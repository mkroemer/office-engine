import { expect, test } from 'bun:test'
import { parasToDocx } from '../src/core/convert'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'
import { useBrowser, viewer, type Params, type Viewer } from './harness'

// A document longer than the browser lets a box scroll (https://github.com/p10node/office-engine/issues/50): what the
// reader or the browser scrolls moves the pages once and as far as it goes, and every page is still reached.

useBrowser()

const PAGES = 200
/** the page that holds lines of text, for the caret */
const LINES_PAGE = 100
const LINES = 60

/** helpers of the page, set by `open` */
interface Long {
  sc(): HTMLElement
  frames(n: number): Promise<void>
  /** the place of the document at the top of the box, counted from the first page, read from the pages as they are drawn */
  place(): number
  /** how far the last page ends above the bottom of the box (24: the end of the document) */
  end(): number | null
}
declare const L: Long

let file: Promise<number[]> | undefined
/** 200 pages, each well over 100,000 px tall (two to three times what a box scrolls over); page 100 holds 60 lines */
function longDocx() {
  return (file ??= (async () => {
    const paras = Array.from({ length: PAGES }, (_, i) => [
      { runs: [{ text: `Page ${i + 1}` }] },
      ...(i + 1 === LINES_PAGE ? Array.from({ length: LINES }, (_, k) => ({ runs: [{ text: `Line ${k + 1}` }] })) : []),
    ]).flat()
    const blob = await parasToDocx(paras)
    const zip = await ZipArchive.open(new BlobSource(blob))
    const xml = (await zip.text('word/document.xml'))
      .replace(/<w:pgSz[^>]*\/>/, '<w:pgSz w:w="12240" w:h="2000000"/>')
      .replaceAll('<w:p><w:r><w:t xml:space="preserve">Page ', '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t xml:space="preserve">Page ')
    return [...new Uint8Array(await (await rebuildZip(zip, new Map([['word/document.xml', xml]]), blob.type)).arrayBuffer())]
  })())
}

async function open(params: Params = {}): Promise<Viewer> {
  const v = await viewer({ onSave: true, ...params })
  await v.page.evaluate((bytes) => window.__setProps({ src: new Uint8Array(bytes), name: 'long.docx' }), await longDocx())
  await v.ready('docx')
  await v.page.waitForFunction((n) => document.querySelector('.de-bar')?.textContent?.includes(`of ${n}`), PAGES)
  await v.page.evaluate((pages) => {
    const sc = () => document.querySelector<HTMLElement>('.de-scroll')!
    const all = () => [...document.querySelectorAll<HTMLElement>('.de-page')]
    ;(window as unknown as { L: Long }).L = {
      sc,
      frames: async (n) => {
        for (let i = 0; i < n; i++) await new Promise((r) => requestAnimationFrame(r))
      },
      place() {
        const box = sc().getBoundingClientRect()
        // the first page that ends below the top of the box (16: the gap between two pages)
        const p = all().find((x) => x.getBoundingClientRect().bottom + 16 > box.top)
        if (!p) return NaN
        const r = p.getBoundingClientRect()
        return (Number(p.dataset.page) - 1) * (r.height + 16) + box.top - r.top
      },
      end() {
        const p = all().find((x) => Number(x.dataset.page) === pages)
        return p ? sc().getBoundingClientRect().bottom - p.getBoundingClientRect().bottom : null
      },
    }
  }, PAGES)
  // the document is longer than its box scrolls over
  expect(await v.page.evaluate(() => L.sc().scrollHeight)).toBeLessThanOrEqual(12_000_000)
  return v
}

/** the thumb dragged to `top`: the place shown there, once its page is drawn */
async function jump(v: Viewer, top: number): Promise<number> {
  await v.page.evaluate((t) => void (L.sc().scrollTop = t), top)
  await v.page.waitForFunction(() => {
    const box = L.sc().getBoundingClientRect()
    return [...document.querySelectorAll('.de-page')].some((p) => p.getBoundingClientRect().top < box.bottom && p.getBoundingClientRect().bottom > box.top)
  })
  await v.page.waitForTimeout(400)
  return v.page.evaluate(() => L.place())
}

/** by the find bar to the page of lines: a place the view itself asks for */
async function findLine20(v: Viewer) {
  await v.btn('Find & replace').click()
  const input = v.page.locator('.de-find input[aria-label="Find"]')
  await input.fill('Line 20')
  await input.press('Enter')
  await v.page.waitForFunction(() => {
    const box = L.sc().getBoundingClientRect()
    const p = [...document.querySelectorAll<HTMLElement>('.de-page [data-pa]')].find((x) => x.textContent === 'Line 20')?.getBoundingClientRect()
    return !!p && p.top >= box.top && p.bottom <= box.bottom
  })
  await input.press('Escape')
}

test('past the cap, the wheel moves the pages as far as it scrolls, and the box at rest is put where its thumb belongs', async () => {
  const v = await open()
  const y0 = await jump(v, 5_000_000)
  // (a jump is not pixel for pixel: the document is more than twice as long as the box's range)
  expect(y0).toBeGreaterThan(9_000_000)
  const box = (await v.page.locator('.de-scroll').boundingBox())!
  await v.page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
  await v.page.mouse.wheel(0, 300)
  await v.page.waitForFunction((y) => Math.abs(L.place() - y - 300) < 1.5, y0)
  // at rest: the box is scrolled back to where its thumb belongs for the place it shows (less far than the step), and nothing moved on screen
  await v.page.waitForFunction(() => L.sc().scrollTop < 5_000_290)
  await v.page.waitForTimeout(300)
  const rest = await v.page.evaluate(() => ({ y: L.place(), top: L.sc().scrollTop }))
  expect(Math.abs(rest.y - y0 - 300)).toBeLessThan(1.5)
  expect(rest.top).toBeGreaterThan(5_000_050)
  expect(rest.top).toBeLessThan(5_000_290)
  // and back, by the wheel again
  await v.page.mouse.wheel(0, -300)
  await v.page.waitForFunction((y) => Math.abs(L.place() - y) < 1.5, y0)
  expect(v.errors).toEqual([])
})

test('past the cap, Page Down moves by one box, not by several', async () => {
  const v = await open()
  const y0 = await jump(v, 5_000_000)
  const h = await v.page.evaluate(() => L.sc().clientHeight)
  await v.page.locator('.de-scroll').focus()
  await v.page.keyboard.press('PageDown')
  await v.page.waitForFunction((y) => L.place() - y > 100, y0)
  await v.page.waitForTimeout(600)
  const moved = (await v.page.evaluate(() => L.place())) - y0
  expect(moved).toBeGreaterThan(h * 0.5)
  expect(moved).toBeLessThanOrEqual(h + 1)
  expect(v.errors).toEqual([])
})

test('past the cap, steps that reach the end of the box before the end of the document go on to the last page', async () => {
  const v = await open()
  // just before the last 1,000,000 px, which scroll pixel for pixel: the document has more left than the box
  const max = await v.page.evaluate(() => L.sc().scrollHeight - L.sc().clientHeight)
  const y0 = await jump(v, max - 1_050_000)
  const out = await v.page.evaluate(async () => {
    const s = L.sc()
    let y = L.place()
    let worst = 0
    let settled = 0
    let steps = 0
    for (; steps < 2000; steps++) {
      const top = s.scrollTop
      s.scrollTop = top + 1900
      const d = s.scrollTop - top
      if (d === 0) break
      await L.frames(1)
      // the box was put back where it has room: only when it could not go on
      if (s.scrollTop < top) settled++
      // (a step past a whole box can end below the pages drawn so far: the next one is drawn a frame later)
      if (Number.isNaN(L.place())) await L.frames(2)
      const now = L.place()
      worst = Math.max(worst, Math.abs(now - y - d))
      y = now
    }
    await L.frames(2)
    return { y, worst, settled, steps, end: L.end(), top: s.scrollTop, max: s.scrollHeight - s.clientHeight }
  })
  // every step moved the document as far as the box, also the ones at which the box was put back
  expect(out.worst).toBeLessThan(1.5)
  expect(out.settled).toBeGreaterThan(0)
  expect(out.settled).toBeLessThan(10)
  // more of the document was scrolled than the box had left
  expect(out.y - y0).toBeGreaterThan(1_100_000)
  // the end of the last page, 24 px above the bottom of the box
  expect(Math.abs(out.end! - 24)).toBeLessThan(1.5)
  expect(Math.abs(out.top - out.max)).toBeLessThan(1.5)
  expect(await v.page.locator(`.de-page[data-page="${PAGES}"]`).textContent()).toContain(`Page ${PAGES}`)
  expect(v.errors).toEqual([])
})

test('past the cap, the caret moved below the box by the arrow keys is followed line by line', async () => {
  const v = await open({ editable: true })
  await findLine20(v)
  // the last line that is whole in the box
  const start = await v.page.evaluate(() => {
    const box = L.sc().getBoundingClientRect()
    const lines = [...document.querySelectorAll<HTMLElement>('.de-page [data-pa]')].filter((x) => /^Line \d+$/.test(x.textContent ?? ''))
    const whole = lines.filter((x) => x.getBoundingClientRect().bottom <= box.bottom)
    return { n: Number(whole.at(-1)!.textContent!.slice(5)), lineH: lines[1].getBoundingClientRect().top - lines[0].getBoundingClientRect().top }
  })
  expect(start.n).toBeGreaterThan(20)
  expect(start.n).toBeLessThan(LINES - 12)
  // the caret at the end of that line, by a click to the right of its text
  const at = await v.page.evaluate((n) => {
    const line = [...document.querySelectorAll<HTMLElement>('.de-page [data-pa]')].find((x) => x.textContent === `Line ${n}`)!
    const r = line.getBoundingClientRect()
    return { x: r.left + 300, y: r.top + r.height / 2 }
  }, start.n)
  await v.page.mouse.click(at.x, at.y)
  await v.page.waitForFunction((t) => (getSelection()?.anchorNode as Node | null)?.parentElement?.closest('[data-pa]')?.textContent === t, `Line ${start.n}`)
  const y0 = await v.page.evaluate(() => L.place())
  let y = y0
  for (let k = 1; k <= 10; k++) {
    await v.page.keyboard.press('ArrowDown')
    await v.page.waitForTimeout(60)
    const now = await v.page.evaluate(() => {
      const box = L.sc().getBoundingClientRect()
      const line = (getSelection()!.anchorNode as Node).parentElement!.closest<HTMLElement>('[data-pa]')!
      // the caret itself (the paragraph also has space after its text)
      const r = getSelection()!.getRangeAt(0).getBoundingClientRect()
      return { y: L.place(), text: line.textContent, h: r.height, above: r.top - box.top, below: box.bottom - r.bottom }
    })
    // the caret is on the next line and in the box: scrolled to by no more than it took
    expect(now.h).toBeGreaterThan(0)
    expect(now.text).toBe(`Line ${start.n + k}`)
    expect(now.above).toBeGreaterThanOrEqual(0)
    expect(now.below).toBeGreaterThanOrEqual(-1)
    expect(now.y - y).toBeGreaterThanOrEqual(0)
    expect(now.y - y).toBeLessThan(start.lineH * 1.6)
    y = now.y
  }
  // (the first line down may still have been in the box)
  expect(y - y0).toBeGreaterThan(start.lineH * 7)
  expect(y - y0).toBeLessThan(start.lineH * 10.5)
  // at rest the box is put where its thumb belongs: the caret and the lines stay where they are
  await v.page.waitForTimeout(500)
  const rest = await v.page.evaluate(() => ({ y: L.place(), text: (getSelection()!.anchorNode as Node).parentElement!.closest('[data-pa]')!.textContent }))
  expect(Math.abs(rest.y - y)).toBeLessThan(1.5)
  expect(rest.text).toBe(`Line ${start.n + 10}`)
  await v.page.keyboard.type('!')
  await v.page.waitForFunction((t) => [...document.querySelectorAll('.de-page [data-pa]')].some((x) => x.textContent === t), `Line ${start.n + 10}!`)
  expect(Math.abs((await v.page.evaluate(() => L.place())) - y)).toBeLessThan(start.lineH * 3)
  expect(v.errors).toEqual([])
})

test('past the cap, a long scroll the browser makes to something on a page that is drawn shows it', async () => {
  const v = await open({ editable: true })
  await findLine20(v)
  await v.page.waitForTimeout(400)
  const y0 = await v.page.evaluate(() => L.place())
  // the first line of the page before, more than 100,000 px up: as a host's own scrollIntoView, or the browser for what takes the focus
  const title = `.de-page[data-page="${LINES_PAGE - 1}"] [data-pa]`
  await v.page.locator(title).first().waitFor({ state: 'attached' })
  await v.page.evaluate((sel) => document.querySelector(sel)!.scrollIntoView({ block: 'center' }), title)
  await v.page.waitForFunction((y) => L.place() < y - 100_000, y0)
  await v.page.waitForTimeout(500)
  const shown = await v.page.evaluate((sel) => {
    const box = L.sc().getBoundingClientRect()
    const r = document.querySelector(sel)!.getBoundingClientRect()
    return { text: document.querySelector(sel)!.textContent, mid: (r.top + r.bottom) / 2 - (box.top + box.bottom) / 2 }
  }, title)
  expect(shown.text).toBe(`Page ${LINES_PAGE - 1}`)
  // in the middle of the box, also once the box has rested
  expect(Math.abs(shown.mid)).toBeLessThan(2)
  expect(v.errors).toEqual([])
})

test('past the cap, another zoom shows the same page', async () => {
  const v = await open()
  const y0 = await jump(v, 6_000_000)
  const page = async () => Number(/(\d+) of 200/.exec((await v.page.locator('.de-bar').textContent()) ?? '')?.[1])
  const n = await page()
  expect(n).toBeGreaterThan(50)
  const h0 = await v.page.evaluate(() => document.querySelector('.de-page')!.getBoundingClientRect().height)
  await v.btn('Zoom out').click()
  await v.page.waitForFunction((h) => document.querySelector('.de-page')!.getBoundingClientRect().height < h - 1, h0)
  await v.page.waitForTimeout(400)
  const h1 = await v.page.evaluate(() => document.querySelector('.de-page')!.getBoundingClientRect().height)
  expect(await page()).toBe(n)
  // as far down the shorter pages as it was down the taller ones
  const y1 = await v.page.evaluate(() => L.place())
  expect(Math.abs(y1 / y0 - h1 / h0)).toBeLessThan(0.002)
  await v.btn('Zoom in').click()
  await v.page.waitForFunction((h) => document.querySelector('.de-page')!.getBoundingClientRect().height > h + 1, h1)
  await v.page.waitForTimeout(400)
  expect(await page()).toBe(n)
  expect(v.errors).toEqual([])
})
