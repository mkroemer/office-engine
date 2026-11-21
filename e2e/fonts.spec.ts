/**
 * The host's fonts in Chromium: a generated TrueType font whose letters are 2 em wide (test/font-builder.ts) and a
 * document set in it. With the font, its text takes about four times the lines, so the page count tells which font
 * the document was laid out with.
 */
import { describe, expect, test } from 'bun:test'
import { parasToDocx, rowsToXlsx } from '../src/core/convert'
import { testFont } from '../test/font-builder'
import { useBrowser, viewer, type Params, type Viewer } from './harness'

useBrowser()

declare global {
  interface Window {
    __frames: { n: number; status: string; font: string }[]
  }
}

const FAMILY = 'Wide Test'
const FONT = Buffer.from(testFont({ family: FAMILY }))
const SENTENCE = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.'

/** a one-word paragraph, then 24 of a few lines, all in `font` */
async function docx(font: string) {
  const paras = [{ runs: [{ text: 'Wide', set: { font } }] }, ...Array.from({ length: 24 }, (_, i) => ({ runs: [{ text: `${i + 1}. ${SENTENCE} ${SENTENCE}`, set: { font } }] }))]
  return Buffer.from(await (await parasToDocx(paras)).arrayBuffer())
}
const DOCS: Partial<Record<string, Promise<Buffer>>> = {
  'wide.docx': docx(FAMILY),
  'calibri.docx': docx('Calibri'),
  // a number in a column 18 characters wide, in Calibri (the workbook's font)
  'pi.xlsx': rowsToXlsx([['3.14159265358979']]).then(async (b) => Buffer.from(await b.arrayBuffer())),
}

const fontsParam = (src: string) => JSON.stringify([{ family: FAMILY, src }])

/**
 * Serves the document and the fonts under /gen/ (a font after `delay` ms, or with `status`), records what is asked
 * for there and, every frame, how many paragraphs page 1 shows, the status text and the state of the test font.
 */
function serve(fonts: Record<string, { delay?: number; status?: number }> = {}) {
  const requested: string[] = []
  const warnings: string[] = []
  /** every font the page fetched, from anywhere */
  const fontRequests: string[] = []
  const setup = async (v: Viewer) => {
    v.page.on('console', (m) => void (m.type() === 'warning' && warnings.push(m.text())))
    v.page.on('request', (r) => void (r.resourceType() === 'font' && fontRequests.push(new URL(r.url()).pathname)))
    await v.ctx.route('**/gen/**', async (route) => {
      const name = new URL(route.request().url()).pathname.split('/').pop()!
      requested.push(name)
      try {
        if (DOCS[name]) return await route.fulfill({ body: await DOCS[name], contentType: 'application/octet-stream' })
        const how = fonts[name] ?? {}
        if (how.delay) await new Promise((r) => setTimeout(r, how.delay))
        if (how.status) return await route.fulfill({ status: how.status, body: 'missing' })
        await route.fulfill({ body: FONT, contentType: 'font/ttf' })
      } catch {
        // the page is gone (the test ended before a slow font was sent)
      }
    })
    await v.page.addInitScript(() => {
      window.__frames = []
      const rec = () => {
        const page1 = document.querySelector('.de-page[data-page="1"]')
        const face = [...document.fonts].find((f) => f.family.replace(/["']/g, '') === 'Wide Test')
        window.__frames.push({ n: page1 ? page1.querySelectorAll('[data-pa]').length : 0, status: document.querySelector('.de-bar .de-status')?.textContent ?? '', font: face?.status ?? 'none' })
        requestAnimationFrame(rec)
      }
      requestAnimationFrame(rec)
    })
  }
  return { requested, warnings, fontRequests, setup }
}

const status = (v: Viewer) => v.page.evaluate(() => document.querySelector('.de-bar .de-status')?.textContent ?? '')

/** the page count once the document is loaded and measured (and stays so for a moment) */
async function pages(v: Viewer, until: (n: number) => boolean = () => true): Promise<number> {
  let last = ''
  for (let i = 0; i < 200; i++) {
    const s = await status(v)
    const m = /^Page \d+ of (\d+) · [\d,]+ words$/.exec(s)
    if (m && s === last && until(Number(m[1]))) return Number(m[1])
    last = s
    await v.page.waitForTimeout(150)
  }
  throw new Error(`not laid out: ${last}`)
}

const open = async (params: Params, fonts?: Record<string, { delay?: number; status?: number }>) => {
  const s = serve(fonts)
  const v = await viewer({ src: 'gen/wide.docx', ...params }, { setup: s.setup })
  return { v, ...s }
}

/** the page count in whatever the browser has for the family: no font declared */
let plain: Promise<number> | undefined
const fallbackPages = () =>
  (plain ??= (async () => {
    const { v } = await open({})
    return pages(v)
  })())

describe('fonts the host supplies', () => {
  test('none declared: nothing is fetched, the text is in a font the browser has', async () => {
    const { v, requested, fontRequests } = await open({})
    const n = await pages(v)
    expect(requested).toEqual(['wide.docx'])
    expect(fontRequests).toEqual([])
    expect(await v.page.evaluate(() => document.fonts.size)).toBe(0)
    expect(n).toBeGreaterThan(0)
    // a font declared for a family the document does not use is not fetched either
    const other = await open({ src: 'samples/sample.docx', fonts: fontsParam('/gen/unused.ttf') })
    await pages(other.v)
    expect(other.requested).toEqual([])
    expect(v.errors).toEqual([])
  })

  test('declared: the document is measured with it, once it is there', async () => {
    const fallback = await fallbackPages()
    const { v, requested, fontRequests } = await open({ fonts: fontsParam('/gen/wide.ttf') }, { 'wide.ttf': { delay: 1200 } })
    const n = await pages(v)
    expect(fontRequests).toEqual(['/gen/wide.ttf'])
    // about four times the lines: far more pages than in the browser's own font
    expect(n).toBeGreaterThan(fallback * 2)
    expect(requested).toEqual(['wide.docx', 'wide.ttf'])

    // the text is in it: the family resolves to the file, a one-word run is 8 em wide and 1 em tall (about 2:1 otherwise)
    expect(await v.page.evaluate((f) => document.fonts.check(`16px "${f}"`), FAMILY)).toBe(true)
    const ratio = await v.page.evaluate(() => {
      const span = [...document.querySelectorAll<HTMLElement>('.de-page [data-pa="0"] span')].find((s) => s.textContent === 'Wide')!
      const r = span.getBoundingClientRect()
      return r.width / r.height
    })
    expect(ratio).toBeGreaterThan(7)
    expect(ratio).toBeLessThan(9)

    // nothing was laid out before it arrived: from the first frame that shows page 1, the font is there and page 1
    // never holds more paragraphs than it ends with (the browser's narrower font would have fitted more first)
    const frames = await v.page.evaluate(() => window.__frames)
    const shown = frames.filter((f) => f.n > 0)
    expect(shown.length).toBeGreaterThan(10)
    expect(shown.filter((f) => f.font !== 'loaded')).toEqual([])
    const final = shown.at(-1)!.n
    expect(shown.filter((f) => f.n > final)).toEqual([])
    const counts = frames.map((f) => /of (\d+)/.exec(f.status)?.[1]).filter((x) => x !== undefined).map(Number)
    expect(counts.every((c, i) => i === 0 || c >= counts[i - 1])).toBe(true)
    expect(v.errors).toEqual([])
  })

  test('as bytes: the same layout, nothing fetched', async () => {
    const fallback = await fallbackPages()
    const { v, requested } = await open({ src: '' })
    await v.page.evaluate(
      ([bytes, family]) => window.__setProps({ src: 'gen/wide.docx', fonts: [{ family: family as string, src: new Uint8Array(bytes as number[]).buffer }] }),
      [[...FONT], FAMILY] as const,
    )
    await v.waitEvent('load')
    const n = await pages(v, (x) => x > fallback * 2)
    expect(n).toBeGreaterThan(fallback * 2)
    expect(requested).toEqual(['wide.docx'])
    expect(v.errors).toEqual([])
  })

  test('declared under the name of a fallback: fetched for the documents in the original (Carlito for Calibri)', async () => {
    const { v, requested } = await open({ src: 'gen/calibri.docx', fonts: JSON.stringify([{ family: 'Carlito', src: '/gen/carlito.ttf' }]) })
    await pages(v)
    expect(requested).toEqual(['calibri.docx', 'carlito.ttf'])
    // (whether the page shows it depends on a Calibri installed here: a font of that name comes first)
    expect(await v.page.evaluate(() => [...document.fonts].find((f) => f.family === 'Carlito')?.status)).toBe('loaded')
    expect(v.errors).toEqual([])
  })

  test('a font that cannot be loaded: reported, and the document is shown in the next font', async () => {
    const fallback = await fallbackPages()
    const { v, warnings } = await open({ fonts: fontsParam('/gen/missing.ttf') }, { 'missing.ttf': { status: 404 } })
    expect(await pages(v)).toBe(fallback)
    const errors = await v.events('error')
    expect(errors).toHaveLength(1)
    expect(errors[0].message).toStartWith(`Font "${FAMILY}" (/gen/missing.ttf) could not be loaded`)
    expect(warnings.filter((w) => w.startsWith(`[office-engine] Font "${FAMILY}"`))).toHaveLength(1)
    expect(v.errors).toEqual([])
  })

  test('a slow font: laid out without it after 3 s, again when it arrives, the text in view kept in view', async () => {
    const fallback = await fallbackPages()
    expect(fallback).toBeGreaterThan(1)
    const { v, warnings } = await open({ fonts: fontsParam('/gen/slow.ttf') }, { 'slow.ttf': { delay: 7000 } })
    expect(await pages(v)).toBe(fallback)
    // shown while the font was still on its way
    const frames = await v.page.evaluate(() => window.__frames)
    expect(frames.find((f) => f.n > 0)?.font).toBe('loading')
    expect(warnings.some((w) => /^\[office-engine\] font "Wide Test" \(\/gen\/slow\.ttf\) is still loading after 3 s/.test(w))).toBe(true)

    // the reader is on page 2: its first paragraph near the top of the view
    await v.page.evaluate(() => {
      const sc = document.querySelector<HTMLElement>('.de-scroll')!
      const page1 = document.querySelector<HTMLElement>('.de-page[data-page="1"]')!
      sc.scrollTop = page1.offsetTop + page1.offsetHeight + 16 + 30
    })
    await v.page.waitForTimeout(300)
    const before = await v.page.evaluate(() => {
      const sc = document.querySelector<HTMLElement>('.de-scroll')!
      const el = document.querySelector<HTMLElement>('.de-page[data-page="2"] [data-pa]')!
      return { pa: el.dataset.pa!, y: el.getBoundingClientRect().top - sc.getBoundingClientRect().top, scroll: sc.scrollTop }
    })
    expect(before.y).toBeGreaterThan(0)

    // the font arrives: measured again, many more pages, and that paragraph where it was on screen
    const n = await pages(v, (x) => x > fallback * 2)
    expect(n).toBeGreaterThan(fallback * 2)
    const after = await v.page.evaluate((pa) => {
      const sc = document.querySelector<HTMLElement>('.de-scroll')!
      const el = document.querySelector<HTMLElement>(`[data-pa="${pa}"]`)
      return el ? { y: el.getBoundingClientRect().top - sc.getBoundingClientRect().top, scroll: sc.scrollTop } : null
    }, before.pa)
    expect(after).not.toBeNull()
    expect(Math.abs(after!.y - before.y)).toBeLessThan(3)
    expect(after!.scroll).toBeGreaterThan(before.scroll)
    expect(v.errors).toEqual([])
  })

  test('editing: the selection is kept when a late font moves its paragraph to another page', async () => {
    const fallback = await fallbackPages()
    const { v } = await open({ editable: true, fonts: fontsParam('/gen/late.ttf') }, { 'late.ttf': { delay: 5000 } })
    await v.page.locator('.de-ebar').first().waitFor()
    expect(await pages(v)).toBe(fallback)
    // paragraph 6: on page 1 in the browser's font, further on in the test font
    const pageOf = () => v.page.evaluate(() => document.querySelector('[data-pa="6"]')?.closest<HTMLElement>('.de-page')?.dataset.page)
    expect(await pageOf()).toBe('1')
    await v.selectText('6. Lorem')
    await pages(v, (x) => x > fallback * 2)
    expect(await pageOf()).not.toBe('1')
    expect(await v.page.evaluate(() => getSelection()?.toString())).toBe('6. Lorem')
    await v.page.keyboard.type('Six:')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('[data-pa="6"]').innerText()).toStartWith('Six: ipsum dolor')
    expect(v.errors).toEqual([])
  })

  test("a font of the page's own (@font-face) that arrives late: the document is measured again", async () => {
    const fallback = await fallbackPages()
    const s = serve({ 'css.ttf': { delay: 1500 } })
    const v = await viewer(
      { src: 'gen/wide.docx' },
      {
        setup: async (v) => {
          await s.setup(v)
          await v.page.addInitScript(() => {
            const style = document.createElement('style')
            style.textContent = '@font-face { font-family: "Wide Test"; src: url(/gen/css.ttf) format("truetype"); }'
            document.addEventListener('DOMContentLoaded', () => document.head.append(style))
          })
        },
      },
    )
    const n = await pages(v, (x) => x > fallback * 2)
    expect(n).toBeGreaterThan(fallback * 2)
    expect(s.requested).toEqual(['wide.docx', 'css.ttf'])
    expect(v.errors).toEqual([])
  })

  /** a number shown in a font 2 em a digit: what fits of it in its column is a few characters */
  const fitted = async (v: Viewer) => {
    for (let i = 0; i < 100; i++) {
      const t = await v.shown('A1')
      if (t.length <= 5) return t
      await v.page.waitForTimeout(150)
    }
    return v.shown('A1')
  }

  test('xlsx: a number is fitted to its column again when the declared font arrives', async () => {
    const { v, requested } = await open({ src: 'gen/pi.xlsx', fonts: JSON.stringify([{ family: 'Calibri', src: '/gen/calibri.ttf' }]) }, { 'calibri.ttf': { delay: 800 } })
    expect(await fitted(v)).toMatch(/^3\.\d{0,3}$/)
    expect(requested).toEqual(['pi.xlsx', 'calibri.ttf'])
    expect(v.errors).toEqual([])
  })

  test("xlsx: ...and when a font of the page's own arrives (the widths measured before were the fallback's)", async () => {
    const s = serve({ 'css-calibri.ttf': { delay: 800 } })
    const v = await viewer(
      { src: 'gen/pi.xlsx' },
      {
        setup: async (v) => {
          await s.setup(v)
          await v.page.addInitScript(() => {
            const style = document.createElement('style')
            style.textContent = '@font-face { font-family: Calibri; src: url(/gen/css-calibri.ttf) format("truetype"); }'
            document.addEventListener('DOMContentLoaded', () => document.head.append(style))
          })
        },
      },
    )
    expect(await fitted(v)).toMatch(/^3\.\d{0,3}$/)
    expect(s.requested).toEqual(['pi.xlsx', 'css-calibri.ttf'])
    expect(v.errors).toEqual([])
  })
})
