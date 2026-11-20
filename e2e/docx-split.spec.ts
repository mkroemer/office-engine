import { describe, expect, test } from 'bun:test'
import type { Paragraph, Table } from '../src/core/docx/types'
import { DocBuilder } from '../test/doc-builder'
import { bareDocx, para, table } from '../test/docx-builder'
import { docxBlocks, paraText, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

/** opens a built document in the fixture, editable, at `zoom`; waits until every page is laid out */
async function open(doc: Blob | Promise<Blob>, zoom = 1, name = 'split.docx') {
  const v = await viewer({ editable: true, onSave: true }, { wait: false })
  const b64 = Buffer.from(await (await doc).arrayBuffer()).toString('base64')
  await v.page.evaluate(([b, z, n]) => window.__setProps({ src: new Blob([Uint8Array.from(atob(b as string), (c) => c.charCodeAt(0))]), name: n as string, zoom: z as number }), [b64, zoom, name] as const)
  await v.ready('docx')
  await v.page.waitForFunction(() => /Page \d+ of \d+ ·/.test(document.querySelector('.de-bar .de-status')?.textContent ?? ''))
  await v.page.waitForTimeout(200)
  return v
}

const pageCount = async (v: Viewer) => Number(/Page \d+ of (\d+)/.exec(await v.status())![1])

/** scrolls page `n` (1-based) to the top of the view */
async function showPage(v: Viewer, n: number) {
  await v.page.evaluate((k) => {
    const sc = document.querySelector<HTMLElement>('.de-scroll')!
    const el = document.querySelector<HTMLElement>(`.de-page[data-page="${k}"]`)
    if (el) sc.scrollTop = el.offsetTop - 16
    else sc.scrollTop = ((k - 1) / Number(/of (\d+)/.exec(document.querySelector('.de-bar .de-status')!.textContent!)![1])) * sc.scrollHeight
  }, n)
  await v.page.locator(`.de-page[data-page="${n}"]`).waitFor()
  await v.page.waitForTimeout(150)
}

/** what each page shows: the text of its paragraphs (by address), the text of its table rows (copies apart) */
async function pageContent(v: Viewer, n: number) {
  await showPage(v, n)
  return v.page.locator(`.de-page[data-page="${n}"]`).evaluate((pg) => {
    const rows = Array.from(pg.querySelectorAll('tr')).map((tr) => ({ copy: tr.classList.contains('de-head-copy'), text: Array.from(tr.cells).map((c) => c.innerText.trim()).join('|') }))
    const paras = Array.from(pg.querySelectorAll<HTMLElement>('.de-page-inner > div:not(.de-hf):not(.de-hf-zone) [data-pa]:not([data-pa*="."])')).map((p) => ({ pa: p.dataset.pa!, text: p.textContent ?? '' }))
    const inner = pg.querySelector<HTMLElement>('.de-page-inner')!
    const box = inner.getBoundingClientRect()
    return { rows, paras, top: pg.getBoundingClientRect().top, height: pg.getBoundingClientRect().height, zoom: box.height / inner.offsetHeight }
  })
}

const LOREM = 'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua'.split(' ')
const words = (n: number, seed = 0) => Array.from({ length: n }, (_, i) => LOREM[(i * 7 + seed) % LOREM.length]).join(' ')
/** `n` one-line paragraphs */
const filler = (n: number) => Array.from({ length: n }, (_, i) => para(`Filler line ${i + 1}`)).join('')

interface Frag {
  page: number
  text: string
  /** px from the top / bottom of the page's text area, at zoom 1 */
  top: number
  bottom: number
  /** offsets in the element's text where its lines start (after the first) */
  lines: number[]
  /** last line: right edge of its last non-blank character, from the text area's right edge */
  right: number
}

/** the parts of top-level paragraph `pa` on every page, in order */
async function fragments(v: Viewer, pa: number): Promise<Frag[]> {
  const out: Frag[] = []
  for (let k = 1; k <= (await pageCount(v)); k++) {
    await showPage(v, k)
    const got = await v.page.locator(`.de-page[data-page="${k}"]`).evaluate((pg, a) => {
      const inner = pg.querySelector<HTMLElement>('.de-page-inner')!
      const z = inner.getBoundingClientRect().height / inner.offsetHeight
      const area = inner.querySelector<HTMLElement>(':scope > div[style*="position: absolute"]:not(.de-hf):not(.de-hf-zone)')!.getBoundingClientRect()
      const pageBottom = pg.getBoundingClientRect().top + (792 - 72) * (4 / 3) * z
      return Array.from(pg.querySelectorAll<HTMLElement>(`[data-pa="${a}"]`)).map((el) => {
        const r = el.getBoundingClientRect()
        const lines: number[] = []
        const range = document.createRange()
        let o = 0
        let lastTop = -1
        let right = 0
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
        for (let t = walker.nextNode() as Text | null; t; t = walker.nextNode() as Text | null) {
          for (let i = 0; i < t.length; i++, o++) {
            range.setStart(t, i)
            range.setEnd(t, i + 1)
            const cr = range.getBoundingClientRect()
            if (lastTop >= 0 && cr.top > lastTop + 2) lines.push(o)
            if (lastTop < 0 || cr.top > lastTop + 2) lastTop = cr.top
            if (cr.top >= lastTop - 2 && /\S/.test(t.data[i])) right = cr.right
          }
        }
        return { text: el.textContent ?? '', top: (r.top - area.top) / z, bottom: (pageBottom - r.bottom) / z, lines, right: (area.right - right) / z }
      })
    }, pa)
    out.push(...got.map((g) => ({ ...g, page: k })))
  }
  return out
}

/** where the lines of `text` start when laid out whole like the paragraph's first part (its look and width) */
async function wholeLines(v: Viewer, pa: number, text: string) {
  await showPage(v, (await fragments(v, pa))[0].page)
  return v.page.evaluate(
    ([a, t]) => {
      const f = document.querySelector<HTMLElement>(`[data-pa="${a}"]`)!
      const el = f.cloneNode(false) as HTMLElement
      el.removeAttribute('data-pa')
      el.classList.remove('de-just')
      el.style.height = ''
      el.style.paddingBottom = getComputedStyle(f).paddingTop
      const sp = document.createElement('span')
      sp.textContent = t as string
      el.append(sp)
      f.parentElement!.append(el)
      const lines: number[] = []
      const range = document.createRange()
      const node = sp.firstChild as Text
      let lastTop = -1
      for (let i = 0; i < node.length; i++) {
        range.setStart(node, i)
        range.setEnd(node, i + 1)
        const top = range.getBoundingClientRect().top
        if (lastTop >= 0 && top > lastTop + 2) lines.push(i)
        if (lastTop < 0 || top > lastTop + 2) lastTop = top
      }
      el.remove()
      return lines
    },
    [pa, text] as const,
  )
}

describe('a paragraph longer than the rest of a page', () => {
  test('goes on at the top of the next page: no gap, no text lost or shown twice, its lines as when whole', async () => {
    const long = words(420)
    // justified, a first-line indent, space before and after
    const pPr = '<w:jc w:val="both"/><w:ind w:firstLine="720"/><w:spacing w:before="240" w:after="240"/>'
    const blob = await bareDocx(filler(30) + para(long, pPr) + para('After'))
    const model = (await docxBlocks(blob)).filter((b): b is Paragraph => b.t === 'p').map(paraText)
    expect(model[30]).toBe(long)
    const v = await open(blob)
    const fr = await fragments(v, 30)
    expect(fr.map((f) => f.page)).toEqual([1, 2])
    expect(fr.map((f) => f.text).join('')).toBe(long)
    // the first part ends less than a line above the bottom of the text area, the second starts at its top
    expect(fr[0].bottom).toBeGreaterThanOrEqual(-1)
    expect(fr[0].bottom).toBeLessThan(18)
    expect(Math.abs(fr[1].top)).toBeLessThan(1)
    // cut at the start of a line: the space it broke at ends the first part
    expect(fr[0].text.endsWith(' ')).toBe(true)
    expect(fr[1].text.startsWith(' ')).toBe(false)
    // justified: the last line of the first part reaches the right edge, the paragraph's last line does not
    expect(Math.abs(fr[0].right)).toBeLessThan(1.5)
    expect(fr[1].right).toBeGreaterThan(20)
    // the same lines as the paragraph laid out whole
    const whole = await wholeLines(v, 30, long)
    const n0 = fr[0].text.length
    expect([...fr[0].lines, n0, ...fr[1].lines.map((x) => x + n0)]).toEqual(whole)
    expect(v.errors).toEqual([])
  })

  test('line breaks inside it and a hanging indent: the same lines on both pages', async () => {
    const lines = Array.from({ length: 30 }, (_, i) => `${i + 1}. ${words(14 + (i % 3) * 9, i)}`)
    const text = lines.join('\n')
    const blob = await bareDocx(filler(28) + para(text, '<w:ind w:left="720" w:hanging="360"/>'))
    const v = await open(blob)
    const fr = await fragments(v, 28)
    expect(fr.length).toBe(2)
    expect(fr.map((f) => f.text).join('')).toBe(text)
    // not cut inside a line: the first part ends with a line break or a space, the second starts with neither
    expect(/[\n ]$/.test(fr[0].text)).toBe(true)
    expect(/^[\n ]/.test(fr[1].text)).toBe(false)
    const whole = await wholeLines(v, 28, text)
    const n0 = fr[0].text.length
    expect([...fr[0].lines, n0, ...fr[1].lines.map((x) => x + n0)]).toEqual(whole)
    expect(v.errors).toEqual([])
  })
})

/** words that occur once each: `n1q n2q …` */
const uniq = (n: number) => Array.from({ length: n }, (_, i) => `n${i + 1}q`).join(' ')

/** the parts of paragraph `pa` rendered now (the pages around the view), in order */
const parts = (v: Viewer, pa: number) =>
  v.page.evaluate(
    (a) =>
      Array.from(document.querySelectorAll<HTMLElement>(`.de-pages [data-pa="${a}"]`)).map((el) => ({
        text: el.textContent ?? '',
        page: Number(el.closest<HTMLElement>('.de-page')!.dataset.page),
        from: Number(el.dataset.from ?? 0),
      })),
    pa,
  )

/** sets the DOM selection in parts of paragraph `pa`: [part, offset in its text ('end': its end)] for anchor and focus */
async function setSel(v: Viewer, pa: number, a: [number, number | 'end'], f = a) {
  await v.page.evaluate(
    ([p, x, y]) => {
      const els = Array.from(document.querySelectorAll<HTMLElement>(`.de-pages [data-pa="${p}"]`))
      const at = ([k, o]: readonly [number, number | 'end']): [Node, number] => {
        const texts: Text[] = []
        const w = document.createTreeWalker(els[k], NodeFilter.SHOW_TEXT)
        for (let t = w.nextNode(); t; t = w.nextNode()) texts.push(t as Text)
        let rem = o === 'end' ? texts.reduce((n, t) => n + t.length, 0) : o
        for (const t of texts) {
          if (rem <= t.length) return [t, rem]
          rem -= t.length
        }
        return [texts[texts.length - 1], texts[texts.length - 1].length]
      }
      document.querySelector<HTMLElement>('.de-pages')!.focus()
      const [an, ao] = at(x)
      const [fn, fo] = at(y)
      getSelection()!.setBaseAndExtent(an, ao, fn, fo)
    },
    [pa, a, f] as const,
  )
  await v.page.waitForTimeout(150)
}

/** where the caret is: its page, and its offset in paragraph `pa` (-1: not in it) */
const caretAt = (v: Viewer, pa: number) =>
  v.page.evaluate((a) => {
    const s = getSelection()!
    const n = s.focusNode!
    const el = (n.nodeType === 1 ? (n as Element) : n.parentElement)!.closest<HTMLElement>('[data-pa]')
    const page = Number(el?.closest<HTMLElement>('.de-page')?.dataset.page ?? -1)
    if (!el || el.dataset.pa !== String(a)) return { page, o: -1 }
    const r = document.createRange()
    r.setStart(el, 0)
    r.setEnd(n, s.focusOffset)
    return { page, o: Number(el.dataset.from ?? 0) + r.toString().length }
  }, pa)

const joined = async (v: Viewer, pa: number) => (await parts(v, pa)).map((p) => p.text).join('')
const savedText = async (v: Viewer, k: number) => paraText((await docxBlocks((await v.save()).bytes)).filter((b): b is Paragraph => b.t === 'p')[k])

describe('editing a paragraph split across pages', () => {
  const long = uniq(420)
  const doc = () => bareDocx(filler(30) + para(long) + para('After'))

  test('typing in its second part, and at the end of its first until the page cut moves: the caret stays with the typing', async () => {
    const v = await open(doc(), 0.5)
    let ps = await parts(v, 30)
    expect(ps.map((p) => p.page)).toEqual([1, 2])
    const cut = ps[1].from
    // after the first word of the second part
    const w = ps[1].text.split(' ')[0]
    await setSel(v, 30, [1, w.length])
    await v.page.keyboard.type('X')
    let text = long.slice(0, cut + w.length) + 'X' + long.slice(cut + w.length)
    expect(await joined(v, 30)).toBe(text)
    expect(await caretAt(v, 30)).toEqual({ page: 2, o: cut + w.length + 1 })

    // a click right of the last line of the first part: the caret at its end, on page 1
    const end = await v.page.evaluate(() => {
      const el = document.querySelector<HTMLElement>('.de-pages [data-pa="30"]')!
      const r = document.createRange()
      r.selectNodeContents(el)
      const rects = r.getClientRects()
      const last = rects[rects.length - 1]
      return { x: el.getBoundingClientRect().right - 2, y: last.top + last.height / 2 }
    })
    await v.page.mouse.click(end.x, end.y)
    await v.page.waitForTimeout(150)
    expect(await caretAt(v, 30)).toEqual({ page: 1, o: cut })
    // typing there: text that does not fit the last line of the page starts the next page, the caret with it
    await v.page.keyboard.type('abc ')
    text = text.slice(0, cut) + 'abc ' + text.slice(cut)
    ps = await parts(v, 30)
    expect(ps.map((p) => p.text).join('')).toBe(text)
    expect(await caretAt(v, 30)).toEqual({ page: ps[1].from <= cut ? 2 : 1, o: cut + 4 })

    // typing inside the last line of the page: the words after the caret go on to the next page, then the word typed
    // does not fit either and goes there too, with the caret; the page cut moves back as the line fills up
    ps = await parts(v, 30)
    const lastSpace = ps[0].text.trimEnd().lastIndexOf(' ')
    let o = lastSpace
    await setSel(v, 30, [0, o])
    const cuts = new Set<number>()
    for (let k = 0; k < 40; k++) {
      await v.page.keyboard.type('x')
      text = text.slice(0, o) + 'x' + text.slice(o)
      o++
      ps = await parts(v, 30)
      expect(ps.map((p) => p.text).join('')).toBe(text)
      const c = await caretAt(v, 30)
      expect(c.o).toBe(o)
      // the caret is on the page of the part that holds it
      expect(c.page).toBe(o > ps[1].from ? 2 : 1)
      cuts.add(ps[1].from)
    }
    expect(cuts.size).toBeGreaterThan(1)
    expect((await caretAt(v, 30)).page).toBe(2)
    // the caret stayed where the typing is
    await v.page.keyboard.type('Y')
    text = text.slice(0, o) + 'Y' + text.slice(o)
    expect(await joined(v, 30)).toBe(text)
    expect(await savedText(v, 30)).toBe(text)
    expect(v.errors).toEqual([])
  })

  test('Backspace at the end of the first part joins two words: the joined word goes to the next page, the caret too', async () => {
    const v = await open(doc(), 0.5)
    const ps = await parts(v, 30)
    const cut = ps[1].from
    await setSel(v, 30, [0, 'end'])
    expect(await caretAt(v, 30)).toEqual({ page: 1, o: cut })
    await v.page.keyboard.press('Backspace')
    await v.page.waitForTimeout(150)
    const after = await parts(v, 30)
    expect(after.map((p) => p.text).join('')).toBe(long.slice(0, cut - 1) + long.slice(cut))
    // the word before the cut and the one after it are one word now, too long for the end of the line
    expect(after[1].from).toBeLessThan(cut - 1)
    expect(await caretAt(v, 30)).toEqual({ page: 2, o: cut - 1 })
    expect(v.errors).toEqual([])
  })

  test('a selection over the page cut is deleted, and undone', async () => {
    const v = await open(doc(), 0.5)
    const ps = await parts(v, 30)
    const cut = ps[1].from
    const back = ps[0].text.trimEnd().split(' ').slice(-3).join(' ').length + 1
    await setSel(v, 30, [0, ps[0].text.length - back], [1, 7])
    await v.page.keyboard.press('Backspace')
    await v.page.waitForTimeout(200)
    const gone = long.slice(0, cut - back) + long.slice(cut + 7)
    expect(await joined(v, 30)).toBe(gone)
    expect((await caretAt(v, 30)).o).toBe(cut - back)
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(200)
    expect(await joined(v, 30)).toBe(long)
    expect(await savedText(v, 30)).toBe(long)
    expect(v.errors).toEqual([])
  })

  test('arrow keys over the page cut move one character, the caret on the page of its text', async () => {
    const v = await open(doc(), 0.5)
    let ps = await parts(v, 30)
    const cut = ps[1].from
    // ← at the start of the second part: before the last character of the first (the space the line broke at)
    await setSel(v, 30, [1, 0])
    expect(await caretAt(v, 30)).toEqual({ page: 2, o: cut })
    await v.page.keyboard.press('ArrowLeft')
    expect(await caretAt(v, 30)).toEqual({ page: 1, o: cut - 1 })
    await v.page.keyboard.press('ArrowRight')
    expect(await caretAt(v, 30)).toEqual({ page: 1, o: cut })
    // → at the end of the first part: after the first character of the second
    await v.page.keyboard.press('ArrowRight')
    expect(await caretAt(v, 30)).toEqual({ page: 2, o: cut + 1 })
    await v.page.keyboard.type('R')
    let text = long.slice(0, cut + 1) + 'R' + long.slice(cut + 1)
    expect(await joined(v, 30)).toBe(text)
    // Shift+← from there: the selection grows back over the cut, one character a key; typing replaces it
    for (let k = 0; k < 3; k++) await v.page.keyboard.press('Shift+ArrowLeft')
    await v.page.keyboard.type('K')
    text = text.slice(0, cut - 1) + 'K' + text.slice(cut + 2)
    expect(await joined(v, 30)).toBe(text)
    // ↓ from the last line of the first part: the first line of the second; ↑ back
    ps = await parts(v, 30)
    await setSel(v, 30, [0, ps[0].text.length - 3])
    await v.page.keyboard.press('ArrowDown')
    const down = await caretAt(v, 30)
    expect(down.page).toBe(2)
    expect(down.o).toBeGreaterThanOrEqual(ps[1].from)
    await v.page.keyboard.press('ArrowUp')
    expect((await caretAt(v, 30)).page).toBe(1)
    text = await joined(v, 30)
    expect(await savedText(v, 30)).toBe(text)
    expect(v.errors).toEqual([])
  })

  test('find: a phrase over the page cut is found once, highlighted on both pages; an input method composes at the cut', async () => {
    const v = await open(doc(), 0.5)
    const ps = await parts(v, 30)
    const cut = ps[1].from
    const phrase = `${ps[0].text.trim().split(' ').pop()} ${ps[1].text.split(' ')[0]}`
    await setSel(v, 30, [0, 0])
    await v.page.keyboard.press('Control+f')
    await v.page.locator('.de-find input').first().fill(phrase)
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-find-count').innerText()).toBe('1 / 1')
    const hl = await v.page.evaluate(() => {
      const reg = (CSS as unknown as { highlights: Map<string, Set<Range>> }).highlights
      const h = reg.get('de-find-cur') ?? reg.get('de-find')
      return h ? [...h].map((r) => ({ text: r.toString(), page: (r.startContainer.parentElement!.closest('.de-page') as HTMLElement).dataset.page })) : []
    })
    expect(hl.map((r) => r.page)).toEqual(['1', '2'])
    expect(hl.map((r) => r.text).join('')).toBe(phrase)
    await v.page.keyboard.press('Escape')

    // an input method composing at the end of the first part
    await setSel(v, 30, [0, 'end'])
    const cdp = await v.ctx.newCDPSession(v.page)
    await cdp.send('Input.imeSetComposition', { text: 'd', selectionStart: 1, selectionEnd: 1 })
    await cdp.send('Input.imeSetComposition', { text: 'đ', selectionStart: 1, selectionEnd: 1 })
    await cdp.send('Input.insertText', { text: 'đ' })
    await v.page.waitForTimeout(200)
    await v.page.keyboard.type('x')
    await v.page.waitForTimeout(200)
    const text = long.slice(0, cut) + 'đx' + long.slice(cut)
    expect(await joined(v, 30)).toBe(text)
    expect(await savedText(v, 30)).toBe(text)
    expect(v.errors).toEqual([])
  })
})

describe('more layouts', () => {
  test('text columns: a paragraph goes on from one column to the next, then to the next page', async () => {
    const long = uniq(900)
    const sectPr = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="720"/>'
    const v = await open(bareDocx(filler(20) + para(long) + para('After'), { sectPr }), 0.5)
    const ps = await parts(v, 20)
    expect(ps.length).toBeGreaterThanOrEqual(3)
    expect(ps.map((p) => p.text).join('')).toBe(long)
    // the first two parts on page 1, in its two columns side by side
    const cols = await v.page.evaluate(() =>
      Array.from(document.querySelectorAll<HTMLElement>('.de-pages [data-pa="20"]')).map((el) => ({ page: el.closest<HTMLElement>('.de-page')!.dataset.page, col: Array.from(el.closest('.de-page-cols')!.children).indexOf(el.closest('.de-page-col')!) })),
    )
    expect(cols.slice(0, 3)).toEqual([
      { page: '1', col: 0 },
      { page: '1', col: 1 },
      { page: '2', col: 0 },
    ])
    expect(v.errors).toEqual([])
  })

  test('a justified paragraph: ↓ from the last line of the first part goes to the next page', async () => {
    const long = uniq(420)
    const v = await open(bareDocx(filler(30) + para(long, '<w:jc w:val="both"/>') + para('After')), 0.5)
    const ps = await parts(v, 30)
    expect(ps.map((p) => p.page)).toEqual([1, 2])
    await setSel(v, 30, [0, ps[0].text.length - 3])
    await v.page.keyboard.press('ArrowDown')
    const c = await caretAt(v, 30)
    expect(c.page).toBe(2)
    expect(c.o).toBeGreaterThanOrEqual(ps[1].from)
    // End on the last line of the first part: the caret at its end, still on page 1
    await setSel(v, 30, [0, ps[0].text.length - 3])
    await v.page.keyboard.press('End')
    expect(await caretAt(v, 30)).toEqual({ page: 1, o: ps[1].from })
    expect(v.errors).toEqual([])
  })
})

describe('a Word 97 file', () => {
  test('converted, it is cut the same way: a long paragraph, a long table with its header row repeated', async () => {
    const d = new DocBuilder()
    for (let i = 0; i < 30; i++) d.p(`Filler line ${i + 1}`)
    const long = uniq(420)
    d.p(long)
    d.table({ widths: [4000, 4000], header: true, borders: true, rows: [[{ paras: ['Head A'] }, { paras: ['Head B'] }], ...Array.from({ length: 70 }, (_, i) => [{ paras: [`row ${i + 1} a`] }, { paras: [`row ${i + 1} b`] }])] })
    d.p('The end')
    const v = await open(new Blob([d.build()]), 0.5, 'split.doc')
    const ps = await parts(v, 30)
    expect(ps.length).toBe(2)
    expect(ps.map((p) => p.text).join('')).toBe(long)
    const n = await pageCount(v)
    let copies = 0
    for (let k = 2; k <= n; k++) {
      await showPage(v, k)
      copies += await v.page.locator(`.de-page[data-page="${k}"] tr.de-head-copy`).count()
    }
    expect(copies).toBeGreaterThan(0)
    expect(v.errors).toEqual([])
  })
})

describe('comments', () => {
  test('a comment on text over the page cut marks it on both pages', async () => {
    const long = uniq(420)
    const v = await open(bareDocx(filler(30) + para(long) + para('After')), 0.5)
    const ps = await parts(v, 30)
    const cut = ps[1].from
    const back = ps[0].text.trimEnd().split(' ').slice(-2).join(' ').length + 1
    await setSel(v, 30, [0, ps[0].text.length - back], [1, 11])
    await v.page.keyboard.press('Control+Alt+KeyM')
    await v.page.locator('.de-cm-panel').waitFor()
    const quote = long.slice(cut - back, cut + 11)
    expect(await v.page.locator('.de-cm-draft .de-cm-quote').innerText()).toBe(quote.trim())
    await v.page.keyboard.type('over the page end')
    await v.page.keyboard.press('Enter')
    const add = await v.waitEvent('comment')
    await v.page.waitForTimeout(200)
    const marked = await v.page.evaluate(
      (id) => Array.from(document.querySelectorAll<HTMLElement>(`.de-pages [data-cm~="${id}"]`)).map((el) => ({ page: el.closest<HTMLElement>('.de-page')!.dataset.page, text: el.textContent })),
      add.id as string,
    )
    expect(marked.map((m) => m.page)).toEqual(['1', '2'])
    expect(marked.map((m) => m.text).join('')).toBe(quote)
    expect(v.errors).toEqual([])
  })
})

describe('print', () => {
  test('the printed pages are the pages shown: a paragraph and a table split the same way, nothing taller than the paper', async () => {
    const long = words(420)
    const rows = [['Head A', 'Head B'], ...Array.from({ length: 70 }, (_, i) => [`row ${i + 1} a`, `row ${i + 1} b`])]
    const v = await open(bareDocx(filler(30) + para(long) + table(rows, { header: 1 }) + para('The end')))
    const shown = await pageCount(v)
    await v.page.evaluate(() => {
      ;(window as any).__printed = null
      new MutationObserver((ms) => {
        for (const m of ms)
          for (const n of m.addedNodes) {
            if ((n as HTMLElement).tagName !== 'IFRAME') continue
            const w = (n as HTMLIFrameElement).contentWindow!
            w.print = () => {
              const d = w.document
              ;(window as any).__printed = {
                pages: d.querySelectorAll('.de-print-page').length,
                tall: d.querySelectorAll('.de-print-tall').length,
                parts: [...d.querySelectorAll<HTMLElement>('[data-pa="30"]')].map((el) => el.textContent),
                copies: d.querySelectorAll('tr.de-head-copy').length,
                tablePages: [...d.querySelectorAll('.de-print-page')].filter((pg) => pg.querySelector('tr')).length,
                rows: [...d.querySelectorAll('tr:not(.de-head-copy)')].map((tr) => (tr as HTMLElement).innerText.replace(/\s+/g, ' ').trim()),
              }
            }
            w.document.execCommand = () => false
          }
      }).observe(document.body, { childList: true })
    })
    await v.page.keyboard.press('Control+p')
    await v.page.waitForFunction(() => (window as any).__printed, null, { timeout: 30_000 })
    const p = await v.page.evaluate(() => (window as any).__printed)
    expect(p.pages).toBe(shown)
    expect(p.tall).toBe(0)
    expect(p.parts.length).toBe(2)
    expect(p.parts.join('')).toBe(long)
    expect(p.tablePages).toBeGreaterThan(1)
    expect(p.copies).toBe(p.tablePages - 1)
    expect(p.rows).toEqual(rows.map((r) => r.join(' ')))
    expect(v.errors).toEqual([])
  })
})

describe('keep rules', () => {
  // every line 12 pt (16 px): the text area of a letter page holds 54 of them
  const EXACT = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>'
  const nLines = (n: number, pPr = '') => para(Array.from({ length: n }, (_, i) => `line ${i + 1}`).join('\n'), pPr)
  /** page and line count of each part of paragraph `pa` */
  const layout = async (body: string, pa: number, pPrDefault = EXACT) => {
    const v = await open(bareDocx(body, { pPrDefault }), 0.5)
    const ps = await parts(v, pa)
    expect(v.errors).toEqual([])
    return ps.map((p) => [p.page, p.text.replace(/\n$/, '').split('\n').length])
  }

  test('widow / orphan control: no first or last line alone on a page; off when the document says so', async () => {
    // room for 1 line: the paragraph starts the next page
    expect(await layout(filler(53) + nLines(6), 53)).toEqual([[2, 6]])
    // room for 5 of 6: 4 stay, 2 go on
    expect(await layout(filler(49) + nLines(6), 49)).toEqual([
      [1, 4],
      [2, 2],
    ])
    const off = EXACT + '<w:widowControl w:val="0"/>'
    expect(await layout(filler(53) + nLines(6), 53, off)).toEqual([
      [1, 1],
      [2, 5],
    ])
    expect(await layout(filler(49) + nLines(6), 49, off)).toEqual([
      [1, 5],
      [2, 1],
    ])
  })

  test('keep lines together: the paragraph starts the next page', async () => {
    expect(await layout(filler(50) + nLines(6, '<w:keepLines/>'), 50)).toEqual([[2, 6]])
    expect(await layout(filler(50) + nLines(6), 50)).toEqual([
      [1, 4],
      [2, 2],
    ])
  })

  test('keep with next: a heading stays with the first lines of its paragraph, a caption with its table', async () => {
    const heading = (t: string) => para(t, '<w:keepNext/>')
    // room for the heading and 2 lines: they stay
    expect(await layout(filler(51) + heading('Heading') + nLines(6), 51)).toEqual([[1, 1]])
    // room for the heading and 1 line: the heading goes with the paragraph to the next page
    expect(await layout(filler(52) + heading('Heading') + nLines(6), 52)).toEqual([[2, 1]])
    // two headings in a row: both go
    expect(await layout(filler(51) + heading('Part') + heading('Heading') + nLines(6), 51)).toEqual([[2, 1]])
    // a caption kept with the table below it
    const rows = Array.from({ length: 8 }, (_, i) => [`r${i + 1}`, 'x'])
    expect(await layout(filler(52) + heading('Caption') + table(rows, { header: 1 }), 52)).toEqual([[2, 1]])
    expect(await layout(filler(52) + para('Caption') + table(rows, { header: 1 }), 52)).toEqual([[1, 1]])
  })
})

describe('a table longer than a page', () => {
  test('a merged cell cut by the page end: its text stays on the first page, an empty cell goes on below', async () => {
    const body = [['Head A', 'Head B'], ...Array.from({ length: 70 }, (_, i) => [i === 0 ? 'Group' : '', `row ${i + 1} b`])]
    const merge = (r: number, c: number) => (c === 0 && r >= 1 ? (r === 1 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>') : '')
    const v = await open(bareDocx(table(body, { header: 1, tcPr: merge })))
    const n = await pageCount(v)
    expect(n).toBeGreaterThanOrEqual(2)
    const second: string[] = []
    let groups = 0
    for (let k = 1; k <= n; k++) {
      await showPage(v, k)
      const c = await v.page.locator(`.de-page[data-page="${k}"]`).evaluate((pg) =>
        Array.from(pg.querySelectorAll('tr:not(.de-head-copy)')).map((tr) => Array.from((tr as HTMLTableRowElement).cells).map((td) => ({ text: td.innerText.trim(), span: td.rowSpan }))),
      )
      groups += c.flat().filter((x) => x.text === 'Group').length
      // the merged cell spans the rows of its page
      const first = k === 1 ? 1 : 0
      expect(c[first][0].span).toBe(c.length - first)
      expect(c[first][0].text).toBe(k === 1 ? 'Group' : '')
      second.push(...c.map((cells) => cells[cells.length - 1].text))
    }
    expect(groups).toBe(1)
    expect(second).toEqual(body.map((r) => r[1]))
    expect(v.errors).toEqual([])
  })

  const rows = [['Head A', 'Head B'], ...Array.from({ length: 70 }, (_, i) => [`row ${i + 1} a`, `row ${i + 1} b`])]
  const doc = () => bareDocx(para('Before the table') + table(rows, { header: 1 }) + para('After the table'))

  test('goes on on the next pages, its header row repeated, every row shown once', async () => {
    const v = await open(doc())
    const n = await pageCount(v)
    expect(n).toBeGreaterThanOrEqual(2)
    const seen: string[] = []
    for (let k = 1; k <= n; k++) {
      const c = await pageContent(v, k)
      const copies = c.rows.filter((r) => r.copy)
      // page 1 holds the header itself, the following pages a copy of it first
      if (k === 1) expect(copies).toEqual([])
      else if (c.rows.length) expect(c.rows[0]).toEqual({ copy: true, text: 'Head A|Head B' })
      seen.push(...c.rows.filter((r) => !r.copy).map((r) => r.text))
      // no page grows past the paper (a row is never taller than the page here)
      expect(c.height / c.zoom).toBeCloseTo(792 * (4 / 3), 0)
    }
    expect(seen).toEqual(rows.map((r) => r.join('|')))
    expect(v.errors).toEqual([])
  })

  test('the copy is not editable; a cell on the next page is; the header edited on page 1 shows in the copy', async () => {
    const v = await open(doc())
    await showPage(v, 2)
    const copy = v.page.locator('.de-page[data-page="2"] tr.de-head-copy').first()
    expect(await copy.evaluate((tr) => (tr as HTMLElement).isContentEditable)).toBe(false)
    expect(await copy.locator('[data-pa]').count()).toBe(0)
    // a click on the copy does not put the caret in it, typing changes nothing there
    await copy.locator('td').first().click()
    await v.page.keyboard.type('Q')
    expect(await copy.innerText()).not.toContain('Q')

    // a cell of a row on page 2
    const cell = v.page.locator('.de-page[data-page="2"] tr:not(.de-head-copy) td').nth(2)
    const before = (await cell.innerText()).trim()
    await v.selectText(before, 'end')
    await v.page.keyboard.type(' edited')
    expect((await cell.innerText()).trim()).toBe(`${before} edited`)

    // the header itself, on page 1
    await showPage(v, 1)
    await v.selectText('Head A', 'end')
    await v.page.keyboard.type('Z')
    await showPage(v, 2)
    expect(await v.page.locator('.de-page[data-page="2"] tr.de-head-copy').first().innerText()).toContain('Head AZ')

    const blocks = await docxBlocks((await v.save()).bytes)
    const t = blocks.find((b): b is Table => b.t === 'tbl')!
    const cellText = (r: number, c: number) => paraText(t.rows[r].cells[c].blocks[0] as Paragraph)
    expect(cellText(0, 0)).toBe('Head AZ')
    expect(t.rows.flatMap((_, r) => [cellText(r, 0), cellText(r, 1)]).filter((x) => x.includes('edited'))).toEqual([`${before} edited`])
    expect(t.rows.length).toBe(rows.length)
    expect(v.errors).toEqual([])
  })
})
