import { describe, expect, test } from 'bun:test'
import type { Paragraph, Table } from '../src/core/docx/types'
import { bareDocx, para, table } from '../test/docx-builder'
import { docxBlocks, paraText, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

/** opens a built document in the fixture, editable, at `zoom`; waits until every page is laid out */
async function open(doc: Blob | Promise<Blob>, zoom = 1) {
  const v = await viewer({ editable: true, onSave: true }, { wait: false })
  const b64 = Buffer.from(await (await doc).arrayBuffer()).toString('base64')
  await v.page.evaluate(([b, z]) => window.__setProps({ src: new Blob([Uint8Array.from(atob(b as string), (c) => c.charCodeAt(0))]), name: 'split.docx', zoom: z as number }), [b64, zoom] as const)
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
