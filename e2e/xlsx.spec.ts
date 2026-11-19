import { describe, expect, test } from 'bun:test'
import { book, rows, unzip } from '../test/xlsx-builder'
import { fromB64, useBrowser, viewer, xlsxSheet, type Viewer } from './harness'

useBrowser()

const open = async () => {
  const v = await viewer({ src: 'samples/sample.xlsx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  return v
}
const saved = async (v: Viewer, sheet = 0) => xlsxSheet((await v.save()).bytes, sheet)

/** the rendered cell element under a reference: box + inline style */
async function cellBox(v: Viewer, ref: string) {
  const p = await v.cellPoint(ref)
  return v.page.evaluate(({ x, y }) => {
    const cell = document.elementsFromPoint(x, y).map((e) => e.closest<HTMLElement>('.de-cell')).find(Boolean)
    if (!cell) return null
    const r = cell.getBoundingClientRect()
    return { x: r.x, y: r.y, w: r.width, h: r.height, text: cell.innerText, style: cell.getAttribute('style') ?? '', weight: getComputedStyle(cell).fontWeight }
  }, p)
}

/** topmost element at the boundary between two cells (a gridline unless something covers it) */
async function topAt(v: Viewer, a: string, b: string) {
  const p = await v.cellPoint(a)
  const q = await v.cellPoint(b)
  const x = p.x === q.x ? p.x : (p.x + q.x) / 2
  const y = p.y === q.y ? p.y : (p.y + q.y) / 2
  return v.page.evaluate(
    ({ x, y }) => {
      // the gridline sits 1px before the boundary: probe a 3px band
      const hits = [-1, 0, 1].flatMap((d) => document.elementsFromPoint(x + d * 0, y + d)[0]?.className ?? '')
      const hitsX = [-1, 0, 1].flatMap((d) => document.elementsFromPoint(x + d, y)[0]?.className ?? '')
      return [...hits, ...hitsX].join(' ')
    },
    { x, y },
  )
}

describe('xlsx editing', () => {
  test('values, formulas and recalculation', async () => {
    const v = await open()
    await v.typeCell('J10', '5')
    await v.typeCell('J11', '7')
    await v.typeCell('J12', '=SUM(J10:J11)')
    expect(await v.shown('J12')).toBe('12')
    expect(await v.formula('J12')).toBe('=SUM(J10:J11)')
    await v.typeCell('J10', '10')
    expect(await v.shown('J12')).toBe('17')
    expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
    const s = await saved(v)
    expect(s.cell('J12')).toBe('17')
    expect(v.errors).toEqual([])
  })

  test('Escape discards the edit: a new value, a changed value, and a formula', async () => {
    const v = await open()
    await v.typeCell('J10', 'keep')
    // typing into an empty cell, then Escape: stays empty
    await v.clickCell('J11')
    await v.page.keyboard.type('dropped')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(150)
    expect(await v.shown('J11')).toBe('')
    // editing an existing value (F2), then Escape: the old value stays
    await v.clickCell('J10')
    await v.page.keyboard.press('F2')
    await v.page.keyboard.type(' changed')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(150)
    expect(await v.shown('J10')).toBe('keep')
    // a formula with autocomplete open: the first Escape closes the list, the second discards
    await v.clickCell('J12')
    await v.page.keyboard.type('=SU')
    await v.page.locator('.de-fassist-at').waitFor()
    await v.page.keyboard.press('Escape')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(150)
    expect(await v.shown('J12')).toBe('')
    // the keyboard is in the grid: in its hidden input field
    expect(await v.page.evaluate(() => !!document.activeElement?.closest('.de-grid'))).toBe(true)
    expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
    const s = await saved(v)
    expect(s.cell('J10')).toBe('keep')
    expect(s.cell('J11')).toBe('')
    expect(s.cell('J12')).toBe('')
    expect(v.errors).toEqual([])
  })

  test('Shift+Enter and Alt+Enter break the line, in the cell and in the formula bar: the cell wraps and its row grows', async () => {
    const v = await open()
    const rowH = async (n: number) => (await v.page.locator('.de-head-r').filter({ hasText: new RegExp(`^${n}$`) }).first().boundingBox())!.height
    const h0 = await rowH(10)
    const fxRef = () => v.page.locator('.de-fx-ref').innerText()
    // in the cell
    await v.clickCell('J10')
    await v.page.keyboard.type('one')
    await v.page.keyboard.press('Shift+Enter')
    await v.page.keyboard.type('two')
    await v.page.keyboard.press('Alt+Enter')
    await v.page.keyboard.type('three')
    expect(await v.page.locator('textarea.de-cell-edit').inputValue()).toBe('one\ntwo\nthree')
    // Enter alone still commits and moves down
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(200)
    expect(await fxRef()).toBe('J11')
    expect(await v.shown('J10')).toBe('one\ntwo\nthree')
    expect((await cellBox(v, 'J10'))!.style).toContain('pre-wrap')
    expect(await rowH(10)).toBe(h0 * 3)
    // in the formula bar: one line shows at a time, the caret's
    await v.clickCell('J12')
    const bar = v.page.locator('.de-fx-input')
    await bar.click()
    await v.page.keyboard.type('a')
    await v.page.keyboard.press('Shift+Enter')
    await v.page.keyboard.type('b')
    await v.page.keyboard.press('Alt+Enter')
    await v.page.keyboard.type('c')
    expect(await bar.inputValue()).toBe('a\nb\nc')
    // the bar snaps to a whole line in its scroll handler, a frame after the browser scrolled the caret into view
    const lines = (el: HTMLElement) => el.scrollTop / parseFloat(getComputedStyle(el).lineHeight)
    await v.page.waitForFunction((el) => el.scrollTop / parseFloat(getComputedStyle(el).lineHeight) === 2, await bar.elementHandle(), { timeout: 2000 }).catch(() => {})
    expect(await bar.evaluate(lines)).toBe(2)
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(200)
    expect(await v.shown('J12')).toBe('a\nb\nc')
    expect(await rowH(12)).toBe(h0 * 3)
    // the bar of a cell with several lines keeps them, from the first
    expect(await v.formula('J10')).toBe('one\ntwo\nthree')
    expect(await bar.evaluate((el) => el.scrollTop)).toBe(0)
    // on the grid Shift+Enter still moves up
    await v.clickCell('L5')
    await v.page.keyboard.press('Shift+Enter')
    expect(await fxRef()).toBe('L4')
    const s = await saved(v)
    expect(s.cell('J10')).toBe('one\ntwo\nthree')
    expect(s.cell('J12')).toBe('a\nb\nc')
    // one undo per cell: the text, the wrap and the row height
    await v.clickCell('N20')
    await v.page.keyboard.press('Control+z')
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('')
    expect(await rowH(10)).toBe(h0)
    expect(await rowH(12)).toBe(h0)
    expect(v.errors).toEqual([])
  })

  test('merge all: one cell, no inner gridlines, other values cleared', async () => {
    const v = await open()
    await v.typeCell('J10', 'merged')
    await v.typeCell('K11', 'lost')
    const a = (await cellBox(v, 'J10'))!
    await v.selectRange('J10', 'K12')
    await v.pick('Merge cells', 'Merge all')
    await v.page.waitForTimeout(300)
    await v.clickCell('N20')

    // the merged cell spans 2 columns × 3 rows and covers the gridlines inside
    const m = (await cellBox(v, 'K11'))!
    expect(m.text).toBe('merged')
    expect(Math.round(m.x)).toBe(Math.round(a.x))
    expect(m.w).toBeGreaterThan(a.w * 1.8)
    expect(m.h).toBeGreaterThan(a.h * 2.5)
    expect(m.style).toContain('linear-gradient')
    expect(await topAt(v, 'J11', 'K11')).not.toContain('de-gl-v')
    expect(await topAt(v, 'K10', 'K11')).not.toContain('de-gl-h')
    // outside the merge gridlines stay
    expect(await topAt(v, 'L11', 'M11')).toContain('de-gl-v')

    const s = await saved(v)
    expect(s.merges).toContainEqual({ r0: 9, c0: 9, r1: 11, c1: 10 })
    expect(s.cell('J10')).toBe('merged')
    expect(s.cell('K11')).toBe('')

    // clicking the merged cell selects it whole: Unmerge is offered
    await v.clickCell('J10')
    await v.pick('Merge cells', 'Unmerge')
    await v.page.waitForTimeout(300)
    expect((await cellBox(v, 'J10'))!.w).toBeLessThan(a.w * 1.2)
    expect((await saved(v)).merges).not.toContainEqual({ r0: 9, c0: 9, r1: 11, c1: 10 })
  })

  test('merge horizontally and vertically', async () => {
    const v = await open()
    await v.selectRange('J10', 'K11')
    await v.pick('Merge cells', 'Merge horizontally')
    await v.selectRange('M10', 'N11')
    await v.pick('Merge cells', 'Merge vertically')
    await v.page.waitForTimeout(300)
    const s = await saved(v)
    expect(s.merges).toEqual(
      expect.arrayContaining([
        { r0: 9, c0: 9, r1: 9, c1: 10 },
        { r0: 10, c0: 9, r1: 10, c1: 10 },
        { r0: 9, c0: 12, r1: 10, c1: 12 },
        { r0: 9, c0: 13, r1: 10, c1: 13 },
      ]),
    )
  })

  test('bold, fill, currency format, borders', async () => {
    const v = await open()
    await v.typeCell('J10', '12.5')
    await v.clickCell('J10')
    await v.page.keyboard.press('Control+b')
    await v.btn('Format as currency').click()
    await v.btn('Fill color').click()
    await v.page.locator('.de-tb-pop .de-tb-grid button[title="#ffff00"]').click()
    await v.page.waitForTimeout(300)
    await v.clickCell('N20')
    const c = (await cellBox(v, 'J10'))!
    expect(c.text).toBe('$12.50')
    expect(Number(c.weight)).toBeGreaterThanOrEqual(700)
    expect(c.style).toMatch(/background: (rgb\(255, 255, 0\)|#ffff00)/)

    await v.selectRange('J12', 'K13')
    await v.btn('Borders').click()
    await v.page.locator('button[aria-label="All borders"]').click()
    await v.page.waitForTimeout(300)
    expect((await cellBox(v, 'K13'))?.style ?? '').toMatch(/border/)
    expect((await saved(v)).cell('J10')).toBe('$12.50')
  })

  test('insert / delete rows and columns, hide / unhide', async () => {
    const v = await open()
    await v.typeCell('J10', 'here')
    await v.clickCell('J10')
    await v.pick('Insert, delete or hide rows & columns', 'Insert row above')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J11')).toBe('here')
    await v.clickCell('J11')
    await v.pick('Insert, delete or hide rows & columns', 'Insert column left')
    await v.page.waitForTimeout(300)
    expect(await v.shown('K11')).toBe('here')
    await v.clickCell('K11')
    await v.pick('Insert, delete or hide rows & columns', 'Delete column(s)')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J11')).toBe('')

    await v.clickCell('C5')
    await v.pick('Insert, delete or hide rows & columns', 'Hide column(s)')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-head-c').filter({ hasText: /^C$/ }).count()).toBe(0)
    await v.selectRange('B5', 'D5')
    await v.pick('Insert, delete or hide rows & columns', 'Unhide columns')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-head-c').filter({ hasText: /^C$/ }).count()).toBe(1)
  })

  test('sheets: add, rename, switch', async () => {
    const v = await open()
    await v.page.locator('.de-tab-add').click()
    await v.page.waitForTimeout(400)
    const tab = v.page.locator('.de-tab[role=tab][aria-selected=true]')
    expect(await tab.innerText()).toMatch(/^Sheet\d+$/)
    await tab.dblclick()
    await v.page.locator('.de-tab-input').fill('Budget')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-tab[role=tab]').allInnerTexts()).toEqual(['Sales', 'Budget', 'Notes', 'Data'])
    await v.typeCell('A1', 'new sheet')
    const s = await saved(v, 1)
    expect(s.names).toEqual(['Sales', 'Budget', 'Notes', 'Data'])
    expect(s.cell('A1')).toBe('new sheet')
  })

  test('freeze panes and gridlines', async () => {
    const v = await open()
    await v.pick('View: freeze, gridlines', 'Unfreeze')
    await v.pick('View: freeze, gridlines', 'Hide gridlines')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-gl-h').count()).toBe(0)
    let s = await saved(v)
    expect(s.frozen ?? { rows: 0, cols: 0 }).toEqual({ rows: 0, cols: 0 })
    expect(s.gridLines).toBe(false)
    await v.clickCell('C4')
    await v.pick('View: freeze, gridlines', 'Freeze up to current cell')
    await v.page.waitForTimeout(300)
    s = await saved(v)
    expect(s.frozen).toEqual({ rows: 3, cols: 2 })
  })

  test('sort a range', async () => {
    const v = await open()
    await v.typeCell('J10', '3')
    await v.typeCell('J11', '1')
    await v.typeCell('J12', '2')
    await v.selectRange('J10', 'J12')
    await v.pick('Sort range', 'Sort range A → Z')
    await v.page.waitForTimeout(300)
    expect([await v.shown('J10'), await v.shown('J11'), await v.shown('J12')]).toEqual(['1', '2', '3'])
  })

  test('functions menu (AutoSum)', async () => {
    const v = await open()
    await v.typeCell('J10', '4')
    await v.typeCell('J11', '6')
    await v.clickCell('J12')
    await v.pick('Functions', 'SUM')
    await v.page.waitForTimeout(300)
    expect(await v.formula('J12')).toBe('=SUM(J10:J11)')
    expect(await v.shown('J12')).toBe('10')
  })

  test('find & replace', async () => {
    const v = await open()
    await v.clickCell('A3')
    await v.page.keyboard.press('Control+f')
    await v.page.locator('.de-find input[aria-label="Find"]').fill('North')
    await v.page.waitForTimeout(400)
    expect(await v.page.locator('.de-find-count').innerText()).toMatch(/^1 \/ \d+$/)
    await v.btn('Replace').click()
    await v.page.locator('.de-find input[aria-label="Replace with"]').fill('Nord')
    await v.page.locator('.de-find button', { hasText: 'All' }).click()
    await v.page.waitForTimeout(400)
    expect(await v.shown('A3')).toBe('Nord')
  })

  test('undo / redo', async () => {
    const v = await open()
    await v.typeCell('J10', 'first')
    await v.typeCell('J10', 'second')
    await v.btn('Undo').click()
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('first')
    await v.btn('Redo').click()
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('second')
    // keyboard
    await v.typeCell('J10', 'third')
    await v.clickCell('N20')
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('second')
    await v.page.keyboard.press('Control+y')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('third')
  })
})

describe('xlsx select all', () => {
  test('selects the whole sheet', async () => {
    const v = await open()
    await v.clickCell('B2')
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.waitForFunction(() => document.querySelector('.de-fx-ref')?.textContent === 'A1:H62')
    expect(await v.page.locator('.de-head-on').count()).toBeGreaterThan(8)
    expect(await v.page.evaluate(() => getSelection()!.toString())).not.toContain('sample.xlsx')
    expect(v.errors).toEqual([])
  })
})

describe('xlsx keyboard focus', () => {
  test('the grid has the focus on load; Ctrl/⌘+A from the toolbar selects the sheet, in a cell being edited its text', async () => {
    const v = await open()
    // the keyboard is in the grid: in its hidden input field
    expect(await v.page.evaluate(() => !!document.activeElement?.closest('.de-grid'))).toBe(true)
    await v.barButton('View').focus()
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.waitForFunction(() => document.querySelector('.de-fx-ref')?.textContent === 'A1:H62')
    await v.clickCell('B2')
    await v.page.keyboard.type('abc')
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.keyboard.type('x')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(150)
    expect(await v.shown('B2')).toBe('x')
    expect(v.errors).toEqual([])
  })
})

describe('xlsx whole selections and the clipboard (QA 18/11/2026)', () => {
  const rowText = async (s: Awaited<ReturnType<typeof saved>>, sheet: number, r: number) => ((await s.engine.call('xlsx.block', [sheet, r, r + 1])) as any).rows[0]?.cells.map((c: any) => c[1]).join('|') ?? ''

  test('C10: Delete with the whole sheet selected clears every row, the ones not shown too', async () => {
    const v = await open()
    await v.page.locator('.de-tab[role=tab]', { hasText: 'Data' }).click()
    await v.page.waitForFunction(() => document.querySelector('.de-tab[role=tab][aria-selected=true]')?.textContent === 'Data')
    await v.clickCell('B2')
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.keyboard.press('Delete')
    await v.page.waitForTimeout(400)
    const s = await saved(v, 2)
    expect(await rowText(s, 2, 1)).toBe('')
    expect(await rowText(s, 2, 2500)).toBe('')
    expect(await rowText(s, 2, 5000)).toBe('')
    // one step back, all of it
    await v.clickCell('B2')
    await v.page.keyboard.press('ControlOrMeta+z')
    await v.page.waitForTimeout(400)
    const u = await saved(v, 2)
    expect(await rowText(u, 2, 5000)).not.toBe('')
    expect(v.errors).toEqual([])
  })

  test('N1-004 / N1-005: copied cells are pasted as they are (formula, exact value); cut ones are moved, references with them', async () => {
    const v = await open()
    await v.typeCell('J10', '3.14159')
    await v.typeCell('J11', '=J10*2')
    await v.typeCell('M20', '=J10+1')
    await v.selectRange('J10', 'J11')
    await v.page.keyboard.press('ControlOrMeta+c')
    await v.page.waitForTimeout(200)
    await v.clickCell('L10')
    await v.page.keyboard.press('ControlOrMeta+v')
    await v.page.waitForTimeout(300)
    expect(await v.formula('L11')).toBe('=L10*2')
    expect(await v.shown('L11')).toBe('6.28318')
    // cut J10:J11 to J14: M20 reads the cell where it is now
    await v.selectRange('J10', 'J11')
    await v.page.keyboard.press('ControlOrMeta+x')
    await v.page.waitForTimeout(300)
    expect(await v.shown('J10')).toBe('')
    await v.clickCell('J14')
    await v.page.keyboard.press('ControlOrMeta+v')
    await v.page.waitForTimeout(300)
    expect(await v.formula('J15')).toBe('=J14*2')
    expect(await v.formula('M20')).toBe('=J14+1')
    expect(await v.shown('M20')).toBe('4.14159')
    expect(v.errors).toEqual([])
  })
})

// ---------------------------------------------------------------- QA 19/11/2026: input, keys, clipboard, sheets, layout

const fxRef = (v: Viewer) => v.page.locator('.de-fx-ref').innerText()
const atRef = (v: Viewer, ref: string) => v.page.waitForFunction((r) => document.querySelector('.de-fx-ref')?.textContent === r, ref, { timeout: 5000 })
/** the shown text of a cell, once it is `text` (the worker answers after the key) */
async function shownIs(v: Viewer, ref: string, text: string | RegExp) {
  let got = ''
  for (let i = 0; i < 50; i++) {
    got = await v.shown(ref)
    if (typeof text === 'string' ? got === text : text.test(got)) return
    await v.page.waitForTimeout(100)
  }
  expect(got).toEqual(typeof text === 'string' ? text : expect.stringMatching(text))
}
/** a paste event with only the given clipboard formats */
const pasteEvent = (v: Viewer, data: Record<string, string>) =>
  v.page.evaluate((d) => {
    const dt = new DataTransfer()
    for (const [k, x] of Object.entries(d)) dt.setData(k, x)
    ;(document.activeElement ?? document.body).dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }))
  }, data)
const headSize = async (v: Viewer, kind: 'c' | 'r', label: string) => {
  const b = (await v.page.locator(`.de-head-${kind}`).filter({ hasText: new RegExp(`^${label}$`) }).first().boundingBox())!
  return kind === 'c' ? b.width : b.height
}

describe('xlsx input and keys (QA 19/11/2026)', () => {
  test('N1-010: Ctrl/Cmd+S in a cell being edited commits it, saves, and keeps the browser dialog closed', async () => {
    const v = await open()
    await v.page.evaluate(() => {
      const w = window as unknown as { __s: boolean[] }
      w.__s = []
      window.addEventListener('keydown', (e) => (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && w.__s.push(e.defaultPrevented))
    })
    await v.clickCell('J10')
    await v.page.keyboard.type('typed then saved')
    const n = (await v.events('save')).length
    await v.page.keyboard.press('ControlOrMeta+s')
    const e = await v.waitEvent('save', n + 1)
    expect((await xlsxSheet(fromB64(e.b64 as string))).cell('J10')).toBe('typed then saved')
    // from the formula bar too
    await v.clickCell('J11')
    await v.page.locator('.de-fx-input').click()
    await v.page.keyboard.type('from the bar')
    await v.page.keyboard.press('ControlOrMeta+s')
    const f = await v.waitEvent('save', n + 2)
    expect((await xlsxSheet(fromB64(f.b64 as string))).cell('J11')).toBe('from the bar')
    expect(await v.page.evaluate(() => (window as unknown as { __s: boolean[] }).__s)).toEqual([true, true])
    expect(v.errors).toEqual([])
  })

  test('N1-011: text that does not come from a key (insertText, an emoji, an IME composition) starts the edit of the selected cell', async () => {
    const v = await open()
    const ed = () => v.page.locator('textarea.de-cell-edit')
    await v.clickCell('J10')
    await v.page.keyboard.insertText('Đúng')
    expect(await ed().inputValue()).toBe('Đúng')
    await v.page.keyboard.press('Enter')
    await shownIs(v, 'J10', 'Đúng')
    await v.clickCell('J11')
    await v.page.keyboard.insertText('👍 ok')
    await v.page.keyboard.press('Enter')
    await shownIs(v, 'J11', '👍 ok')
    // Telex: the first d is composed with the second one into đ, in the editor that the composition opened
    const cdp = await v.ctx.newCDPSession(v.page)
    await v.clickCell('J12')
    await cdp.send('Input.imeSetComposition', { text: 'd', selectionStart: 1, selectionEnd: 1 })
    await v.page.waitForFunction(() => document.querySelector<HTMLTextAreaElement>('textarea.de-cell-edit')?.value === 'd')
    await cdp.send('Input.imeSetComposition', { text: 'đ', selectionStart: 1, selectionEnd: 1 })
    await cdp.send('Input.insertText', { text: 'đ' })
    await v.page.keyboard.type('ung')
    await v.page.keyboard.press('Enter')
    await shownIs(v, 'J12', 'đung')
    // a letter key still types its letter, and the cell keeps no stray text when nothing is typed
    await v.clickCell('J13')
    await v.page.keyboard.type('abc')
    await v.page.keyboard.press('Enter')
    await shownIs(v, 'J13', 'abc')
    expect(v.errors).toEqual([])
  })

  test('N1-014: Tab, Tab, Enter returns to the column the row was started in; an arrow key forgets it', async () => {
    const v = await open()
    await v.clickCell('K14')
    await v.page.keyboard.type('a')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.type('b')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.type('c')
    await v.page.keyboard.press('Enter')
    await atRef(v, 'K15')
    await shownIs(v, 'M14', 'c')
    // the same on the grid, without typing
    await v.clickCell('L20')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.press('Enter')
    await atRef(v, 'L21')
    // an arrow in between: Enter goes down from where it is
    await v.clickCell('K30')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.press('ArrowRight')
    await v.page.keyboard.press('Enter')
    await atRef(v, 'M31')
    // Shift+Tab then Enter: back to where the Shift+Tab began
    await v.clickCell('M40')
    await v.page.keyboard.press('Shift+Tab')
    await v.page.keyboard.press('Enter')
    await atRef(v, 'M41')
  })

  test('N1-014: Home, End, Ctrl+Home and Ctrl+End', async () => {
    const v = await open()
    await v.clickCell('K20')
    await v.page.keyboard.type('a')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.type('b')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.type('c')
    await v.page.keyboard.press('Enter')
    await v.clickCell('L25')
    await v.page.keyboard.press('Home')
    await atRef(v, 'A25')
    await v.clickCell('K20')
    await v.page.keyboard.press('End')
    await atRef(v, 'M20')
    await v.page.keyboard.press('ControlOrMeta+Home')
    await atRef(v, 'A1')
    // the last cell of the used range: the last row of the data, the last column
    await v.page.keyboard.press('ControlOrMeta+End')
    await atRef(v, 'M62')
    // Shift extends
    await v.page.keyboard.press('ControlOrMeta+Home')
    await v.page.keyboard.press('Shift+ControlOrMeta+End')
    await atRef(v, 'A1:M62')
    expect(v.errors).toEqual([])
  })

  test('N1-014: Ctrl+Arrow stops at the edge of the block of filled cells, like Excel', async () => {
    const v = await open()
    for (const r of [3, 4, 5]) await v.typeCell(`K${r}`, 'x')
    for (const r of [7, 8]) await v.typeCell(`K${r}`, 'y')
    await v.typeCell('M3', 'z')
    await v.clickCell('K3')
    await v.page.keyboard.press('ControlOrMeta+ArrowDown')
    await atRef(v, 'K5')
    await v.page.keyboard.press('ControlOrMeta+ArrowDown')
    await atRef(v, 'K7')
    await v.page.keyboard.press('ControlOrMeta+ArrowDown')
    await atRef(v, 'K8')
    // past the last one: the end of the data
    await v.page.keyboard.press('ControlOrMeta+ArrowDown')
    await atRef(v, 'K62')
    await v.page.keyboard.press('ControlOrMeta+ArrowUp')
    await atRef(v, 'K8')
    await v.page.keyboard.press('ControlOrMeta+ArrowUp')
    await atRef(v, 'K7')
    // along a row: from K3 over the empty L3 to the next filled cell
    await v.page.keyboard.press('ControlOrMeta+Home')
    await v.clickCell('K3')
    await v.page.keyboard.press('ControlOrMeta+ArrowRight')
    await atRef(v, 'M3')
    await v.page.keyboard.press('ControlOrMeta+ArrowLeft')
    await atRef(v, 'K3')
    // a run of three: the end of the run, then the start of it
    await v.page.keyboard.press('ControlOrMeta+Home')
    await v.clickCell('K20')
    await v.page.keyboard.type('a')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.type('b')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.type('c')
    await v.page.keyboard.press('Enter')
    await v.clickCell('K20')
    await v.page.keyboard.press('ControlOrMeta+ArrowRight')
    await atRef(v, 'M20')
    await v.page.keyboard.press('ControlOrMeta+ArrowLeft')
    await atRef(v, 'K20')
    // Shift extends from the block edge
    await v.clickCell('K3')
    await v.page.keyboard.press('Shift+ControlOrMeta+ArrowDown')
    await atRef(v, 'K3:K5')
    expect(v.errors).toEqual([])
  })

  test('N1-025: Shift+Space, Ctrl+Space, Alt+=, Ctrl+; and Ctrl+Enter', async () => {
    const v = await open()
    await v.clickCell('K10')
    await v.page.keyboard.press('Shift+Space')
    await v.page.waitForFunction(() => /^A10:[A-Z]+10$/.test(document.querySelector('.de-fx-ref')!.textContent!))
    await v.clickCell('K10')
    await v.page.keyboard.press('Control+Space')
    await v.page.waitForFunction(() => /^K1:K\d+$/.test(document.querySelector('.de-fx-ref')!.textContent!))
    // AutoSum over the numbers above
    for (const [r, n] of [[24, '1'], [25, '2'], [26, '3']] as const) await v.typeCell(`K${r}`, n)
    await v.clickCell('K27')
    await v.page.keyboard.press('Alt+Equal')
    await shownIs(v, 'K27', '6')
    expect(await v.formula('K27')).toBe('=SUM(K24:K26)')
    // today, as a date
    await v.clickCell('L28')
    await v.page.keyboard.press('ControlOrMeta+;')
    const today = await v.page.evaluate(() => {
      const d = new Date()
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
    })
    await shownIs(v, 'L28', today)
    // Ctrl+Enter: the entry goes into every cell of the selection
    await v.selectRange('K10', 'K12')
    await v.page.keyboard.type('7')
    await v.page.keyboard.press('ControlOrMeta+Enter')
    await shownIs(v, 'K12', '7')
    await shownIs(v, 'K11', '7')
    await shownIs(v, 'K10', '7')
    expect(await fxRef(v)).toBe('K10:K12')
    // a formula keeps its relative references
    await v.selectRange('M10', 'M12')
    await v.page.keyboard.type('=A1')
    await v.page.keyboard.press('ControlOrMeta+Enter')
    await v.page.waitForTimeout(300)
    expect(await v.formula('M12')).toBe('=A3')
    expect(v.errors).toEqual([])
  })
})

describe('xlsx paste (QA 19/11/2026, N1-016)', () => {
  test('one copied value fills the selection; formulas keep their relative references', async () => {
    const v = await open()
    await v.typeCell('K10', '5')
    await v.typeCell('L10', '=A1')
    await v.clickCell('K10')
    await v.page.keyboard.press('ControlOrMeta+c')
    await v.page.waitForTimeout(200)
    await v.selectRange('K12', 'K15')
    await v.page.keyboard.press('ControlOrMeta+v')
    for (const r of [12, 13, 14, 15]) await shownIs(v, `K${r}`, '5')
    await v.clickCell('L10')
    await v.page.keyboard.press('ControlOrMeta+c')
    await v.page.waitForTimeout(200)
    await v.selectRange('L12', 'L14')
    await v.page.keyboard.press('ControlOrMeta+v')
    await v.page.waitForTimeout(300)
    expect(await v.formula('L12')).toBe('=A3')
    expect(await v.formula('L14')).toBe('=A5')
    // plain text from another program
    await v.selectRange('M20', 'N21')
    await pasteEvent(v, { 'text/plain': 'hi' })
    for (const ref of ['M20', 'N20', 'M21', 'N21']) await shownIs(v, ref, 'hi')
    // one undo step
    await v.page.keyboard.press('ControlOrMeta+z')
    await shownIs(v, 'N21', '')
    expect(v.errors).toEqual([])
  })

  test('a paste that would drop data over merged cells is refused with a message', async () => {
    const v = await open()
    await v.selectRange('K30', 'L31')
    await v.pick('Merge cells', 'Merge all')
    await v.page.waitForTimeout(300)
    await v.clickCell('K30')
    await pasteEvent(v, { 'text/plain': 'p\tq\nr\ts' })
    expect(await v.toast()).toMatch(/merged/i)
    expect(await v.shown('K30')).toBe('')
    // one value into the merged cell is fine
    await pasteEvent(v, { 'text/plain': 'one' })
    await shownIs(v, 'K30', 'one')
    expect(v.errors).toEqual([])
  })

  test('a clipboard with only text/html pastes its table', async () => {
    const v = await open()
    await v.clickCell('K22')
    await pasteEvent(v, { 'text/html': '<meta charset="utf-8"><table><tbody><tr><td>a</td><td><b>b</b></td></tr><tr><td>c</td><td>d&amp;e</td></tr></tbody></table>' })
    await shownIs(v, 'K22', 'a')
    await shownIs(v, 'L22', 'b')
    await shownIs(v, 'K23', 'c')
    await shownIs(v, 'L23', 'd&e')
    expect(v.errors).toEqual([])
  })
})

describe('xlsx sheets and merging (QA 19/11/2026, N1-018, N1-019)', () => {
  const rename = async (v: Viewer, tab: string, to: string) => {
    await v.page.locator('.de-tab[role=tab]', { hasText: new RegExp(`^${tab.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`) }).dblclick()
    await v.page.locator('.de-tab-input').fill(to)
    await v.page.keyboard.press('Enter')
  }
  const tabs = (v: Viewer) => v.page.locator('.de-tab[role=tab]').allInnerTexts()
  /** the name of tab `i`, once it is `want` (the message about a name comes before the worker has renamed the sheet) */
  const tabName = async (v: Viewer, i: number, want: string) => {
    for (let k = 0; k < 60 && (await tabs(v))[i] !== want; k++) await v.page.waitForTimeout(50)
    return (await tabs(v))[i]
  }

  test('N1-018: a refused sheet name says why; a name that was cut or changed says so', async () => {
    const v = await open()
    await rename(v, 'Notes', 'sales')
    expect(await v.toast()).toMatch(/already/i)
    expect(await tabs(v)).toEqual(['Sales', 'Notes', 'Data'])
    await rename(v, 'Notes', '   ')
    expect(await v.toast()).toMatch(/empty/i)
    expect(await tabs(v)).toEqual(['Sales', 'Notes', 'Data'])
    await rename(v, 'Notes', 'x'.repeat(40))
    expect(await v.toast()).toMatch(/31 characters/)
    expect(await tabName(v, 1, 'x'.repeat(31))).toBe('x'.repeat(31))
    await rename(v, 'x'.repeat(31), 'a/b?c')
    expect(await v.toast()).toMatch(/cannot contain/i)
    expect(await tabName(v, 1, 'a_b_c')).toBe('a_b_c')
    // a plain rename says nothing
    await v.page.waitForTimeout(100)
    await rename(v, 'a_b_c', 'Budget')
    await v.page.waitForFunction(() => [...document.querySelectorAll('.de-tab[role=tab]')].some((t) => t.textContent === 'Budget'))
    expect(v.errors).toEqual([])
  })

  test('N1-018: merging cells that hold data says what was dropped, and Undo brings it back', async () => {
    const v = await open()
    await v.typeCell('K24', 'a')
    await v.typeCell('L24', 'b')
    await v.selectRange('K24', 'L24')
    await v.pick('Merge cells', 'Merge all')
    expect(await v.toast()).toMatch(/only the upper-left/i)
    await shownIs(v, 'K24', 'a')
    await v.page.keyboard.press('ControlOrMeta+z')
    await v.page.waitForTimeout(400)
    expect((await saved(v)).cell('L24')).toBe('b')
    // nothing dropped, nothing said
    await v.selectRange('K28', 'L28')
    const n = await v.page.locator('.de-toast').count()
    await v.pick('Merge cells', 'Merge all')
    await v.page.waitForTimeout(400)
    expect(await v.page.locator('.de-toast').count()).toBeLessThanOrEqual(n)
    expect(v.errors).toEqual([])
  })

  test('N1-019: "Duplicate sheet" copies the sheet: cells, formats, widths, merges', async () => {
    const v = await open()
    await v.typeCell('K10', 'dup me')
    await v.typeCell('K11', '=K10&"!"')
    await v.selectRange('M10', 'N11')
    await v.pick('Merge cells', 'Merge all')
    await v.page.waitForTimeout(300)
    await v.page.locator('.de-tab[role=tab]', { hasText: /^Sales$/ }).click({ button: 'right' })
    await v.page.locator('.de-menu-item, [role=menuitem]', { hasText: 'Duplicate sheet' }).first().click()
    await v.page.waitForFunction(() => document.querySelector('.de-tab[role=tab][aria-selected=true]')?.textContent === 'Sales (2)')
    expect(await tabs(v)).toEqual(['Sales', 'Sales (2)', 'Notes', 'Data'])
    await shownIs(v, 'K10', 'dup me')
    await shownIs(v, 'K11', 'dup me!')
    const s = await saved(v, 1)
    expect(s.cell('K10')).toBe('dup me')
    expect(s.merges).toContainEqual({ r0: 9, c0: 12, r1: 10, c1: 13 })
    // the original is as it was
    const o = await saved(v, 0)
    expect(o.cell('K10')).toBe('dup me')
    expect(v.errors).toEqual([])
  })
})

describe('xlsx fill handle and F4 (QA 19/11/2026, N1-019)', () => {
  test('dragging the corner of the selection fills down and right', async () => {
    const v = await open()
    await v.typeCell('K10', '5')
    await v.typeCell('L10', '=A1')
    await v.selectRange('K10', 'L10')
    const h = (await v.page.locator('.de-fill-handle').first().boundingBox())!
    const to = await v.cellPoint('K13')
    await v.page.mouse.move(h.x + h.width / 2, h.y + h.height / 2)
    await v.page.mouse.down()
    await v.page.mouse.move(to.x, to.y, { steps: 6 })
    await v.page.mouse.up()
    for (const r of [11, 12, 13]) await shownIs(v, `K${r}`, '5')
    // the selection grows with the fill
    expect(await fxRef(v)).toBe('K10:L13')
    expect(await v.formula('L13')).toBe('=A4')
    // right, from one cell
    await v.typeCell('K20', 'r')
    await v.clickCell('K20')
    const g = (await v.page.locator('.de-fill-handle').first().boundingBox())!
    const right = await v.cellPoint('M20')
    await v.page.mouse.move(g.x + g.width / 2, g.y + g.height / 2)
    await v.page.mouse.down()
    await v.page.mouse.move(right.x, right.y, { steps: 6 })
    await v.page.mouse.up()
    await shownIs(v, 'M20', 'r')
    // one undo step for a drag
    await v.page.keyboard.press('ControlOrMeta+z')
    await shownIs(v, 'M20', '')
    expect(v.errors).toEqual([])
  })

  test('F4 cycles $ on the reference at the caret of a formula', async () => {
    const v = await open()
    await v.clickCell('K10')
    await v.page.keyboard.type('=SUM(A1:B2)+C3')
    const ed = v.page.locator('textarea.de-cell-edit')
    // at the end: C3
    await v.page.keyboard.press('F4')
    expect(await ed.inputValue()).toBe('=SUM(A1:B2)+$C$3')
    await v.page.keyboard.press('F4')
    expect(await ed.inputValue()).toBe('=SUM(A1:B2)+C$3')
    await v.page.keyboard.press('F4')
    expect(await ed.inputValue()).toBe('=SUM(A1:B2)+$C3')
    await v.page.keyboard.press('F4')
    expect(await ed.inputValue()).toBe('=SUM(A1:B2)+C3')
    // inside a range: both ends
    await v.page.keyboard.press('F2')
    await v.page.keyboard.press('Home')
    for (let i = 0; i < 8; i++) await v.page.keyboard.press('ArrowRight')
    await v.page.keyboard.press('F4')
    expect(await ed.inputValue()).toBe('=SUM($A$1:$B$2)+C3')
    await v.page.keyboard.press('Escape')
    expect(v.errors).toEqual([])
  })
})

describe('xlsx layout (QA 19/11/2026, N1-021, N1-023)', () => {
  test('N1-023: Wrap text grows the row to show the text', async () => {
    const v = await open()
    await v.typeCell('K7', 'a long text that does not fit in one line of a narrow column at all, really it does not')
    const h0 = await headSize(v, 'r', '7')
    await v.clickCell('K7')
    await v.btn('Wrap text').click()
    await v.page.waitForFunction(([h]) => [...document.querySelectorAll('.de-head-r')].find((e) => e.textContent === '7')!.getBoundingClientRect().height > (h as number) * 3, [h0])
    // the text is all there: the cell is as tall as the row
    const box = (await cellBox(v, 'K7'))!
    expect(box.h).toBeGreaterThan(h0 * 3)
    expect(v.errors).toEqual([])
  })

  test('N1-023: double click on the edge of a column fits it to its text', async () => {
    const v = await open()
    await v.typeCell('K5', 'fit this column to the text in it')
    const w0 = await headSize(v, 'c', 'K')
    const b = (await v.page.locator('.de-grip-c[data-grip-col="10"]').boundingBox())!
    await v.page.mouse.dblclick(b.x + b.width / 2, b.y + b.height / 2)
    await v.page.waitForFunction(([w]) => [...document.querySelectorAll('.de-head-c')].find((e) => e.textContent === 'K')!.getBoundingClientRect().width > (w as number) * 2, [w0])
    const w1 = await headSize(v, 'c', 'K')
    // the text fits (a little air at the edge), and the width was written
    const s = (await cellBox(v, 'K5'))!
    expect(s.w).toBe(w1)
    expect(await v.page.evaluate(() => { const c = [...document.querySelectorAll<HTMLElement>('.de-cell')].find((e) => e.textContent === 'fit this column to the text in it')!; return c.scrollWidth <= c.clientWidth + 1 })).toBe(true)
    expect(v.errors).toEqual([])
  })

  test('N1-023: a column widens when a currency typed into it would show ####', async () => {
    const v = await open()
    await v.typeCell('K20', '1234567.891')
    const w0 = await headSize(v, 'c', 'K')
    await v.clickCell('K20')
    await v.btn('Format as currency').click()
    await shownIs(v, 'K20', '$1,234,567.89')
    expect(await headSize(v, 'c', 'K')).toBeGreaterThan(w0)
    expect(v.errors).toEqual([])
  })

  test('N1-012: a number in a column of the default width is shown in what fits, in the font the browser has: never wider than its cell', async () => {
    const v = await open()
    const typed = ['1234567', '0.123456789', '123456789012', '-1234567.891', '0.000012345678']
    for (const [i, t] of typed.entries()) await v.typeCell(`K${20 + i}`, t)
    await v.clickCell('J10')
    for (const [i, t] of typed.entries()) {
      const p = await v.cellPoint(`K${20 + i}`)
      const box = await v.page.evaluate(({ x, y }) => {
        const cell = document.elementsFromPoint(x, y).map((e) => e.closest<HTMLElement>('.de-cell')).find(Boolean)!
        const span = cell.querySelector('span')!
        return { text: span.textContent ?? '', left: span.getBoundingClientRect().left, cellLeft: cell.getBoundingClientRect().left }
      }, p)
      // (a number is drawn from the right edge: one too wide loses its first digits)
      expect(box.left).toBeGreaterThanOrEqual(box.cellLeft)
      expect(box.text).not.toMatch(/^#/)
      // (how many digits fit is the font's business: the first one at least)
      expect(Number(box.text)).toBeCloseTo(Number(t), -Math.floor(Math.log10(Math.abs(Number(t)))))
    }
    expect(v.errors).toEqual([])
  })

  test('N1-023: a number format typed by hand ("Custom format…")', async () => {
    const v = await open()
    await v.typeCell('J10', '1234567')
    await v.typeCell('J11', '2024-01-15')
    const custom = async (ref: string, code: string) => {
      await v.clickCell(ref)
      await v.pick('Number format', 'Custom format…')
      const input = v.page.locator('.de-tb-pop .de-prompt input')
      await input.fill(code)
      await input.press('Enter')
      await v.page.waitForTimeout(300)
    }
    await custom('J10', '#,##0 "đ"')
    expect(await v.shown('J10')).toBe('1,234,567 đ')
    await custom('J11', 'ddd dd/mm')
    expect(await v.shown('J11')).toBe('Mon 15/01')
    // the grid has the keys again, and the format of the cell is offered as "Custom"
    await v.clickCell('J10')
    expect(await v.page.locator('button[aria-label^="Number format"]').first().getAttribute('aria-label')).toBe('Number format: Custom')
    const s = await saved(v)
    expect(s.cell('J10')).toBe('1,234,567 đ')
    expect(s.cell('J11')).toBe('Mon 15/01')
    expect(v.errors).toEqual([])
  })

  test('P2-037: text tilted from the bar is shown whole: its row grows, and it goes out over the cells to the right', async () => {
    const v = await open()
    await v.typeCell('J10', 'Rotate this text 45')
    await v.clickCell('J10')
    const before = (await cellBox(v, 'J10'))!.h
    await v.pick('Text rotation', 'Tilt up')
    await v.page.waitForTimeout(500)
    await v.clickCell('L12')
    const p = await v.cellPoint('J10')
    const box = await v.page.evaluate(({ x, y }) => {
      const cell = document.elementsFromPoint(x, y).map((e) => e.closest<HTMLElement>('.de-cell')).find(Boolean)!
      const c = cell.getBoundingClientRect()
      const t = cell.querySelector('span')!.getBoundingClientRect()
      return { top: c.top, bottom: c.bottom, left: c.left, h: c.height, tTop: t.top, tBottom: t.bottom, tLeft: t.left }
    }, p)
    expect(box.h).toBeGreaterThan(before + 30)
    // the turned text is inside the rows of its cell, and starts in it
    expect(box.tTop).toBeGreaterThanOrEqual(box.top - 1)
    expect(box.tBottom).toBeLessThanOrEqual(box.bottom + 1)
    expect(box.tLeft).toBeGreaterThanOrEqual(box.left - 1)
    expect(v.errors).toEqual([])
  })

  test('N1-021: at 390 px wide every button of the top bar can be reached', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', onSave: true }, { viewport: { width: 390, height: 844 } })
    const r = await v.page.evaluate(() => {
      const bar = document.querySelector<HTMLElement>('.de-bar')!
      const edge = Math.max(...[...bar.querySelectorAll<HTMLElement>('button')].map((b) => b.getBoundingClientRect().right))
      const scrolls = ['auto', 'scroll'].includes(getComputedStyle(bar).overflowX)
      return { edge, over: bar.scrollWidth > bar.clientWidth + 1, scrolls, comments: !!bar.querySelector('.de-cm-toggle') }
    })
    expect(r.comments).toBe(true)
    // every button is on the screen, or the bar scrolls to it
    expect(r.edge <= 391 || (r.over && r.scrolls)).toBe(true)
    expect(await v.page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390)
  })
})

describe('xlsx rendering (QA 19/11/2026, P2-037)', () => {
  // a shared string with parts in fonts of their own
  const RICH = {
    parts: {
      'xl/sharedStrings.xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">' +
        '<si><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> then plain</t></r></si></sst>',
    },
    rels: '<Relationship Id="rIdSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
    types: '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
  }

  test('pattern fills are mixed colors, diagonal borders are drawn, shrink to fit scales the text', async () => {
    const styles =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill></fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"><color rgb="FFFF0000"/></diagonal></border></borders>' +
      '<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="1" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment shrinkToFit="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment textRotation="45"/></xf></cellXfs></styleSheet>'
    const blob = await book({ sheets: [{ name: 'Look', rows: rows([[{ s: 1, v: 'pat' }, { s: 2, v: 'dia' }, { s: 3, v: 'shrink this long text into its cell' }, 'plain long text that is not shrunk at all'], [{ s: 4, v: 'Rotate 45 degrees' }]]) + '<row r="3"><c r="A3" t="s"><v>0</v></c></row>' }], styles, ...RICH })
    const b64 = Buffer.from(await blob.arrayBuffer()).toString('base64')
    const v = await viewer({ src: `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${b64}`, name: 'look.xlsx' })
    const look = (text: string) =>
      v.page.evaluate((t) => {
        const cell = [...document.querySelectorAll<HTMLElement>('.de-cell')].find((c) => c.textContent === t)!
        const span = cell.querySelector('span') as HTMLElement
        const cs = getComputedStyle(cell)
        return { bg: cs.backgroundColor, image: cs.backgroundImage, transform: span.style.transform, textW: span.getBoundingClientRect().width, cellW: cell.getBoundingClientRect().width }
      }, text)
    expect((await look('pat')).bg).toBe('rgb(255, 223, 223)')
    const d = await look('dia')
    expect(d.image).toContain('linear-gradient')
    expect(d.image).toContain('rgb(255, 0, 0)')
    const sh = await look('shrink this long text into its cell')
    expect(sh.transform).toContain('scale(')
    expect(sh.textW).toBeLessThanOrEqual(sh.cellW)
    expect((await look('plain long text that is not shrunk at all')).transform).toBe('')
    // tilted text in a row that says no height: the row is as tall as the text needs, in the view
    const tilt = await v.page.evaluate(() => {
      const cell = [...document.querySelectorAll<HTMLElement>('.de-cell')].find((c) => c.textContent === 'Rotate 45 degrees')!
      const c = cell.getBoundingClientRect()
      const t = cell.querySelector('span')!.getBoundingClientRect()
      return { h: c.height, top: c.top, bottom: c.bottom, tTop: t.top, tBottom: t.bottom }
    })
    expect(tilt.h).toBeGreaterThan(60)
    expect(tilt.tTop).toBeGreaterThanOrEqual(tilt.top - 1)
    expect(tilt.tBottom).toBeLessThanOrEqual(tilt.bottom + 1)
    // rich text: each part in its font
    const rich = await v.page.evaluate(() => {
      const cell = [...document.querySelectorAll<HTMLElement>('.de-cell')].find((c) => c.textContent === 'Bold then plain')!
      return [...cell.querySelectorAll<HTMLElement>('span > span')].map((x) => ({ t: x.textContent, w: getComputedStyle(x).fontWeight, color: getComputedStyle(x).color }))
    })
    expect(rich.map((x) => x.t)).toEqual(['Bold', ' then plain'])
    expect(Number(rich[0].w)).toBeGreaterThanOrEqual(700)
    expect(rich[0].color).toBe('rgb(255, 0, 0)')
    expect(Number(rich[1].w)).toBeLessThan(700)
    expect(v.errors).toEqual([])
  })
})

describe('xlsx whole columns and rows (QA 19/11/2026, P2-042)', () => {
  const yellow = /background: (rgb\(255, 255, 0\)|#ffff00)/
  const fill = async (v: Viewer) => {
    await v.btn('Fill color').click()
    await v.page.locator('.de-tb-pop .de-tb-grid button[title="#ffff00"]').click()
    await v.page.waitForTimeout(400)
  }

  test('a fill given to a column or a row selected by its header is the format of the column / row: shown in its empty cells, kept by what is typed there, saved without a cell per place', async () => {
    const v = await open()
    await v.page.locator('.de-head-c[data-col="12"]').click()
    await fill(v)
    await v.clickCell('J10')
    expect((await cellBox(v, 'M20'))?.style ?? '').toMatch(yellow)
    expect((await cellBox(v, 'L20'))?.style ?? '').not.toMatch(yellow)
    await v.typeCell('M15', 'typed')
    expect((await cellBox(v, 'M15'))?.style ?? '').toMatch(yellow)

    await v.page.locator('.de-head-r[data-row="21"]').click()
    await fill(v)
    await v.clickCell('J10')
    expect((await cellBox(v, 'K22'))?.style ?? '').toMatch(yellow)
    expect((await cellBox(v, 'K23'))?.style ?? '').not.toMatch(yellow)

    const z = await unzip(new Blob([(await v.save()).bytes]))
    const xml = await z.text(z.paths.find((p) => /worksheets\/sheet1\.xml$/.test(p))!)
    expect(xml).toMatch(/<col [^>]*min="13"[^>]*style="\d+"/)
    expect(xml).toMatch(/<row r="22"[^>]* s="\d+"[^>]*customFormat="1"|<row r="22"[^>]*customFormat="1"[^>]* s="\d+"/)
    // the column holds what was typed and where the row crosses it, not a cell per row
    expect((xml.match(/<c r="M\d+"/g) ?? []).length).toBeLessThan(10)

    // Undo takes the formats back, one at a time (the Save button has the focus: back to the grid)
    await v.clickCell('J10')
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(300)
    expect((await cellBox(v, 'K22'))?.style ?? '').not.toMatch(yellow)
    expect((await cellBox(v, 'M20'))?.style ?? '').toMatch(yellow)
    await v.page.keyboard.press('Control+z')
    await v.page.keyboard.press('Control+z')
    await v.page.waitForTimeout(300)
    expect((await cellBox(v, 'M20'))?.style ?? '').not.toMatch(yellow)
    expect(v.errors).toEqual([])
  })
})
