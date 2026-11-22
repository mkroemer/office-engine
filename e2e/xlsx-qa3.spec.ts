/** QA round 3 regressions of the spreadsheet editor */
import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer, xlsxSheet, type Viewer } from './harness'

useBrowser()

const open = async () => {
  const v = await viewer({ src: 'samples/sample.xlsx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  return v
}
const head = (v: Viewer, kind: 'c' | 'r', text: string) => v.page.locator(`.de-head-${kind}`).filter({ hasText: new RegExp(`^${text}$`) }).first()
const selected = (v: Viewer) => v.page.locator('.de-fx-ref').innerText()
const menu = async (v: Viewer) => (await v.page.locator('.de-ctx [role=menuitem] .de-ctx-label').allInnerTexts()).map((t) => t.trim())

describe('QA 3: whole columns and rows next to a merged title', () => {
  // sample.xlsx: the title in A1 is merged over A1:F1, the table is A3:F9
  test('a column header selects its column only; Delete clears that column, not the table', async () => {
    const v = await open()
    const before = await xlsxSheet((await v.save()).bytes)
    expect(before.merges).toContainEqual({ r0: 0, c0: 0, r1: 0, c1: 5 })
    await head(v, 'c', 'C').click()
    expect(await selected(v)).toMatch(/^C1:C\d+$/)
    await v.page.keyboard.press('Delete')
    await v.page.waitForTimeout(200)
    const s = await xlsxSheet((await v.save()).bytes)
    // the title and the other columns are as they were, column C is empty
    for (const ref of ['A1', 'A3', 'A4', 'B4', 'D4', 'F4']) expect([ref, s.cell(ref)]).toEqual([ref, before.cell(ref)])
    expect(before.cell('C4')).not.toBe('')
    expect(s.cell('C4')).toBe('')
    // Ctrl+Space from a cell of the table: its column
    await v.clickCell('D5')
    await v.page.keyboard.press('Control+Space')
    expect(await selected(v)).toMatch(/^D1:D\d+$/)
    // from the merged title: the columns it spans, as in Excel
    await v.clickCell('B1')
    await v.page.keyboard.press('Control+Space')
    expect(await selected(v)).toMatch(/^A1:F\d+$/)
    expect(v.errors).toEqual([])
  })

  test('typing with a column selected goes to the merged title over its first cell, not into a covered cell', async () => {
    const v = await open()
    await head(v, 'c', 'C').click()
    await v.page.keyboard.type('New title')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(200)
    const s = await xlsxSheet((await v.save()).bytes)
    expect(s.cell('A1')).toBe('New title')
    expect(s.cell('C1')).toBe('')
    expect(v.errors).toEqual([])
  })

  test('the menu of a row header offers rows, the menu of a column header columns', async () => {
    const v = await open()
    await head(v, 'r', '5').click({ button: 'right' })
    const rows = await menu(v)
    expect(rows).toContain('Insert 1 row above')
    expect(rows).toContain('Delete row 5')
    expect(rows.filter((t) => /column/i.test(t) && !/sort|freeze/i.test(t))).toEqual([])
    await v.page.keyboard.press('Escape')
    await head(v, 'c', 'C').click({ button: 'right' })
    const cols = await menu(v)
    expect(cols).toContain('Insert 1 column left')
    expect(cols).toContain('Delete column C')
    expect(cols.filter((t) => /\brows?\b/i.test(t) && !/freeze/i.test(t))).toEqual([])
    await v.page.keyboard.press('Escape')
    // a cell: both
    await v.clickCell('B4')
    const p = await v.cellPoint('B4')
    await v.page.mouse.click(p.x, p.y, { button: 'right' })
    const both = await menu(v)
    expect(both).toContain('Insert 1 row above')
    expect(both).toContain('Insert 1 column left')
    expect(v.errors).toEqual([])
  })
})

describe('QA 3: what cannot be undone says so', () => {
  test('Undo and Redo are off without a step to take; deleting a sheet asks first', async () => {
    const v = await open()
    const undo = v.page.locator('.de-ebar button[aria-label^="Undo"]')
    const redo = v.page.locator('.de-ebar button[aria-label^="Redo"]')
    expect([await undo.isDisabled(), await redo.isDisabled()]).toEqual([true, true])
    await v.typeCell('J10', 'typed')
    expect([await undo.isDisabled(), await redo.isDisabled()]).toEqual([false, true])
    await undo.click()
    await v.page.waitForTimeout(150)
    expect([await undo.isDisabled(), await redo.isDisabled()]).toEqual([true, false])
    await redo.click()
    await v.page.waitForTimeout(150)
    // the sheet's menu: Delete asks, Cancel keeps the sheet
    const tabs = () => v.page.locator('.de-tab[role=tab]').allInnerTexts()
    const before = await tabs()
    expect(before.length).toBeGreaterThan(1)
    const last = v.page.locator('.de-tab[role=tab]').last()
    const name = (await last.innerText()).trim()
    await last.click({ button: 'right' })
    await v.page.locator('.de-ctx [role=menuitem]', { hasText: 'Delete sheet' }).click()
    const dialog = v.page.locator('[role=alertdialog]')
    expect(await dialog.innerText()).toContain(`Delete “${name}”?`)
    expect(await dialog.innerText()).toContain('cannot be undone')
    await dialog.locator('button', { hasText: 'Cancel' }).click()
    expect(await tabs()).toEqual(before)
    // confirmed: the sheet goes, and with it the steps to undo (the button says so)
    await last.click({ button: 'right' })
    await v.page.locator('.de-ctx [role=menuitem]', { hasText: 'Delete sheet' }).click()
    await dialog.locator('button', { hasText: 'Delete' }).click()
    await v.page.waitForTimeout(300)
    expect(await tabs()).toEqual(before.slice(0, -1))
    expect([await undo.isDisabled(), await redo.isDisabled()]).toEqual([true, true])
    expect(v.errors).toEqual([])
  })
})
