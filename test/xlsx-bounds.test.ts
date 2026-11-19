import { expect, test } from 'bun:test'
import { rowsToXlsx } from '../src/core/convert'
import { OpcPackage } from '../src/core/opc'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { MAX_COL, MAX_ROWS, rowSizesOf, XlsxWorkbook } from '../src/core/xlsx/workbook'
import type { SheetProgress } from '../src/core/xlsx/types'

async function workbook(xml: string) {
  const zip = await ZipArchive.open(new BlobSource(await rowsToXlsx([['base']])))
  const pkg = new OpcPackage(zip)
  pkg.setPart('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${xml}</worksheet>`)
  return XlsxWorkbook.open(pkg, 'xl/workbook.xml')
}

async function parsed(xml: string) {
  const wb = await workbook(xml)
  const progress: SheetProgress[] = []
  await wb.loadSheet(0, (p) => progress.push(p))
  const d = wb.sheetsData.get(0)!
  expect(d.error).toBeUndefined()
  return { wb, d, progress }
}

test('worksheet coordinates reject malformed records and retain valid explicit and implicit cells', async () => {
  const invalidRows = ['0', '-1', '1.5', 'NaN', 'Infinity', '1048577', '']
    .map((r) => `<row r="${r}" ht="30"><c r="A1"><v>9</v></c></row>`).join('')
  const invalidCells = ['A0', 'A-1', 'A1.5', 'XFE1', 'ZZZ1', 'A1048577', 'A2', '', '1', 'A1:B2']
    .map((r) => `<c r="${r}"><v>9</v></c>`).join('')
  const { wb, d, progress } = await parsed(
    '<dimension ref="A1:ZZZ4000000000"/><sheetData>' + invalidRows +
    `<row r="1">${invalidCells}<c r="A1"><v>1</v></c><c><v>2</v></c></row>` +
    '<row><c><v>3</v></c></row></sheetData>',
  )
  expect(d.maxRow).toBe(1)
  expect(d.maxCol).toBe(1)
  expect(d.rows.length).toBe(2)
  expect(d.rowAttrs.size).toBe(0)
  expect(progress.at(-1)).toMatchObject({ rows: 2, cols: 2, done: true })
  expect((await wb.block(0, 0, 2)).rows.map((r) => r.cells.map((c) => c[1]))).toEqual([['1', '2'], ['3']])
})

test('worksheet boundaries preserve XFD1048576 and skip implicit overflow', async () => {
  const { wb, d } = await parsed(
    '<sheetData><row r="1048576"><c r="XFD1048576"><v>7</v></c><c><v>8</v></c></row>' +
    '<row><c><v>9</v></c></row></sheetData>',
  )
  expect(d.maxRow).toBe(MAX_ROWS - 1)
  expect(d.maxCol).toBe(MAX_COL - 1)
  expect(d.rows.length).toBe(MAX_ROWS)
  expect((await wb.block(0, MAX_ROWS - 1, MAX_ROWS)).rows[0].cells).toEqual([[MAX_COL - 1, '7', 0, 1]])
})

test('worksheet dimensions, column records, merges, and sizes cannot bypass coordinate validation', async () => {
  const invalidCols = [
    ['-Infinity', '4', '10'], ['0', '4', '10'], ['1.5', '4', '10'], ['1', '16385', '10'],
    ['4', '3', '10'], ['NaN', '4', '10'], ['1', '4', 'Infinity'], ['1', '4', '-1'], ['1', '4', '256'],
  ].map(([min, max, width]) => `<col min="${min}" max="${max}" width="${width}"/>`).join('')
  const invalidMerges = ['A1:A4000000000', 'A1:XFE2', 'A0:B2', 'A1.5:B2', 'B2:A1', 'A1:B2:C3', 'NaN', 'A1:']
    .map((ref) => `<mergeCell ref="${ref}"/>`).join('')
  const { d, progress } = await parsed(
    '<dimension ref="A1:XFD4000000000"/>' +
    '<sheetFormatPr defaultRowHeight="1e300" defaultColWidth="Infinity" baseColWidth="-1"/>' +
    '<sheetViews><sheetView><pane state="frozen" xSplit="Infinity" ySplit="4000000000"/></sheetView></sheetViews>' +
    `<cols>${invalidCols}<col min="16384" max="16384" width="255"/><col min="2" max="3" hidden="1"/></cols>` +
    '<sheetData><row r="1" ht="Infinity"><c><v>1</v></c></row><row r="2" ht="-1"/>' +
    '<row r="3" ht="409.5"/><row r="4" hidden="1"/><row r="5" ht="0"/></sheetData>' +
    `<mergeCells>${invalidMerges}<mergeCell ref="$A$1:$B$2"/><mergeCell ref="XFC1048575:XFD1048576"/></mergeCells>`,
  )
  // columns past the last one and a width past the largest are kept as far as the sheet goes, not dropped
  expect(d.colInfo).toEqual([{ min: 0, max: MAX_COL - 1, w: 70 }, { min: 0, max: 3, w: 1785 }, { min: MAX_COL - 1, max: MAX_COL - 1, w: 1785 }, { min: 1, max: 2, w: 0 }])
  expect(d.merges).toEqual([{ r0: 0, c0: 0, r1: 1, c1: 1 }, { r0: MAX_ROWS - 2, c0: MAX_COL - 2, r1: MAX_ROWS - 1, c1: MAX_COL - 1 }])
  expect(d.defaultRowH).toBe(20)
  expect(d.defaultColW).toBe(64)
  expect(d.frozen).toEqual({ rows: 0, cols: 0 })
  expect(rowSizesOf(d)).toEqual([[2, 546], [3, 0], [4, 0]])
  expect(progress.at(-1)).toMatchObject({ rows: 5, cols: 1 })
})

test('valid worksheet sizes and panes are preserved', async () => {
  const { d, progress } = await parsed(
    '<sheetFormatPr defaultRowHeight="22.5" defaultColWidth="12"/>' +
    '<sheetViews><sheetView><pane state="frozen" xSplit="2" ySplit="3"/></sheetView></sheetViews>' +
    '<cols><col min="1" max="2" width="9.5"/></cols><sheetData><row r="1" ht="18.75"/></sheetData>',
  )
  expect(progress[0]).toMatchObject({ defaultRowH: 30, defaultColW: 84, frozen: { rows: 3, cols: 2 } })
  expect(d.colInfo).toEqual([{ min: 0, max: 1, w: 67 }])
  expect(rowSizesOf(d)).toEqual([[0, 25]])
})

test('tiny defaults stay positive and font-grown OOXML row heights stay bounded', async () => {
  const wb = await workbook('<sheetFormatPr defaultRowHeight="0.001" defaultColWidth="0.001"/>' +
    '<sheetData><row r="1"><c s="1"><v>1</v></c></row></sheetData>')
  wb.styles.xfs.push({ ...wb.styles.xfs[0], fontPt: 1e300 })
  await wb.loadSheet(0, () => {})
  const d = wb.sheetsData.get(0)!
  expect(d.defaultRowH).toBe(1)
  expect(d.defaultColW).toBe(1)
  expect(rowSizesOf(d)).toEqual([[0, 546]])
})

test('non-OOXML loaders retain 500 million rows and larger counts', async () => {
  const wb = await workbook('<sheetData/>')
  for (const count of [500_000_000, 4_000_000_000]) {
    wb.sheetsData.clear()
    wb.loaders.set(wb.sheets[0].path, async (d) => { d.maxRow = count - 1 })
    let last: SheetProgress | undefined
    await wb.loadSheet(0, (p) => { last = p })
    expect(last).toMatchObject({ rows: count, done: true })
  }
})

test('corrupt multi-billion row records remain bounded in an isolated process', async () => {
  // A separate, forcibly time-limited process contains sparse-array/CPU regressions in this repro.
  const proc = Bun.spawn([process.execPath, '--eval', `
    import assert from 'node:assert/strict'
    import { rowsToXlsx } from './src/core/convert.ts'
    import { OpcPackage } from './src/core/opc.ts'
    import { BlobSource } from './src/core/source.ts'
    import { ZipArchive } from './src/core/zip.ts'
    import { XlsxWorkbook } from './src/core/xlsx/workbook.ts'
    const pkg = new OpcPackage(await ZipArchive.open(new BlobSource(await rowsToXlsx([['base']]))))
    pkg.setPart('xl/worksheets/sheet1.xml', '<worksheet><sheetData>' +
      ['4000000000', '9007199254740992', '1e300', 'Infinity'].map(r =>
        '<row r="' + r + '"><c r="A1"><v>9</v></c></row>').join('') +
      '<row r="1"><c r="A1"><v>7</v></c><c r="ZZZZZZZZZZZZZZ1"><v>9</v></c></row></sheetData></worksheet>')
    const wb = await XlsxWorkbook.open(pkg, 'xl/workbook.xml')
    await wb.loadSheet(0, p => { assert.ok(p.rows <= 1); assert.ok(p.cols <= 1) })
    const d = wb.sheetsData.get(0)
    assert.equal(d.error, undefined)
    assert.equal(d.rows.length, 1)
    assert.equal(d.maxRow, 0)
    assert.equal(d.maxCol, 0)
    assert.equal((await wb.block(0, 0, 1)).rows[0].cells[0][1], '7')
  `], { cwd: `${import.meta.dir}/..`, stdout: 'pipe', stderr: 'pipe', timeout: 5000, killSignal: 'SIGKILL' })
  const [exit, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()])
  expect(stderr).toBe('')
  expect(exit).toBe(0)
}, 10_000)
