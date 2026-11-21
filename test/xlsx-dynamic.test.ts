/** Dynamic array formulas in a workbook (spilling, #SPILL!, saving), defined names, tables and 3-D references. */
import { describe, expect, test } from 'bun:test'
import { shiftedClip } from '../src/core/xlsx/clip'
import { at, book, openBook, rows } from './xlsx-builder'

const SHEET = 'xl/worksheets/sheet1.xml'
type Opened = Awaited<ReturnType<typeof openBook>>
const shown = async (o: Opened, refs: string[], sheet = 0) => Promise.all(refs.map((ref) => o.shown(sheet, at(ref).r, at(ref).c)))
const cellXml = (xml: string, ref: string) => new RegExp(`<c r="${ref}"[^>]*?(?:/>|>.*?</c>)`).exec(xml)?.[0] ?? ''

describe('dynamic array formulas', () => {
  test('a formula that gives several values fills the cells next to it', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[5], [6], [7]]) }] }))
    await o.set(0, 'B1', '=SEQUENCE(3)')
    expect(await shown(o, ['B1', 'B2', 'B3', 'B4'])).toEqual(['1', '2', '3', ''])
    await o.set(0, 'D1', '=A1:A3*2')
    expect(await shown(o, ['D1', 'D2', 'D3'])).toEqual(['10', '12', '14'])
    // what reads the cells filled follows them
    await o.set(0, 'C1', '=SUM(B1:B4)')
    await o.set(0, 'C2', '=SUM(B1#)')
    expect(await shown(o, ['C1', 'C2'])).toEqual(['6', '6'])
    await o.set(0, 'B1', '=SEQUENCE(4)')
    expect(await shown(o, ['B4', 'C1', 'C2'])).toEqual(['4', '10', '10'])
    // fewer values: the cells it filled before are empty again
    await o.set(0, 'B1', '=SEQUENCE(2)')
    expect(await shown(o, ['B1', 'B2', 'B3', 'B4', 'C1', 'C2'])).toEqual(['1', '2', '', '', '3', '3'])
    // one value: no cells filled
    await o.set(0, 'B1', '=SUM(A1:A3)')
    expect(await shown(o, ['B1', 'B2'])).toEqual(['18', ''])
  })

  test('a cell in the way: #SPILL!, until it is cleared', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'B2', 'x')
    await o.set(0, 'B1', '=SEQUENCE(3)')
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['#SPILL!', 'x', ''])
    await o.set(0, 'B2', '')
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['1', '2', '3'])
    // typing over a value it filled blocks it; undoing the typing lets it fill the cells again
    await o.set(0, 'B3', 'y')
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['#SPILL!', '', 'y'])
    await o.engine.call('xlsx.undo', [])
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['1', '2', '3'])
    await o.engine.call('xlsx.redo', [])
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['#SPILL!', '', 'y'])
    await o.engine.call('xlsx.undo', [])
    // a formula in the way blocks it too
    await o.set(0, 'B2', '=1+1')
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['#SPILL!', '2', ''])
    await o.set(0, 'B2', '')
    // the formula removed: its cells are empty
    await o.set(0, 'B1', '')
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['', '', ''])
    await o.engine.call('xlsx.undo', [])
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['1', '2', '3'])
  })

  test('the values follow the cells the formula reads', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[5], [1], [7], [2]]) }] }))
    await o.set(0, 'C1', '=FILTER(A1:A4,A1:A4>4)')
    await o.set(0, 'E1', '=SORT(A1:A4)')
    expect(await shown(o, ['C1', 'C2', 'C3'])).toEqual(['5', '7', ''])
    expect(await shown(o, ['E1', 'E2', 'E3', 'E4'])).toEqual(['1', '2', '5', '7'])
    await o.set(0, 'A2', '9')
    expect(await shown(o, ['C1', 'C2', 'C3'])).toEqual(['5', '9', '7'])
    expect(await shown(o, ['E1', 'E2', 'E3', 'E4'])).toEqual(['2', '5', '7', '9'])
    await o.set(0, 'A1', '0')
    expect(await shown(o, ['C1', 'C2', 'C3'])).toEqual(['9', '7', ''])
    await o.set(0, 'A2', '0')
    await o.set(0, 'A3', '0')
    expect(await shown(o, ['C1', 'C2'])).toEqual(['#CALC!', ''])
  })

  test('a formula of the file that reads a range where one value goes still takes the one of its row', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[5, { f: 'A1:A3*2', v: 10 }], [6, { f: 'A1:A3*2', v: 12 }], [7]]) }] }))
    await o.set(0, 'A2', '8')
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['10', '16', ''])
  })

  test('a formula of the file filled down, typed over and put back, stays what it was', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[5, { f: 'A1:A3*2', v: 10 }], [6], [7]]) }] }))
    // filled down: each copy takes the value of its own row, as the formula it is a copy of
    await o.engine.call('xlsx.sheetOp', [0, { op: 'fill', range: { r0: 0, c0: 1, r1: 2, c1: 1 }, dir: 'down' }])
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['10', '12', '14'])
    expect(await o.formula(0, 1, 1)).toBe('=A2:A4*2')
    // typed over and undone: not a dynamic array formula after
    await o.set(0, 'B1', 'x')
    await o.engine.call('xlsx.undo', [])
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['10', '12', '14'])
    // entered again as it is: nothing changes
    await o.set(0, 'B1', '=A1:A3*2')
    expect(await shown(o, ['B1', 'B2'])).toEqual(['10', '12'])
    // another formula typed: a dynamic array formula, and the cells under it are in its way
    await o.set(0, 'B1', '=A1:A3*3')
    expect(await shown(o, ['B1', 'B2'])).toEqual(['#SPILL!', '12'])
    await o.engine.call('xlsx.undo', [])
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['10', '12', '14'])
    await o.engine.call('xlsx.redo', [])
    expect(await shown(o, ['B1', 'B2'])).toEqual(['#SPILL!', '12'])
  })

  test('copied with the cells it fills, it fills the cells where it is pasted', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[5], [6], [7]]) }] }))
    await o.set(0, 'C1', '=A1:A3*2')
    const clip: any = await o.engine.call('xlsx.copy', [0, { r0: 0, c0: 2, r1: 2, c1: 2 }])
    expect(clip.text).toBe('10\n12\n14')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'paste', range: { r0: 0, c0: 4, r1: 2, c1: 4 }, cells: clip.cells, from: { sheet: clip.sheet, r0: 0, c0: 2 } }])
    expect(await o.formula(0, 0, 4)).toBe('=C1:C3*2')
    expect(await shown(o, ['E1', 'E2', 'E3'])).toEqual(['20', '24', '28'])
    expect(await o.formula(0, 1, 4)).toBeUndefined()
    // one of the values alone is a value
    const one: any = await o.engine.call('xlsx.copy', [0, { r0: 1, c0: 2, r1: 1, c1: 2 }])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'paste', range: { r0: 5, c0: 0, r1: 5, c1: 0 }, cells: one.cells, from: { sheet: one.sheet, r0: 1, c0: 2 } }])
    expect(await shown(o, ['A6'])).toEqual(['12'])
    // filled to the right: a dynamic array formula too
    await o.engine.call('xlsx.sheetOp', [0, { op: 'fill', range: { r0: 0, c0: 6, r1: 0, c1: 6 }, dir: 'right' }])
    await o.set(0, 'G1', '=SEQUENCE(2)')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'fill', range: { r0: 0, c0: 6, r1: 0, c1: 7 }, dir: 'right' }])
    expect(await shown(o, ['G1', 'G2', 'H1', 'H2'])).toEqual(['1', '2', '1', '2'])
  })

  test('#35: a copied block repeated by the view keeps its formulas dynamic in every copy', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[5], [6], [7]]) }] }))
    await o.set(0, 'C1', '=A1:A3*2')
    const clip: any = await o.engine.call('xlsx.copy', [0, { r0: 0, c0: 2, r1: 0, c1: 2 }])
    const cell = clip.cells[0][0]
    expect(cell[4]).toBe(true)
    // (as the view repeats one copied cell into E1:F1: the worker moves both by the first copy's distance)
    const second = shiftedClip(cell, 0, 1)
    expect(second).toEqual([cell[0], cell[1], cell[2], 'B1:B3*2', true])
    expect(shiftedClip(cell, 0, 0)).toBe(cell)
    expect(shiftedClip(null, 1, 1)).toBeNull()
    await o.engine.call('xlsx.sheetOp', [0, { op: 'paste', range: { r0: 0, c0: 4, r1: 0, c1: 5 }, cells: [[cell, second]], from: { sheet: clip.sheet, r0: 0, c0: 2 } }])
    expect(await o.formula(0, 0, 5)).toBe('=D1:D3*2')
    expect(await shown(o, ['E1', 'E3', 'F1', 'F3'])).toEqual(['20', '28', '0', '0'])
    expect(cellXml(await (await o.save()).text(SHEET), 'F1')).toMatch(/ cm="\d+"/)
  })

  test('#SPILL! of a file read: cleared when the cell in the way is', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'B2', 'x')
    await o.set(0, 'B1', '=SEQUENCE(3)')
    const re = await openBook((await o.save()).blob)
    expect(await shown(re, ['B1', 'B2'])).toEqual(['#SPILL!', 'x'])
    await re.set(0, 'B2', '')
    expect(await shown(re, ['B1', 'B2', 'B3'])).toEqual(['1', '2', '3'])
  })

  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  const META =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="${NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>` +
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
  /** a workbook as Excel saves dynamic array formulas: no word to compute it when it is read */
  const excel = (sheet: string) =>
    book({
      sheets: [{ name: 'S', rows: sheet }],
      parts: { 'xl/metadata.xml': META },
      rels: '<Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>',
      types: '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>',
    })

  test('a file of Excel: its dynamic array formulas, one of them showing #SPILL!', async () => {
    const o = await openBook(
      await excel(
        '<row r="1"><c r="A1"><v>5</v></c><c r="B1" cm="1" t="e"><f t="array" ref="B1">_xlfn.SEQUENCE(3)</f><v>#SPILL!</v></c><c r="D1" cm="1"><f t="array" ref="D1:D2">_xlfn._xlws.FILTER(A1:A3,A1:A3&gt;4)</f><v>5</v></c></row>' +
          '<row r="2"><c r="A2"><v>1</v></c><c r="B2" t="inlineStr"><is><t>x</t></is></c><c r="D2"><v>7</v></c></row><row r="3"><c r="A3"><v>7</v></c></row>',
      ),
    )
    expect(await shown(o, ['B1', 'B2', 'D1', 'D2'])).toEqual(['#SPILL!', 'x', '5', '7'])
    expect(await o.formula(0, 0, 1)).toBe('=SEQUENCE(3)')
    // the cell in the way cleared: the formula fills its cells
    await o.set(0, 'B2', '')
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['1', '2', '3'])
    // the other follows the cells it reads, with the cells it fills
    await o.set(0, 'A2', '6')
    expect(await shown(o, ['D1', 'D2', 'D3'])).toEqual(['5', '6', '7'])
    await o.set(0, 'A1', '0')
    expect(await shown(o, ['D1', 'D2', 'D3'])).toEqual(['6', '7', ''])
    const xml = await (await o.save()).text(SHEET)
    expect(cellXml(xml, 'B1')).toContain('<f t="array" ref="B1:B3">')
    expect(cellXml(xml, 'B1')).toContain('cm="1"')
    expect(cellXml(xml, 'D3')).toBe('')
  })

  test('a formula waiting for a cell another formula fills goes on when that one gives it back', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'E1', '=SEQUENCE(2)')
    await o.set(0, 'A2', '=SEQUENCE(1,5)')
    expect(await shown(o, ['A2', 'E2'])).toEqual(['#SPILL!', '2'])
    await o.set(0, 'E1', '=SEQUENCE(1)')
    expect(await shown(o, ['A2', 'B2', 'E2', 'E1'])).toEqual(['1', '2', '5', '1'])
    // the formula put back finds the cell taken: the one that fills it keeps it
    await o.engine.call('xlsx.undo', [])
    expect(await shown(o, ['A2', 'E2', 'E1'])).toEqual(['1', '5', '#SPILL!'])
    // and when the formula in the way is replaced by one that fills nothing
    const p = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await p.set(0, 'E1', '=SEQUENCE(2)')
    await p.set(0, 'A2', '=SEQUENCE(1,5)')
    await p.set(0, 'E1', '=1+1')
    expect(await shown(p, ['A2', 'E2', 'E1'])).toEqual(['1', '5', '2'])
  })

  test('the sheet reaches as far as the cells filled', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    const res: any = await o.set(0, 'B1', '=SEQUENCE(10,3)')
    expect(res.grown).toEqual({ rows: 10, cols: 4 })
    await o.set(0, 'F1', '=SUM(B:B)')
    await o.set(0, 'F2', '=COUNT(2:2)')
    expect(await shown(o, ['F1', 'F2'])).toEqual(['145', '3'])
    const xml = await (await o.save()).text(SHEET)
    expect(cellXml(xml, 'D10')).toContain('<v>30</v>')
  })

  test('a range holding a dynamic array formula is not sorted', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[3], [1], [2]]) }] }))
    await o.set(0, 'B1', '=A1:A3*2')
    await expect(o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range: { r0: 0, c0: 0, r1: 2, c1: 1 }, col: 0 }])).rejects.toThrow(/dynamic array/)
    // the cells it reads are sorted, and it follows
    await o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range: { r0: 0, c0: 0, r1: 2, c1: 0 }, col: 0 }])
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['2', '4', '6'])
  })

  test('saved as an array formula marked dynamic, and read back as one', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[5], [1], [7]]) }] }))
    await o.set(0, 'C1', '=FILTER(A1:A3,A1:A3>4)')
    await o.set(0, 'E1', '=LET(x,A1,x*2)')
    await o.set(0, 'F1', '=UNIQUE(A1:A1)')
    const saved = await o.save()
    const xml = await saved.text(SHEET)
    expect(cellXml(xml, 'C1')).toContain('<f t="array" ref="C1:C2">_xlfn._xlws.FILTER(A1:A3,A1:A3&gt;4)</f>')
    expect(cellXml(xml, 'C1')).toContain('cm="1"')
    expect(cellXml(xml, 'C2')).toContain('<v>7</v>')
    expect(cellXml(xml, 'C2')).not.toContain('<f')
    // a formula of single values is written as any other; the names of LET with their prefix
    expect(cellXml(xml, 'E1')).toContain('<f>_xlfn.LET(_xlpm.x,A1,_xlpm.x*2)</f>')
    expect(cellXml(xml, 'E1')).not.toContain('cm=')
    // one that gives a single value for now is dynamic all the same
    expect(cellXml(xml, 'F1')).toContain('<f t="array" ref="F1">')
    expect(saved.has('xl/metadata.xml')).toBe(true)
    expect(await saved.text('xl/metadata.xml')).toContain('fDynamic="1"')
    expect(await saved.text('[Content_Types].xml')).toContain('/xl/metadata.xml')
    expect(await saved.text('xl/_rels/workbook.xml.rels')).toContain('sheetMetadata')

    const re = await openBook(saved.blob)
    expect(await re.formula(0, 0, 2)).toBe('=FILTER(A1:A3,A1:A3>4)')
    expect(await re.formula(0, 0, 4)).toBe('=LET(x,A1,x*2)')
    expect(await shown(re, ['C1', 'C2', 'C3', 'E1'])).toEqual(['5', '7', '', '10'])
    await re.set(0, 'A2', '6')
    expect(await shown(re, ['C1', 'C2', 'C3'])).toEqual(['5', '6', '7'])
    await re.set(0, 'A1', '1')
    await re.set(0, 'A2', '1')
    expect(await shown(re, ['C1', 'C2', 'C3'])).toEqual(['7', '', ''])
    const again = await (await re.save()).text(SHEET)
    expect(cellXml(again, 'C1')).toContain('<f t="array" ref="C1">')
    expect(cellXml(again, 'C2')).not.toContain('<v>')
  })

  test('an array formula entered over a range keeps its cells', async () => {
    const sheet = `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="B1:B3">A1:A3*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>4</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>6</v></c></row>`
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: sheet }] }))
    await o.set(0, 'A2', '10')
    expect(await shown(o, ['B1', 'B2', 'B3'])).toEqual(['2', '20', '6'])
    const xml = await (await o.save()).text(SHEET)
    expect(cellXml(xml, 'B1')).toContain('<f t="array" ref="B1:B3">')
    expect(cellXml(xml, 'B1')).not.toContain('cm=')
  })

  test('rows inserted and deleted through the cells it fills', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[5], [6], [7]]) }] }))
    await o.set(0, 'C1', '=A1:A3')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 1, n: 1 }])
    expect(await shown(o, ['C1', 'C2', 'C3', 'C4', 'C5'])).toEqual(['5', '0', '6', '7', ''])
    await o.engine.call('xlsx.undo', [])
    expect(await shown(o, ['C1', 'C2', 'C3', 'C4'])).toEqual(['5', '6', '7', ''])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'deleteRows', at: 2, n: 1 }])
    expect(await shown(o, ['C1', 'C2', 'C3'])).toEqual(['5', '6', ''])
  })
})

describe('names, tables, 3-D references', () => {
  const TABLE =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B4" totalsRowShown="0">' +
    '<autoFilter ref="A1:B4"/><tableColumns count="2"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Qty"/></tableColumns></table>'
  const spec = () =>
    book({
      sheets: [
        { name: 'S', rows: rows([['Item', 'Qty'], ['a', 2], ['b', 3], ['c', 5]]), tail: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' },
        { name: 'T', rows: rows([[10]]) },
        { name: 'U', rows: rows([[20]]) },
      ],
      names: '<definedName name="Qty">S!$B$2:$B$4</definedName><definedName name="Rate">0.5</definedName><definedName name="Twice">_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)</definedName><definedName name="Local" localSheetId="1">T!$A$1</definedName>',
      parts: {
        'xl/tables/table1.xml': TABLE,
        'xl/worksheets/_rels/sheet1.xml.rels':
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
      },
      types: '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
    })

  test('defined names', async () => {
    const o = await openBook(await spec())
    await o.set(0, 'D1', '=SUM(Qty)')
    await o.set(0, 'D2', '=MAX(Qty)*Rate')
    await o.set(0, 'D3', '=Twice(21)')
    await o.set(0, 'D4', '=Nope+1')
    expect(await shown(o, ['D1', 'D2', 'D3', 'D4'])).toEqual(['10', '2.5', '42', '#NAME?'])
    // in a cell of the table: an optional parameter of LAMBDA is not a column of it
    await o.set(0, 'A3', '=LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(5)')
    expect(await shown(o, ['A3'])).toEqual(['5'])
    expect(cellXml(await (await o.save()).text(SHEET), 'A3')).toContain('_xlfn.LAMBDA(_xlpm.a,[_xlpm.b],')
    await o.set(0, 'A3', 'b')
    // a change in the cells a name stands for computes what uses the name
    await o.set(0, 'B3', '30')
    expect(await shown(o, ['D1', 'D2'])).toEqual(['37', '15'])
    // the name follows the rows
    await o.engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 2, n: 1 }])
    await o.set(0, 'B3', '100')
    expect(await shown(o, ['D1'])).toEqual(['137'])
    await o.engine.call('xlsx.undo', [])
    await o.engine.call('xlsx.undo', [])
    expect(await shown(o, ['D1'])).toEqual(['37'])
    // a name local to a sheet is known there only
    await o.load(1)
    await o.set(1, 'B1', '=Local*2')
    await o.set(0, 'D5', '=Local*2')
    await o.set(0, 'D6', '=T!Local*2')
    expect(await o.shown(1, 0, 1)).toBe('20')
    expect(await shown(o, ['D5', 'D6'])).toEqual(['#NAME?', '20'])
  })

  test('a name local to a sheet, in a workbook that lists a chart sheet before it', async () => {
    const o = await openBook(
      await book({
        sheets: [{ name: 'Chart', kind: 'chartsheet' }, { name: 'S', rows: rows([[1]]) }, { name: 'T', rows: rows([[10]]) }],
        names: '<definedName name="Loc" localSheetId="2">T!$A$1</definedName>',
      }),
    )
    await o.load(1)
    await o.set(1, 'B1', '=Loc*2')
    await o.set(0, 'B1', '=Loc*2')
    expect(await o.shown(1, 0, 1)).toBe('20')
    expect(await o.shown(0, 0, 1)).toBe('#NAME?')
  })

  test('structured references', async () => {
    const o = await openBook(await spec())
    await o.set(0, 'D1', '=SUM(Sales[Qty])')
    await o.set(0, 'D2', '=Sales[@Qty]*10')
    // (without a table name: of the table the formula is in)
    await o.set(0, 'A3', '=[@Qty]*10')
    await o.set(0, 'C3', '=[@Qty]*10')
    await o.set(0, 'D4', '=COUNTA(Sales[#All])')
    await o.set(0, 'E1', '=Sales[[#Headers],[Item]]')
    await o.set(0, 'E2', '=ROWS(Sales)')
    expect(await shown(o, ['D1', 'D2', 'A3', 'C3', 'D4', 'E1', 'E2'])).toEqual(['10', '20', '30', '#REF!', '8', 'Item', '3'])
    await o.set(0, 'B2', '7')
    expect(await shown(o, ['D1', 'D2'])).toEqual(['15', '70'])
    // a file writes the row of the formula out, with the table's name
    const xml = await (await o.save()).text(SHEET)
    expect(cellXml(xml, 'D2')).toContain('Sales[[#This Row],[Qty]]*10')
    expect(cellXml(xml, 'A3')).toContain('Sales[[#This Row],[Qty]]*10')
  })

  test('3-D references', async () => {
    const o = await openBook(await spec())
    await o.load(1)
    await o.load(2)
    await o.set(0, 'D1', '=SUM(T:U!A1)')
    await o.set(0, 'D2', "=COUNT('S:U'!A1:B4)")
    expect(await shown(o, ['D1', 'D2'])).toEqual(['30', '5'])
    await o.set(2, 'A1', '25')
    expect(await shown(o, ['D1'])).toEqual(['35'])
  })

  test('SUBTOTAL leaves out the rows the filter hides', async () => {
    const hidden = `<row r="1"><c r="A1"><v>1</v></c></row><row r="2" hidden="1"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>4</v></c></row>`
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: hidden, tail: '<autoFilter ref="A1:A3"/>' }] }))
    await o.set(0, 'C1', '=SUBTOTAL(9,A1:A3)')
    await o.set(0, 'C2', '=SUBTOTAL(109,A1:A3)')
    await o.set(0, 'C3', '=SUM(A1:A3)')
    await o.set(0, 'C4', '=SUBTOTAL(9,C1:C3)')
    expect(await shown(o, ['C1', 'C2', 'C3', 'C4'])).toEqual(['5', '5', '7', '7'])
  })
})
