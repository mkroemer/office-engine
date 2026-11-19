/**
 * Regressions from the QA pass of 18/11/2026 on v0.14.0 (ids as in its report: C1…C21 critical, P1-/P2-/N1- per tester).
 */
import { describe, expect, test } from 'bun:test'
import { at, book, openBook, rows } from './xlsx-builder'

const cellsOf = (xml: string, r: number) => [...(new RegExp(`<row r="${r}"[^>]*>(.*?)</row>`).exec(xml)?.[1] ?? '').matchAll(/<c r="([A-Z]+\d+)"/g)].map((m) => m[1])

describe('xlsx: the saved file is a valid workbook', () => {
  test('C1: cells typed into a row that has some are written in column order', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1], [1, null, null, null, 5]]) }] }))
    await o.engine.call('xlsx.set', [0, [{ ...at('B1'), input: 'b' }, { ...at('C1'), input: 'c' }, { ...at('D1'), input: 'd' }]])
    await o.set(0, 'C2', 'c')
    await o.set(0, 'B2', 'b')
    await o.set(0, 'G2', 'g')
    await o.set(0, 'F2', 'f')
    const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect(cellsOf(xml, 1)).toEqual(['A1', 'B1', 'C1', 'D1'])
    expect(cellsOf(xml, 2)).toEqual(['A2', 'B2', 'C2', 'E2', 'F2', 'G2'])
  })

  test('C2: the values of the cells an edit recomputes are saved', async () => {
    const src = await book({ sheets: [{ name: 'S', rows: rows([[1], [2], [{ f: 'SUM(A1:A2)', v: 3 }], [{ f: 'A3*2', v: 6 }], [{ f: 'A1&"x"', v: '1x' }]]) }] })
    const o = await openBook(src)
    await o.set(0, 'A1', '10')
    expect(await o.shown(0, 2, 0)).toBe('12')
    const saved = await o.save()
    const xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<c r="A3"><f>SUM(A1:A2)</f><v>12</v></c>')
    expect(xml).toContain('<f>A3*2</f><v>24</v>')
    expect(xml).toContain('<f>A1&amp;&quot;x&quot;</f><v>10x</v>')
    const re = await openBook(saved.blob)
    expect(await re.shown(0, 2, 0)).toBe('12')
    expect(await re.shown(0, 3, 0)).toBe('24')
    expect(await re.shown(0, 4, 0)).toBe('10x')
  })

  test('C2: so are those of another sheet, and of a shared formula', async () => {
    const src = await book({
      sheets: [
        { name: 'In', rows: rows([[1, 2]]) },
        { name: 'Out', rows: rows([[{ f: 'In!A1+1', v: 2, shared: { si: 0, ref: 'A1:B1' } }, { f: '', v: 3, shared: { si: 0 } }]]) },
      ],
    })
    const o = await openBook(src)
    await o.set(0, 'A1', '10')
    await o.set(0, 'B1', '20')
    const saved = await o.save()
    const xml = await saved.text('xl/worksheets/sheet2.xml')
    expect(xml).toContain('<v>11</v>')
    expect(xml).toContain('<v>21</v>')
    // the shared formula is still one
    expect(xml).toContain('t="shared"')
    const re = await openBook(saved.blob, 1)
    expect(await re.shown(1, 0, 0)).toBe('11')
    expect(await re.shown(1, 0, 1)).toBe('21')
    expect(await re.formula(1, 0, 1)).toBe('=In!B1+1')
  })

  test('C3: adding, deleting and moving sheets keeps localSheetId and activeTab on their sheets', async () => {
    const names = '<definedName name="_xlnm.Print_Area" localSheetId="2">C!$A$1:$B$2</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">A!$A$1</definedName><definedName name="G">B!$A$1</definedName>'
    const mk = () => book({ sheets: ['A', 'B', 'C'].map((name) => ({ name, rows: rows([[1]]) })), names, view: 'activeTab="2" firstSheet="2"' })
    const local = (xml: string) => [...xml.matchAll(/<definedName name="_xlnm.Print_Area" localSheetId="(\d+)">(\w+)!/g)].map((m) => `${m[2]}=${m[1]}`).sort()
    const order = (xml: string) => [...xml.matchAll(/<sheet name="([^"]+)"/g)].map((m) => m[1])

    let o = await openBook(await mk())
    await o.engine.call('xlsx.bookOp', [{ op: 'delete', index: 0 }])
    let xml = await (await o.save()).text('xl/workbook.xml')
    expect(order(xml)).toEqual(['B', 'C'])
    expect(local(xml)).toEqual(['C=1'])
    expect(xml).toContain('activeTab="1"')
    expect(xml).toContain('firstSheet="1"')

    o = await openBook(await mk())
    await o.engine.call('xlsx.bookOp', [{ op: 'add', after: 0, name: 'N' }])
    xml = await (await o.save()).text('xl/workbook.xml')
    expect(order(xml)).toEqual(['A', 'N', 'B', 'C'])
    expect(local(xml)).toEqual(['A=0', 'C=3'])
    expect(xml).toContain('activeTab="3"')

    o = await openBook(await mk())
    await o.engine.call('xlsx.bookOp', [{ op: 'move', from: 2, to: 0 }])
    xml = await (await o.save()).text('xl/workbook.xml')
    expect(order(xml)).toEqual(['C', 'A', 'B'])
    expect(local(xml)).toEqual(['A=1', 'C=0'])
    expect(xml).toContain('activeTab="0"')

    // the last sheet deleted while it is the active one: the tab shown is one that exists
    o = await openBook(await mk())
    await o.engine.call('xlsx.bookOp', [{ op: 'delete', index: 2 }])
    xml = await (await o.save()).text('xl/workbook.xml')
    expect(local(xml)).toEqual(['A=0'])
    expect(Number(/activeTab="(\d+)"/.exec(xml)?.[1] ?? 0)).toBeLessThan(2)
    expect(Number(/firstSheet="(\d+)"/.exec(xml)?.[1] ?? 0)).toBeLessThan(2)
  })

  test('C4: a chart sheet in the list: renaming, deleting and moving act on the sheet meant', async () => {
    const mk = () =>
      book({
        sheets: [{ name: 'Chart', kind: 'chartsheet' }, { name: 'One', rows: rows([[1]]) }, { name: 'Dlg', kind: 'dialogsheet' }, { name: 'Two', rows: rows([[2]]) }, { name: 'Three', rows: rows([[3]]) }],
        names: '<definedName name="_xlnm.Print_Area" localSheetId="3">Two!$A$1</definedName>',
      })
    const order = (xml: string) => [...xml.matchAll(/<sheet name="([^"]+)"[^>]*r:id="(\w+)"/g)].map((m) => `${m[1]}:${m[2]}`)

    let o = await openBook(await mk())
    expect(o.info.xlsx.sheets.map((s: any) => s.name)).toEqual(['One', 'Two', 'Three'])
    await o.engine.call('xlsx.bookOp', [{ op: 'rename', index: 0, name: 'Uno' }])
    let saved = await o.save()
    expect(order(await saved.text('xl/workbook.xml'))).toEqual(['Chart:rId1', 'Uno:rId2', 'Dlg:rId3', 'Two:rId4', 'Three:rId5'])

    o = await openBook(await mk())
    await o.engine.call('xlsx.bookOp', [{ op: 'delete', index: 1 }])
    saved = await o.save()
    let xml = await saved.text('xl/workbook.xml')
    expect(order(xml)).toEqual(['Chart:rId1', 'One:rId2', 'Dlg:rId3', 'Three:rId5'])
    expect(saved.has('xl/worksheets/sheet4.xml')).toBe(false)
    expect(saved.has('xl/chartsheets/sheet1.xml')).toBe(true)
    expect(saved.has('xl/dialogsheets/sheet3.xml')).toBe(true)
    expect(xml).not.toContain('Print_Area')
    const rels = await saved.text('xl/_rels/workbook.xml.rels')
    expect(rels).toContain('rId1"')
    expect(rels).not.toContain('worksheets/sheet4.xml')

    // moving the 3rd worksheet before the 1st: the chart and dialog sheets stay where they are among the others
    o = await openBook(await mk())
    await o.engine.call('xlsx.bookOp', [{ op: 'move', from: 2, to: 0 }])
    xml = await (await o.save()).text('xl/workbook.xml')
    expect(order(xml).filter((x) => !/Chart|Dlg/.test(x))).toEqual(['Three:rId5', 'One:rId2', 'Two:rId4'])
    expect(/localSheetId="(\d+)">Two/.exec(xml)?.[1]).toBe(String(order(xml).findIndex((x) => x.startsWith('Two'))))

    o = await openBook(await mk())
    await o.engine.call('xlsx.bookOp', [{ op: 'add', after: 0, name: 'New' }])
    xml = await (await o.save()).text('xl/workbook.xml')
    expect(order(xml).map((x) => x.split(':')[0]).filter((x) => !/Chart|Dlg/.test(x))).toEqual(['One', 'New', 'Two', 'Three'])
    expect(/localSheetId="(\d+)">Two/.exec(xml)?.[1]).toBe(String(order(xml).findIndex((x) => x.startsWith('Two'))))
  })

  test('C5: merged cells of a filtered sheet are written after the filter, not inside it', async () => {
    const tail = '<autoFilter ref="A1:B5"><filterColumn colId="0"><filters><filter val="x"/></filters></filterColumn><sortState ref="A2:B5"><sortCondition ref="A2:A5"/></sortState></autoFilter>'
    for (const t of [tail, '<sheetProtection sheet="0"/><autoFilter ref="A1:B5"/><sortState ref="A2:B5"><sortCondition descending="1" ref="A2:A5"/></sortState><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>']) {
      const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([['h', 'k'], ['x', 1]]), tail: t }] }))
      await o.engine.call('xlsx.sheetOp', [0, { op: 'merge', range: { r0: 6, c0: 0, r1: 6, c1: 2 } }])
      const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
      const merge = '<mergeCells count="1"><mergeCell ref="A7:C7"/></mergeCells>'
      expect(xml).toContain(merge)
      const before = xml.slice(0, xml.indexOf(merge))
      const count = (s: string, re: RegExp) => (s.match(re) ?? []).length
      // every filter / sort element opened before the merges is closed before them
      expect(count(before, /<autoFilter\b[^>]*[^/]>/g)).toBe(count(before, /<\/autoFilter>/g))
      expect(count(before, /<sortState\b[^>]*[^/]>/g)).toBe(count(before, /<\/sortState>/g))
      expect(before).toContain('sortCondition')
      expect(xml.indexOf('<pageMargins') < 0 || xml.indexOf('<pageMargins') > xml.indexOf(merge)).toBe(true)
    }
  })

  test('C7: a number too large to be one is kept as the text typed', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'A1', '1e999')
    await o.set(0, 'A2', '-1e999')
    await o.set(0, 'A3', '=A1*2')
    await o.set(0, 'A4', '=1e308*10')
    const saved = await o.save()
    const xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).not.toContain('Infinity')
    expect(xml).not.toContain('NaN')
    expect(await o.shown(0, 0, 0)).toBe('1e999')
    expect(await o.shown(0, 3, 0)).toBe('#NUM!')
  })

  test('C8: rows and columns are not inserted when that pushes cells off the sheet', async () => {
    const src = await book({ sheets: [{ name: 'S', rows: `<row r="1"><c r="A1"><v>1</v></c><c r="XFD1"><v>2</v></c></row><row r="1048576"><c r="A1048576"><v>3</v></c></row>` }] })
    const o = await openBook(src)
    await expect(o.engine.call('xlsx.sheetOp', [0, { op: 'insertCols', at: 1, n: 1 }])).rejects.toThrow(/last column/i)
    await expect(o.engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 1, n: 2 }])).rejects.toThrow(/last row/i)
    // nothing moved, and the workbook still saves
    expect(await o.shown(0, 0, 16383)).toBe('2')
    expect(await o.shown(0, 1048575, 0)).toBe('3')
    const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect(xml).not.toContain('XFE')
    // cells that only hold a format at the edge fall off without stopping anything
    const o2 = await openBook(await book({ sheets: [{ name: 'S', rows: `<row r="1"><c r="A1"><v>1</v></c><c r="XFD1" s="0"/></row>` }] }))
    await o2.engine.call('xlsx.sheetOp', [0, { op: 'insertCols', at: 0, n: 1 }])
    const x2 = await (await o2.save()).text('xl/worksheets/sheet1.xml')
    expect(x2).toContain('<c r="B1"><v>1</v></c>')
    expect(x2).not.toContain('XFE')
  })
})

describe('xlsx: structure edits and the other sheets', () => {
  test('C9: undoing an insert or a delete puts back the formulas of the other sheets', async () => {
    const src = await book({ sheets: [{ name: 'A', rows: rows([[1], [2], [3], [4], [5]]) }, { name: 'B', rows: rows([[{ f: 'A!A5', v: 5 }, { f: 'SUM(A!A1:A5)', v: 15 }]]) }] })
    const o = await openBook(src)
    await o.load(1)
    await o.engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 0, n: 2 }])
    expect(await o.formula(1, 0, 0)).toBe('=A!A7')
    await o.engine.call('xlsx.undo', [])
    expect(await o.formula(1, 0, 0)).toBe('=A!A5')
    expect(await o.formula(1, 0, 1)).toBe('=SUM(A!A1:A5)')
    expect(await o.shown(1, 0, 0)).toBe('5')
    await o.engine.call('xlsx.redo', [])
    expect(await o.formula(1, 0, 0)).toBe('=A!A7')
    expect(await o.formula(1, 0, 1)).toBe('=SUM(A!A3:A7)')
    // a delete that takes the cell referred to: #REF!, and back
    await o.engine.call('xlsx.undo', [])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'deleteRows', at: 4, n: 1 }])
    expect(await o.formula(1, 0, 0)).toBe('=A!#REF!')
    await o.engine.call('xlsx.undo', [])
    expect(await o.formula(1, 0, 0)).toBe('=A!A5')
    expect(await o.shown(1, 0, 0)).toBe('5')
    expect(await o.shown(1, 0, 1)).toBe('15')
    const xml = await (await o.save()).text('xl/worksheets/sheet2.xml')
    expect(xml).toContain('<f>A!A5</f>')
    expect(xml).toContain('<f>SUM(A!A1:A5)</f>')
  })

  test('C12: formulas across sheets follow their sheets when one is added, moved, renamed or deleted after an edit', async () => {
    const mk = () =>
      book({
        sheets: [
          { name: 'In', rows: rows([[1]]) },
          { name: 'Two', rows: rows([[{ f: 'In!A1*2', v: 2 }]]) },
          { name: 'Three', rows: rows([[{ f: 'In!A1*3', v: 3 }, { f: 'Two!A1+1', v: 3 }]]) },
        ],
      })
    const open = async () => {
      const o = await openBook(await mk())
      await o.load(1)
      await o.load(2)
      await o.set(0, 'A1', '5') // the dependency graph exists from here on
      expect(await o.shown(1, 0, 0)).toBe('10')
      return o
    }
    // a sheet added in the middle
    let o = await open()
    await o.engine.call('xlsx.bookOp', [{ op: 'add', after: 0 }])
    await o.set(0, 'A1', '7')
    expect(await o.shown(2, 0, 0)).toBe('14')
    expect(await o.shown(3, 0, 0)).toBe('21')
    expect(await o.shown(3, 0, 1)).toBe('15')
    expect(await o.shown(1, 0, 0)).toBe('')
    let re = await openBook((await o.save()).blob)
    await re.load(2)
    await re.load(3)
    expect(await re.shown(2, 0, 0)).toBe('14')
    expect(await re.shown(3, 0, 1)).toBe('15')

    // moved
    o = await open()
    await o.engine.call('xlsx.bookOp', [{ op: 'move', from: 0, to: 2 }])
    expect(o.info.xlsx.sheets.length).toBe(3)
    await o.set(2, 'A1', '7')
    expect(await o.shown(0, 0, 0)).toBe('14')
    expect(await o.shown(1, 0, 0)).toBe('21')
    expect(await o.shown(1, 0, 1)).toBe('15')

    // renamed: the formulas name the sheet as it is called now, and still compute
    o = await open()
    await o.engine.call('xlsx.bookOp', [{ op: 'rename', index: 0, name: 'Data 1' }])
    expect(await o.formula(1, 0, 0)).toBe("='Data 1'!A1*2")
    await o.set(0, 'A1', '7')
    expect(await o.shown(1, 0, 0)).toBe('14')
    expect(await o.shown(2, 0, 0)).toBe('21')
    re = await openBook((await o.save()).blob, 1)
    expect(await re.shown(1, 0, 0)).toBe('14')
    expect(await re.formula(1, 0, 0)).toBe("='Data 1'!A1*2")

    // deleted: what read it is #REF! (now, and in the file), the others go on
    o = await open()
    await o.engine.call('xlsx.bookOp', [{ op: 'delete', index: 1 }])
    expect(await o.shown(1, 0, 1)).toBe('#REF!')
    await o.set(0, 'A1', '7')
    expect(await o.shown(1, 0, 0)).toBe('21')
    expect(await o.shown(1, 0, 1)).toBe('#REF!')
    re = await openBook((await o.save()).blob, 1)
    expect(await re.shown(1, 0, 0)).toBe('21')
    expect(await re.shown(1, 0, 1)).toBe('#REF!')
  })

  test('C13: a formula reading a sheet not shown yet computes on its cells', async () => {
    const src = await book({
      sheets: [
        { name: 'Ask', rows: rows([[2, { f: 'VLOOKUP(A1,Table!A1:B3,2,FALSE)', v: 'two' }, { f: 'SUM(Table!A:A)', v: 6 }]]) },
        { name: 'Table', rows: rows([[1, 'one'], [2, 'two'], [3, 'three']]) },
        { name: 'Far', rows: rows([[{ f: 'Ask!A1*10', v: 20 }]]) },
      ],
    })
    const o = await openBook(src)
    await o.set(0, 'A1', '3')
    expect(await o.shown(0, 0, 1)).toBe('three')
    expect(await o.shown(0, 0, 2)).toBe('6')
    // and a sheet never shown that reads the edited cell is saved with its new value
    const saved = await o.save()
    expect(await saved.text('xl/worksheets/sheet3.xml')).toContain('<v>30</v>')
    const re = await openBook(saved.blob, 2)
    expect(await re.shown(2, 0, 0)).toBe('30')
  })

  test('C14: a factorial too large is #NUM! at once', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    const t0 = performance.now()
    await o.set(0, 'A2', '=FACT(1E15)')
    await o.set(0, 'A3', '=FACT(170)')
    await o.set(0, 'A4', '=FACT(171)')
    await o.set(0, 'A5', '=REPT("x",1E9)')
    expect(performance.now() - t0).toBeLessThan(1000)
    expect(await o.shown(0, 1, 0)).toBe('#NUM!')
    expect(await o.shown(0, 2, 0)).toMatch(/^7\.257/)
    expect(await o.shown(0, 3, 0)).toBe('#NUM!')
    expect(await o.shown(0, 4, 0)).toBe('#VALUE!')
  })
})

describe('xlsx: formulas that were not computed', () => {
  test('P1-004: the cells of a shared formula compute, show their formula, and keep it when formatted or typed over and undone', async () => {
    const mk = () =>
      book({
        sheets: [{ name: 'S', rows: rows([[1, { f: 'A1*2', v: 2, shared: { si: 0, ref: 'B1:B3' } }], [2, { f: '', v: 4, shared: { si: 0 } }], [3, { f: '', v: 6, shared: { si: 0 } }]]) }],
        styles: '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>',
      })
    let o = await openBook(await mk())
    expect(await o.formula(0, 2, 1)).toBe('=A3*2')
    await o.set(0, 'A3', '30')
    expect(await o.shown(0, 2, 1)).toBe('60')
    // formatted: still formulas, in the file too
    await o.engine.call('xlsx.format', [0, { r0: 0, c0: 1, r1: 2, c1: 1 }, { b: true }])
    let re = await openBook((await o.save()).blob)
    expect(await re.formula(0, 1, 1)).toBe('=A2*2')
    expect(await re.formula(0, 2, 1)).toBe('=A3*2')
    expect(await re.shown(0, 2, 1)).toBe('60')
    // typed over, undone: the formula is back
    o = await openBook(await mk())
    await o.set(0, 'B2', '99')
    await o.engine.call('xlsx.undo', [])
    expect(await o.formula(0, 1, 1)).toBe('=A2*2')
    await o.set(0, 'A2', '5')
    expect(await o.shown(0, 1, 1)).toBe('10')
    // the first cell typed over: the others keep theirs
    await o.set(0, 'B1', 'x')
    re = await openBook((await o.save()).blob)
    expect(await re.formula(0, 0, 1)).toBeUndefined()
    expect(await re.formula(0, 1, 1)).toBe('=A2*2')
    expect(await re.formula(0, 2, 1)).toBe('=A3*2')
    // filled over the first cell: same
    o = await openBook(await mk())
    await o.set(0, 'C1', '=A1+100')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'fill', range: { r0: 0, c0: 2, r1: 0, c1: 2 }, dir: 'right' }])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range: { r0: 0, c0: 0, r1: 2, c1: 0 }, col: 0, desc: true }])
    re = await openBook((await o.save()).blob)
    expect(await re.formula(0, 2, 1)).toBe('=A3*2')
    expect(await re.shown(0, 2, 1)).toBe('2')
  })

  test('P1-004: an array formula fills its cells again, what reads them follows, and it stays one when rows move', async () => {
    const src = await book({
      sheets: [{ name: 'S', rows: `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="B1:B3">A1:A3*10</f><v>10</v></c><c r="C1"><f>B3+1</f><v>31</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>20</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>30</v></c></row>` }],
    })
    const o = await openBook(src)
    await o.set(0, 'A3', '7')
    expect(await o.shown(0, 2, 1)).toBe('70')
    expect(await o.shown(0, 0, 2)).toBe('71')
    let saved = await o.save()
    let xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<f t="array" ref="B1:B3">A1:A3*10</f><v>10</v>')
    expect(xml).toContain('<c r="B3"><v>70</v></c>')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 0, n: 1 }])
    saved = await o.save()
    xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<f t="array" ref="B2:B4">A2:A4*10</f>')
    const re = await openBook(saved.blob)
    await re.set(0, 'A2', '5')
    expect(await re.shown(0, 1, 1)).toBe('50')
    expect(await re.shown(0, 3, 1)).toBe('70')
  })

  test('P1-003: a workbook that says to compute on load is computed, every sheet in the order its formulas need', async () => {
    const src = await book({
      sheets: [
        { name: 'A', rows: rows([[{ f: 'B!A1+1', v: 0 }, 5]]) },
        { name: 'B', rows: rows([[{ f: 'A!B1*2', v: 0 }, { f: 'PMT(1,2,3)', v: 42 }]]) },
      ],
    })
    const parts = await (await import('./xlsx-builder')).unzip(src)
    const { rebuildZip } = await import('../src/core/zipwrite')
    const { ZipArchive } = await import('../src/core/zip')
    const { BlobSource } = await import('../src/core/source')
    const wbXml = (await parts.text('xl/workbook.xml')).replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
    const full = await rebuildZip(await ZipArchive.open(new BlobSource(src)), new Map([['xl/workbook.xml', wbXml]]))
    const o = await openBook(full)
    while (!o.events.some((e) => e.event === 'xlsx:progress' && e.data.recalc)) await new Promise((r) => setTimeout(r, 5))
    expect(await o.shown(0, 0, 0)).toBe('11')
    await o.load(1)
    expect(await o.shown(1, 0, 0)).toBe('10')
    // a function the engine does not have: the value of the file stays
    expect(await o.shown(1, 0, 1)).toBe('42')
    // nothing was edited: computing is not a change to save
    expect(await o.engine.call('dirty', [])).toBe(false)
  })

  test('P1-005: INDIRECT and OFFSET follow the cells they end up reading', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1, 'A1', { f: 'INDIRECT(B1)', v: 1 }, { f: 'SUM(OFFSET(A1,0,0,2,1))', v: 3 }, { f: 'C1+1', v: 2 }], [2]]) }] }))
    await o.set(0, 'A1', '10')
    expect(await o.shown(0, 0, 2)).toBe('10')
    expect(await o.shown(0, 0, 3)).toBe('12')
    expect(await o.shown(0, 0, 4)).toBe('11')
    await o.set(0, 'A2', '5')
    expect(await o.shown(0, 0, 3)).toBe('15')
    await o.set(0, 'B1', 'A2')
    expect(await o.shown(0, 0, 2)).toBe('5')
    expect(await o.shown(0, 0, 4)).toBe('6')
  })

  test('P1-012: a long chain of formulas computes (each row reading the one above)', async () => {
    const N = 20_000
    const grid: any[][] = [[1]]
    for (let r = 1; r < N; r++) grid.push([{ f: `A${r}+1`, v: r + 1 }])
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows(grid) }] }))
    await o.set(0, 'A1', '100')
    expect(await o.shown(0, 1, 0)).toBe('101')
    expect(await o.shown(0, 700, 0)).toBe('800')
    expect(await o.shown(0, N - 1, 0)).toBe(String(100 + N - 1))
    // a circular reference longer than the stack is deep: an error, not a hang
    await o.set(0, 'A1', `=A${N}`)
    expect(await o.shown(0, N - 1, 0)).toBe('#REF!')
    expect(await o.shown(0, 0, 0)).toBe('#REF!')
    await o.set(0, 'A1', '1')
    expect(await o.shown(0, N - 1, 0)).toBe(String(N))
  }, 30000)

  test('N1-027 / P2-008 / P1-014: what the engine cannot compute keeps the value of the file', async () => {
    const src = await book({
      sheets: [{ name: 'S', rows: rows([[1000, { f: 'PMT(0.05/12,60,A1)', v: -18.87 }, { f: 'IFERROR(PMT(0.05/12,60,A1),0)', v: -18.87 }, { f: 'Rate*A1', v: 50 }, { f: 'B1*2', v: -37.74 }, { f: 'Table1[Total]', v: 7 }]]) }],
      names: '<definedName name="Rate">0.05</definedName>',
    })
    const o = await openBook(src)
    await o.set(0, 'A1', '2000')
    expect(await o.shown(0, 0, 1)).toBe('-18.87')
    expect(await o.shown(0, 0, 2)).toBe('-18.87')
    expect(await o.shown(0, 0, 3)).toBe('50')
    expect(await o.shown(0, 0, 5)).toBe('7')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 0, n: 1 }])
    expect(await o.shown(0, 1, 1)).toBe('-18.87')
    const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect(xml).not.toContain('#NAME?')
    expect(xml).toContain('<v>-18.87</v>')
    // typed now, there is no value to keep: it says what is wrong
    await o.set(0, 'H1', '=NOSUCHFN(1)')
    expect(await o.shown(0, 0, 7)).toBe('#NAME?')
    await o.set(0, 'H2', '=IFERROR(NOSUCHFN(1),0)')
    expect(await o.shown(0, 1, 7)).toBe('#NAME?')
  })

  test('P1-007: functions newer than the format are saved with their prefix, and shown without', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1, 2], [{ f: '_xlfn.XLOOKUP(1,A1:A1,B1:B1)', v: 2 }]]) }] }))
    await o.set(0, 'C1', '=XLOOKUP(1,A1:A1,B1:B1)')
    await o.set(0, 'C2', '=IFS(A1>0,"p",TRUE,"n")&TEXTJOIN(",",TRUE,A1:B1)&SUM(A1:B1)')
    await o.set(0, 'C3', '=_xlfn.CONCAT(A1,B1)')
    expect(await o.formula(0, 1, 0)).toBe('=XLOOKUP(1,A1:A1,B1:B1)')
    const saved = await o.save()
    const xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<f>_xlfn.XLOOKUP(1,A1:A1,B1:B1)</f>')
    expect(xml).toContain('<f>_xlfn.IFS(A1&gt;0,"p",TRUE,"n")&amp;_xlfn.TEXTJOIN(",",TRUE,A1:B1)&amp;SUM(A1:B1)</f>')
    expect(xml).toContain('<f>_xlfn.CONCAT(A1,B1)</f>')
    expect(xml).not.toContain('_xlfn._xlfn')
    const re = await openBook(saved.blob)
    expect(await re.formula(0, 0, 2)).toBe('=XLOOKUP(1,A1:A1,B1:B1)')
    expect(await re.shown(0, 0, 2)).toBe('2')
  })
})

describe('xlsx: sheets and names', () => {
  test('P2-017 / P2-011: a sheet name takes no apostrophe at its ends; renaming follows into the defined names', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'A', rows: rows([[1]]) }, { name: 'B', rows: rows([[{ f: 'A!A1', v: 1 }]]) }], names: '<definedName name="_xlnm.Print_Area" localSheetId="0">A!$A$1:$B$2</definedName><definedName name="Top">A!$A$1</definedName>' }))
    await o.engine.call('xlsx.bookOp', [{ op: 'rename', index: 0, name: "Q1 'draft'" }])
    const xml = await (await o.save()).text('xl/workbook.xml')
    const name = /<sheet name="([^"]+)"/.exec(xml)![1].replace(/&apos;/g, "'")
    expect(name.startsWith("'") || name.endsWith("'")).toBe(false)
    expect(name).toBe("Q1 'draft_")
    expect(xml).toContain(`localSheetId="0">'Q1 ''draft_'!$A$1:$B$2</definedName>`)
    expect(xml).toContain(`<definedName name="Top">'Q1 ''draft_'!$A$1</definedName>`)
    await o.load(1)
    expect(await o.formula(1, 0, 0)).toBe("='Q1 ''draft_'!A1")
  })

  test('C6: a workbook with macros is saved as one: its type, its extension, its macros', async () => {
    const src = await book({
      sheets: [{ name: 'S', rows: rows([[1]]) }],
      mainType: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      parts: { 'xl/vbaProject.bin': new Uint8Array([1, 2, 3, 4]) },
      types: '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>',
    })
    const o = await openBook(src)
    expect(o.info.ext).toBe('xlsm')
    await o.set(0, 'A1', '2')
    const saved = await o.save()
    expect(saved.blob.type.toLowerCase()).toBe('application/vnd.ms-excel.sheet.macroenabled.12')
    expect(saved.has('xl/vbaProject.bin')).toBe(true)
    expect(await saved.text('[Content_Types].xml')).toContain('macroEnabled.main+xml')
    // a plain workbook says nothing of the kind
    const plain = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    expect(plain.info.ext).toBeUndefined()
    await plain.set(0, 'A1', '2')
    expect((await plain.save()).blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  })
})

describe('xlsx: cells cleared, copied, cut and pasted in the worker', () => {
  const big = (n: number) => rows(Array.from({ length: n }, (_, r) => [r + 1, `t${r}`, { f: `A${r + 1}*2`, v: (r + 1) * 2 }]))

  test('C10: clearing a whole column clears every row of it, not the ones shown', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: big(3000) }] }))
    await o.engine.call('xlsx.sheetOp', [0, { op: 'clear', range: { r0: 0, c0: 0, r1: 1_048_575, c1: 0 } }])
    expect(await o.shown(0, 0, 0)).toBe('')
    expect(await o.shown(0, 2999, 0)).toBe('')
    expect(await o.shown(0, 2999, 1)).toBe('t2999')
    expect(await o.shown(0, 2999, 2)).toBe('0')
    await o.engine.call('xlsx.undo', [])
    expect(await o.shown(0, 2999, 0)).toBe('3000')
    expect(await o.shown(0, 2999, 2)).toBe('6000')
    // the whole sheet
    await o.engine.call('xlsx.sheetOp', [0, { op: 'clear', range: { r0: 0, c0: 0, r1: 1_048_575, c1: 16_383 } }])
    const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect(xml).not.toContain('<v>')
    expect(xml).not.toContain('<f>')
  })

  test('C10 / N1-004: a whole column copied is all there; pasted here, values are exact and formulas are formulas', async () => {
    const o = await openBook(
      await book({
        sheets: [{ name: 'S', rows: rows([[3.14159, 'a\tb', { f: 'A1*2', v: 6.28318 }], ...Array.from({ length: 2999 }, (_, r) => [r, 'x', { f: `A${r + 2}*2`, v: r * 2 }])]) }, { name: 'T', rows: rows([[1]]) }],
        styles: '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>',
      }),
    )
    await o.engine.call('xlsx.format', [0, { r0: 0, c0: 0, r1: 0, c1: 0 }, { xf: 1 }])
    expect(await o.shown(0, 0, 0)).toBe('3.14')
    const clip: any = await o.engine.call('xlsx.copy', [0, { r0: 0, c0: 0, r1: 1_048_575, c1: 2 }])
    expect(clip.cells.length).toBe(3000)
    const lines = clip.text.split('\n')
    expect(lines[0]).toBe('3.14\t"a\tb"\t6.28318')
    expect(lines[2999]).toBe('2998\tx\t5996')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'paste', range: { r0: 0, c0: 5, r1: 2999, c1: 7 }, cells: clip.cells, from: { sheet: clip.sheet, r0: 0, c0: 0 } }])
    expect(((await o.engine.call('xlsx.cell', [0, 0, 5])) as any).value).toBe('3.14159')
    expect(await o.shown(0, 0, 5)).toBe('3.14')
    expect(await o.shown(0, 0, 6)).toBe('a\tb')
    expect(await o.formula(0, 0, 7)).toBe('=F1*2')
    expect(await o.formula(0, 2999, 7)).toBe('=F3000*2')
    expect(await o.shown(0, 2999, 7)).toBe('5996')
    // on another sheet
    await o.load(1)
    await o.engine.call('xlsx.sheetOp', [1, { op: 'paste', range: { r0: 1, c0: 0, r1: 1, c1: 2 }, cells: clip.cells.slice(0, 1), from: { sheet: clip.sheet, r0: 0, c0: 0 } }])
    expect(await o.formula(1, 1, 2)).toBe('=A2*2')
    await o.engine.call('xlsx.undo', [])
    expect(await o.shown(1, 1, 0)).toBe('')
  })

  test('N1-005: cells cut and pasted are moved: what read them reads them where they are', async () => {
    const src = await book({
      sheets: [
        { name: 'S', rows: rows([[1, 2, { f: 'A1+B1', v: 3 }], [{ f: 'SUM(A1:B1)', v: 3 }, { f: 'A1*10', v: 10 }, { f: 'SUM(A1:C1)', v: 6 }]]) },
        { name: 'T', rows: rows([[{ f: 'S!A1+S!C1', v: 4 }]]) },
      ],
    })
    const o = await openBook(src)
    await o.load(1)
    const range = { r0: 0, c0: 0, r1: 0, c1: 1 }
    const clip: any = await o.engine.call('xlsx.copy', [0, range])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'clear', range }])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'paste', range: { r0: 4, c0: 3, r1: 4, c1: 4 }, cells: clip.cells, from: { sheet: clip.sheet, r0: 0, c0: 0 }, moved: true }])
    expect(await o.shown(0, 4, 3)).toBe('1')
    expect(await o.formula(0, 0, 2)).toBe('=D5+E5')
    expect(await o.shown(0, 0, 2)).toBe('3')
    expect(await o.formula(0, 1, 0)).toBe('=SUM(D5:E5)')
    expect(await o.formula(0, 1, 1)).toBe('=D5*10')
    // a range only partly moved stays where it was
    expect(await o.formula(0, 1, 2)).toBe('=SUM(A1:C1)')
    expect(await o.formula(1, 0, 0)).toBe('=S!D5+S!C1')
    expect(await o.shown(1, 0, 0)).toBe('4')
    await o.engine.call('xlsx.undo', [])
    expect(await o.formula(0, 0, 2)).toBe('=A1+B1')
    expect(await o.formula(1, 0, 0)).toBe('=S!A1+S!C1')
    // to another sheet: the formulas name it
    await o.engine.call('xlsx.sheetOp', [1, { op: 'paste', range: { r0: 2, c0: 0, r1: 2, c1: 1 }, cells: clip.cells, from: { sheet: clip.sheet, r0: 0, c0: 0 }, moved: true }])
    expect(await o.formula(0, 0, 2)).toBe('=T!A3+T!B3')
    expect(await o.formula(1, 0, 0)).toBe('=A3+S!C1')
    expect(await o.shown(0, 0, 2)).toBe('3')
  })

  test('N1-003: cells with line breaks, tabs and quotes read from the clipboard as the cells they are', async () => {
    const { parseTsv, tsvCell } = await import('../src/core/xlsx/clip')
    expect(parseTsv('a\tb\r\nc\td\r\n')).toEqual([['a', 'b'], ['c', 'd']])
    expect(parseTsv('"line 1\nline 2"\tx\n"say ""hi"""\t"t\tab"\n')).toEqual([['line 1\nline 2', 'x'], ['say "hi"', 't\tab']])
    expect(parseTsv('"quoted" then more\t5')).toEqual([['"quoted" then more', '5']])
    expect(parseTsv('"never closed\t5\nz')).toEqual([['"never closed', '5'], ['z']])
    expect(parseTsv('\t\n\t')).toEqual([['', ''], ['', '']])
    expect(parseTsv('')).toEqual([['']])
    const cells = ['plain', 'two\nlines', 'a\tb', 'q"q', '']
    expect(parseTsv(cells.map(tsvCell).join('\t'))).toEqual([cells])
  })
})

describe('xlsx: arrays where a function takes one value (P1-006)', () => {
  test('criteria, ranks and looked-up values given as arrays compute per element', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([['x', 5, 'k1', 'one'], ['y', 9, 'k2', 'two'], ['x', 7, 'k3', 'three'], ['z', 1, null, null]]) }] }))
    const calc = async (f: string) => {
      await o.set(0, 'H1', f)
      return o.shown(0, 0, 7)
    }
    // distinct count
    expect(await calc('=SUMPRODUCT(1/COUNTIF(A1:A4,A1:A4))')).toBe('3')
    expect(await calc('=SUM(COUNTIF(A1:A4,{"x","y"}))')).toBe('3')
    expect(await calc('=SUM(SUMIF(A1:A4,{"x","z"},B1:B4))')).toBe('13')
    expect(await calc('=SUM(SUMIFS(B1:B4,A1:A4,{"x","y"}))')).toBe('21')
    expect(await calc('=SUM(LARGE(B1:B4,{1,2,3}))')).toBe('21')
    expect(await calc('=SUM(SMALL(B1:B4,{1,2}))')).toBe('6')
    expect(await calc('=SUM(RANK(B1:B4,B1:B4))')).toBe('10')
    expect(await calc('=SUM(MATCH({"k2","k3"},C1:C3,0))')).toBe('5')
    // looking up to the left
    expect(await calc('=VLOOKUP("two",CHOOSE({1,2},D1:D3,C1:C3),2,FALSE)')).toBe('k2')
    expect(await calc('=SUM(VLOOKUP("x",A1:B4,{2,2},FALSE))')).toBe('10')
    // one value: as before
    expect(await calc('=COUNTIF(A1:A4,"x")')).toBe('2')
    expect(await calc('=CHOOSE(2,"a","b")')).toBe('b')
    expect(await calc('=LARGE(B1:B4,2)')).toBe('7')
    expect(await calc('=VLOOKUP("k3",C1:D3,2,FALSE)')).toBe('three')
  })
})

describe('xlsx: what is typed (N1-006, N1-007, N1-024, P2-015)', () => {
  const STYLES = (fmts: string[]) =>
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${fmts.length}">${fmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${f}"/>`).join('')}</numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${fmts.length + 2}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>${fmts.map((_, i) => `<xf numFmtId="${164 + i}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('')}</cellXfs></styleSheet>`
  const raw = async (o: Awaited<ReturnType<typeof openBook>>, ref: string) => (await o.engine.call('xlsx.cell', [0, at(ref).r, at(ref).c])) as any

  test('a cell formatted as Text keeps what is typed; a date that is none is not turned into another', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[{ s: 1 }, { s: 1 }, { s: 1 }, { s: 2 }, { s: 3 }]]) }], styles: STYLES(['dd/mm/yyyy', '0%']) }))
    await o.set(0, 'A1', '0912345678')
    await o.set(0, 'B1', '001234567890')
    await o.set(0, 'C1', '=1+1')
    expect(await o.shown(0, 0, 0)).toBe('0912345678')
    expect((await raw(o, 'A1')).type).toBe('string')
    expect(await o.shown(0, 0, 1)).toBe('001234567890')
    expect(await o.shown(0, 0, 2)).toBe('=1+1')
    expect((await raw(o, 'C1')).formula).toBeUndefined()
    const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<t xml:space="preserve">0912345678</t>')
    // dates
    await o.set(0, 'D1', '1/15/2024')
    expect(await o.shown(0, 0, 3)).toBe('15/01/2024')
    await o.set(0, 'D1', '31/02/2024')
    expect(await o.shown(0, 0, 3)).toBe('31/02/2024')
    expect((await raw(o, 'D1')).type).toBe('string')
    await o.set(0, 'D1', '15/15/2024')
    expect((await raw(o, 'D1')).type).toBe('string')
    await o.set(0, 'D1', '05/03/2024')
    expect(await o.shown(0, 0, 3)).toBe('05/03/2024')
    expect((await raw(o, 'D1')).value).toBe('45356')
    // percentages
    await o.set(0, 'E1', '50')
    expect(await o.shown(0, 0, 4)).toBe('50%')
    expect((await raw(o, 'E1')).value).toBe('0.5')
    await o.set(0, 'E1', '0.25')
    expect(await o.shown(0, 0, 4)).toBe('25%')
    await o.set(0, 'E1', '12%')
    expect(await o.shown(0, 0, 4)).toBe('12%')
  })

  test('a date or a percentage typed into a cell without a format shows as typed, and is a number', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'B1', '50%')
    expect(await o.shown(0, 0, 1)).toBe('50%')
    expect((await raw(o, 'B1')).value).toBe('0.5')
    await o.set(0, 'B2', '12.5%')
    expect(await o.shown(0, 1, 1)).toBe('12.5%')
    await o.set(0, 'B3', '2024-01-15')
    expect(await o.shown(0, 2, 1)).toBe('2024-01-15')
    expect((await raw(o, 'B3')).value).toBe('45306')
    await o.set(0, 'B4', '15/01/2024 08:30')
    expect(await o.shown(0, 3, 1)).toBe('15/01/2024 08:30')
    await o.set(0, 'B5', '=B3+1')
    await o.set(0, 'B6', 'not/a/date')
    expect((await raw(o, 'B6')).type).toBe('string')
    const re = await openBook((await o.save()).blob)
    expect(await re.shown(0, 0, 1)).toBe('50%')
    expect(await re.shown(0, 2, 1)).toBe('2024-01-15')
    // (a date moved by a day shows as a date: the format of the cell it reads)
    expect(await re.shown(0, 4, 1)).toBe('2024-01-16')
    // undone in one step, format with it
    await o.set(0, 'D1', '75%')
    await o.engine.call('xlsx.undo', [])
    expect(await o.shown(0, 0, 3)).toBe('')
  })

  test('P2-015: in a workbook counting from 1904 a typed date is that date', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[{ s: 2, v: 0 }]]) }], pr: 'date1904="1"', styles: STYLES(['yyyy-mm-dd']) }))
    expect(o.info.xlsx.date1904).toBe(true)
    expect(await o.shown(0, 0, 0)).toBe('1904-01-01')
    await o.set(0, 'A1', '2024-01-15')
    expect(await o.shown(0, 0, 0)).toBe('2024-01-15')
    expect((await raw(o, 'A1')).value).toBe(String(45306 - 1462))
    await o.set(0, 'B1', '15/01/2024')
    expect(await o.shown(0, 0, 1)).toBe('15/01/2024')
  })
})

describe('xlsx: sorting a selection (N1-008, P2-035)', () => {
  test('the table around the selection, and whether its first row is a header', async () => {
    const o = await openBook(
      await book({
        sheets: [
          { name: 'S', rows: rows([['Name', 'Qty', null, 'Other'], ['b', 2, null, 9], ['a', 3, null, 8], ['c', 1, null, 7], [null], ['far', 5]]) },
          { name: 'Text', rows: rows([['x', 'y'], ['b', 'q'], ['a', 'p']]) },
        ],
      }),
    )
    // one cell in the table: the table, not the one next to it past the empty column, nor the rows past the empty row
    expect(await o.engine.call('xlsx.region', [0, { r0: 2, c0: 1, r1: 2, c1: 1 }])).toEqual({ range: { r0: 0, c0: 0, r1: 3, c1: 1 }, header: true })
    // a column of it selected whole: the same rows, whole
    expect(await o.engine.call('xlsx.region', [0, { r0: 0, c0: 0, r1: 1_048_575, c1: 0 }])).toEqual({ range: { r0: 0, c0: 0, r1: 5, c1: 1 }, header: true })
    expect(await o.engine.call('xlsx.region', [0, { r0: 1, c0: 3, r1: 1, c1: 3 }])).toEqual({ range: { r0: 0, c0: 3, r1: 3, c1: 3 }, header: true })
    await o.load(1)
    // text over text, nothing bold: no telling a header
    expect(await o.engine.call('xlsx.region', [1, { r0: 1, c0: 0, r1: 1, c1: 0 }])).toEqual({ range: { r0: 0, c0: 0, r1: 2, c1: 1 }, header: false })
    // sorted as the view asks: the rows under the header, whole
    await o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range: { r0: 1, c0: 0, r1: 3, c1: 1 }, col: 0 }])
    expect([await o.shown(0, 0, 0), await o.shown(0, 1, 0), await o.shown(0, 1, 1), await o.shown(0, 3, 0), await o.shown(0, 3, 1), await o.shown(0, 1, 3)]).toEqual(['Name', 'a', '3', 'c', '1', '9'])
  })
})

describe('xlsx: what points at cells follows them (P2-009, P2-010, P2-011)', () => {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const mk = () =>
    book({
      sheets: [
        {
          name: 'Data',
          rows: rows([['Name', 'Qty', 'Kind'], ['a', 5, 'x'], ['b', 20, 'y'], ['c', 7, 'x'], ['d', 30, 'y']]),
          tail:
            '<autoFilter ref="A1:C5"/><conditionalFormatting sqref="B2:B5"><cfRule type="expression" priority="1"><formula>$B2&gt;10</formula></cfRule></conditionalFormatting>' +
            '<dataValidations count="1"><dataValidation type="list" sqref="C2:C5"><formula1>Other!$A$1:$A$3</formula1></dataValidation></dataValidations>' +
            '<hyperlinks><hyperlink ref="A3" location="Other!A1" display="b"/><hyperlink ref="C5" location="Other!A2"/></hyperlinks>' +
            '<rowBreaks count="1" manualBreakCount="1"><brk id="3" max="16383" man="1"/></rowBreaks><drawing r:id="rId1"/><tableParts count="1"><tablePart r:id="rId2"/></tableParts>',
        },
        { name: 'Other', rows: rows([['x'], ['y'], ['z']]) },
      ],
      names: '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$C$5</definedName><definedName name="Total">Data!$B$2:$B$5</definedName><definedName name="Kinds">Other!$A$1:$A$3</definedName>',
      parts: {
        'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId2" Type="${R}/table" Target="../tables/table1.xml"/></Relationships>`,
        'xl/tables/table1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:C5"><autoFilter ref="A1:C5"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Kind"/></tableColumns></table>',
        'xl/drawings/drawing1.xml':
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to></xdr:twoCellAnchor></xdr:wsDr>',
        'xl/charts/chart1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
      },
      types:
        '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    })

  test('rows inserted and deleted: names, filter, validation, conditional format, links, breaks, table, picture, chart', async () => {
    const o = await openBook(await mk())
    await o.engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 1, n: 2 }])
    let s = await o.save()
    let wb = await s.text('xl/workbook.xml')
    expect(wb).toContain('localSheetId="0">Data!$A$1:$C$7</definedName>')
    expect(wb).toContain('<definedName name="Total">Data!$B$4:$B$7</definedName>')
    expect(wb).toContain('<definedName name="Kinds">Other!$A$1:$A$3</definedName>')
    let sh = await s.text('xl/worksheets/sheet1.xml')
    expect(sh).toContain('<autoFilter ref="A1:C7"/>')
    expect(sh).toContain('<conditionalFormatting sqref="B4:B7"><cfRule type="expression" priority="1"><formula>$B4&gt;10</formula>')
    expect(sh).toContain('<dataValidation type="list" sqref="C4:C7"><formula1>Other!$A$1:$A$3</formula1>')
    expect(sh).toContain('<hyperlink ref="A5" location="Other!A1" display="b"/><hyperlink ref="C7" location="Other!A2"/>')
    expect(sh).toContain('<brk id="5" max="16383" man="1"/>')
    expect(await s.text('xl/tables/table1.xml')).toContain('ref="A1:C7"><autoFilter ref="A1:C7"/>')
    expect(await s.text('xl/drawings/drawing1.xml')).toContain('<xdr:row>8</xdr:row>')
    expect(await s.text('xl/drawings/drawing1.xml')).toContain('<xdr:row>14</xdr:row>')
    expect(await s.text('xl/charts/chart1.xml')).toContain('<c:f>Data!$B$4:$B$7</c:f>')
    // undone: as the file had it
    await o.engine.call('xlsx.undo', [])
    s = await o.save()
    expect(await s.text('xl/workbook.xml')).toContain('<definedName name="Total">Data!$B$2:$B$5</definedName>')
    sh = await s.text('xl/worksheets/sheet1.xml')
    expect(sh).toContain('<autoFilter ref="A1:C5"/>')
    expect(sh).toContain('<hyperlink ref="A3" location="Other!A1" display="b"/>')
    expect(await s.text('xl/tables/table1.xml')).toContain('ref="A1:C5"')
    expect(await s.text('xl/charts/chart1.xml')).toContain('<c:f>Data!$B$2:$B$5</c:f>')
    // rows deleted: what was only on them goes, the rest shrinks
    await o.engine.call('xlsx.sheetOp', [0, { op: 'deleteRows', at: 2, n: 1 }])
    s = await o.save()
    wb = await s.text('xl/workbook.xml')
    expect(wb).toContain('<definedName name="Total">Data!$B$2:$B$4</definedName>')
    sh = await s.text('xl/worksheets/sheet1.xml')
    expect(sh).toContain('<autoFilter ref="A1:C4"/>')
    expect(sh).toContain('<hyperlinks><hyperlink ref="C4" location="Other!A2"/></hyperlinks>')
    expect(sh).toContain('sqref="C2:C4"')
    expect(sh).toContain('<brk id="2" ')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'deleteRows', at: 3, n: 1 }])
    sh = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect(sh).not.toContain('<hyperlinks')
  })

  test('columns inserted in a table and deleted from it are its columns; an array formula stays one', async () => {
    const o = await openBook(await mk())
    await o.engine.call('xlsx.sheetOp', [0, { op: 'insertCols', at: 1, n: 1 }])
    let s = await o.save()
    let t = await s.text('xl/tables/table1.xml')
    expect(t).toContain('ref="A1:D5"')
    expect(t).toContain('<tableColumns count="4"><tableColumn id="1" name="Name"/><tableColumn id="4" name="Column4"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Kind"/></tableColumns>')
    expect(await o.shown(0, 0, 1)).toBe('Column4')
    expect(await s.text('xl/workbook.xml')).toContain('<definedName name="Total">Data!$C$2:$C$5</definedName>')
    expect(await s.text('xl/drawings/drawing1.xml')).toContain('<xdr:col>5</xdr:col>')
    expect(await s.text('xl/worksheets/sheet1.xml')).toContain('<conditionalFormatting sqref="C2:C5"><cfRule type="expression" priority="1"><formula>$C2&gt;10</formula>')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'deleteCols', at: 1, n: 2 }])
    s = await o.save()
    t = await s.text('xl/tables/table1.xml')
    expect(t).toContain('ref="A1:B5"')
    expect(t).toContain('<tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="3" name="Kind"/></tableColumns>')
    expect(await s.text('xl/workbook.xml')).toContain('<definedName name="Total">Data!#REF!</definedName>')
  })

  test('a sheet renamed: charts, validations and names read it by its name of now', async () => {
    const o = await openBook(await mk())
    await o.engine.call('xlsx.bookOp', [{ op: 'rename', index: 1, name: 'Lists 1' }])
    await o.engine.call('xlsx.bookOp', [{ op: 'rename', index: 0, name: 'D2' }])
    const s = await o.save()
    const sh = await s.text('xl/worksheets/sheet1.xml')
    expect(sh).toContain("<formula1>'Lists 1'!$A$1:$A$3</formula1>")
    expect(sh).toContain('<autoFilter ref="A1:C5"/>')
    expect(await s.text('xl/charts/chart1.xml')).toContain("<c:f>'D2'!$B$2:$B$5</c:f>")
    const wb = await s.text('xl/workbook.xml')
    expect(wb).toContain("<definedName name=\"Kinds\">'Lists 1'!$A$1:$A$3</definedName>")
    expect(wb).toContain("localSheetId=\"0\">'D2'!$A$1:$C$5</definedName>")
  })
})

describe('xlsx: whole columns of a sheet longer than Excel allows (P2-024)', () => {
  test('SUM(A:A) adds every row the sheet has', async () => {
    const { evaluateFormula } = await import('../src/core/xlsx/formula')
    const rowsInSheet = 1_048_576 + 1000
    let read = 0
    const v = evaluateFormula('SUM(A:A)', {
      provider: {
        value: (_s, r, c) => (read++, c === 0 && (r < 3 || r >= rowsInSheet - 2) ? 1 : null),
        bounds: () => ({ rows: rowsInSheet, cols: 1 }),
      },
      r: 0,
      c: 5,
    })
    expect(v).toBe(5)
    expect(read).toBe(rowsInSheet)
    // a sheet of a few rows: its rows, not a million
    read = 0
    expect(evaluateFormula('SUM(A:A)', { provider: { value: () => (read++, 2), bounds: () => ({ rows: 4, cols: 1 }) }, r: 0, c: 5 })).toBe(8)
    expect(read).toBe(4)
  })
})

describe('xlsx: a sheet with a formula per row', () => {
  test('an edit finds what reads the cell without looking at every range of the sheet', async () => {
    // 40 000 rows, each looking a table up, adding to a running total and summing a few rows: quadratic, this takes minutes
    const N = 40_000
    const grid: any[][] = []
    for (let r = 1; r <= N; r++) grid.push([r, 2, { f: 'VLOOKUP(1,$A$1:$B$50,2,FALSE)', v: 2 }, { f: r === 1 ? 'C1' : `D${r - 1}+C${r}`, v: 2 * r }, { f: `SUM(B${r}:B${r + 2})`, v: 6 }])
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows(grid) }] }))
    const t0 = performance.now()
    await o.set(0, 'B1', '5')
    await o.set(0, 'B30000', '7')
    expect(performance.now() - t0).toBeLessThan(8000)
    expect(await o.shown(0, 0, 2)).toBe('5')
    expect(await o.shown(0, N - 1, 2)).toBe('5')
    expect(await o.shown(0, N - 1, 3)).toBe(String(5 * N))
    expect(await o.shown(0, 0, 4)).toBe('9')
    expect(await o.shown(0, 29_997, 4)).toBe('11')
    expect(await o.shown(0, 29_999, 4)).toBe('11')
    expect(await o.shown(0, 30_000, 4)).toBe('6')
  }, 30000)
})

describe('xlsx: files written loosely are read without losing cells (PR25-002, -003, -004)', () => {
  test('a cell whose reference names another row, or has leading zeros, is kept; cells out of order are put in order', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: '<row r="2"><c r="A3"><v>33</v></c><c r="B2"><v>22</v></c></row><row r="3"><c r="A03"><v>2</v></c></row><row r="4"><c r="D4"><v>4</v></c><c r="B4"><v>2</v></c><c r="C4"><f>B4+1</f><v>3</v></c><c r="B4"><v>20</v></c></row>' }] }))
    expect(await o.shown(0, 1, 0)).toBe('33')
    expect(await o.shown(0, 1, 1)).toBe('22')
    expect(await o.shown(0, 2, 0)).toBe('2')
    // row 4 as B, C, D: an edit finds each cell once
    const b: any = await o.engine.call('xlsx.block', [0, 3, 4])
    expect(b.rows[0].cells.map((c: any) => [c[0], c[1]])).toEqual([[1, '20'], [2, '3'], [3, '4']])
    await o.set(0, 'B4', '5')
    await o.set(0, 'A4', '1')
    expect(await o.shown(0, 3, 2)).toBe('6')
    const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect([...xml.matchAll(/<c r="([A-Z]+4)"/g)].map((m) => m[1])).toEqual(['A4', 'B4', 'C4', 'D4'])
  })

  test('a column record past the last column, a width or a height past the largest: kept, as large as allowed', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', head: '<cols><col min="2" max="16385" width="20" hidden="1"/><col min="1" max="1" width="300"/></cols>', rows: '<row r="1" ht="500"><c r="A1"><v>1</v></c></row>' }] }))
    const p = o.events.filter((e) => e.event === 'xlsx:progress').map((e) => e.data)
    expect(p.find((x) => x.colInfo)?.colInfo).toEqual([{ min: 1, max: 16383, w: 0, prevW: 140 }, { min: 0, max: 0, w: 1785 }])
    expect(p.flatMap((x) => x.rowSizes)).toContainEqual([0, 546])
  })
})

describe('xlsx: TEXT() shows a number as its cell does (PR24-001)', () => {
  test('large numbers keep their digits', async () => {
    const { evaluateFormula } = await import('../src/core/xlsx/formula')
    const text = (f: string) => evaluateFormula(f, { provider: { value: () => null }, r: 0, c: 0 })
    expect(text('TEXT(1E+21,"0.00")')).toBe('1000000000000000000000.00')
    expect(text('TEXT(1E+21,"#,##0.00")')).toBe('1,000,000,000,000,000,000,000.00')
    expect(text('TEXT(1E+21,"0%")')).toBe('100000000000000000000000%')
    expect(text('TEXT(9.995E+20,"0%")')).toBe('99950000000000000000000%')
    expect(text('TEXT(770467519760130,"0.00")')).toBe('770467519760130.00')
    // as before for the numbers of every day
    expect(text('TEXT(1234.5678,"#,##0.00")')).toBe('1,234.57')
    expect(text('TEXT(2.675,"0.00")')).toBe('2.68')
    expect(text('TEXT(-0.004,"0.00")')).toBe('0.00')
    expect(text('TEXT(-1.5,"0.0")')).toBe('-1.5')
    expect(text('TEXT(0.256,"0.0%")')).toBe('25.6%')
    expect(text('TEXT(5,"000")')).toBe('005')
    expect(text('TEXT(0.5,"#.##")')).toBe('.5')
  })
})

describe('xlsx: sheets, replace (P2-011 links, P2-018, P2-019, P2-023 / N1-017)', () => {
  test('the text put by Replace is taken as it is', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([['price'], ['abc'], ['USD 5']]) }] }))
    await o.engine.call('xlsx.replace', [0, 'price', '$&$&'])
    await o.engine.call('xlsx.replace', [0, 'abc', "$'X$`"])
    await o.engine.call('xlsx.replace', [0, 'USD', '$1'])
    expect([await o.shown(0, 0, 0), await o.shown(0, 1, 0), await o.shown(0, 2, 0)]).toEqual(['$&$&', "$'X$`", '$1 5'])
  })

  test('the last sheet that shows is not deleted; a deleted sheet takes the calculation chain with it', async () => {
    const o = await openBook(
      await book({
        sheets: [{ name: 'Shown', rows: rows([[1]]) }, { name: 'Hidden', state: 'hidden', rows: rows([[{ f: 'Shown!A1', v: 1 }]]) }, { name: 'Other', rows: rows([[{ f: 'A2', v: 0 }]]) }],
        parts: { 'xl/calcChain.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="2"/><c r="A1" i="3"/></calcChain>' },
        types: '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>',
        rels: '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>',
      }),
    )
    await o.engine.call('xlsx.bookOp', [{ op: 'delete', index: 2 }])
    await expect(o.engine.call('xlsx.bookOp', [{ op: 'delete', index: 0 }])).rejects.toThrow(/not hidden/)
    const s = await o.save()
    expect(s.has('xl/calcChain.xml')).toBe(false)
    expect(await s.text('xl/_rels/workbook.xml.rels')).not.toContain('calcChain')
    expect(await s.text('[Content_Types].xml')).not.toContain('calcChain')
    expect((await openBook(s.blob)).info.xlsx.sheets.map((x: any) => x.name)).toEqual(['Shown', 'Hidden'])
  })

  test('a link to a place in the workbook follows the sheet renamed and the rows moved', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'A', rows: rows([['go'], ['here']]), tail: `<hyperlinks><hyperlink ref="A1" location="'Sheet 2'!B2"/><hyperlink ref="A2" location="A!A9"/><hyperlink ref="A2" location="Top"/></hyperlinks>` }, { name: 'Sheet 2', rows: rows([[1]]) }] }))
    await o.engine.call('xlsx.bookOp', [{ op: 'rename', index: 1, name: 'Targets' }])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'insertRows', at: 4, n: 2 }])
    const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<hyperlink ref="A1" location="Targets!B2"/><hyperlink ref="A2" location="A!A11"/><hyperlink ref="A2" location="Top"/>')
  })
})
