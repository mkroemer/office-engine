/**
 * Regressions from the second QA pass (P2-012…P2-030, N1-026): what a saved workbook keeps of the file it was read
 * from, and what a protected workbook refuses.
 */
import { describe, expect, test } from 'bun:test'
import { at, book, openBook, rows } from './xlsx-builder'

const SHEET = 'xl/worksheets/sheet1.xml'
const colEls = (xml: string) => [...(/<cols>(.*?)<\/cols>/.exec(xml)?.[1] ?? '').matchAll(/<col\b[^>]*\/>/g)].map((m) => m[0])
const attr = (el: string, name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(el)?.[1]
const rowTag = (xml: string, r: number) => new RegExp(`<row r="${r}"[^>]*>`).exec(xml)?.[0]
const call = (o: Awaited<ReturnType<typeof openBook>>, op: any) => o.engine.call('xlsx.sheetOp', [0, op])

const COLS =
  '<cols><col min="1" max="1" width="20" style="3" customWidth="1"/><col min="2" max="3" width="9.140625" outlineLevel="1" hidden="1"/>' +
  '<col min="4" max="4" width="12" customWidth="1" bestFit="1"/><col min="5" max="16384" width="9.140625" style="2"/></cols>'

describe('P2-012: columns and rows keep what the model does not edit', () => {
  test('changing a width keeps the style of the column, outline levels and the width of hidden columns', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', head: COLS, rows: rows([[1, 2, 3, 4, 5]]) }] }))
    await call(o, { op: 'colWidth', c0: 3, c1: 3, px: 120 })
    const cols = colEls(await (await o.save()).text(SHEET))
    expect(attr(cols[0], 'style')).toBe('3')
    const hidden = cols.find((c) => attr(c, 'min') === '2')!
    expect(attr(hidden, 'outlineLevel')).toBe('1')
    expect(attr(hidden, 'hidden')).toBe('1')
    expect(Number(attr(hidden, 'width'))).toBeGreaterThan(9)
    const d = cols.find((c) => attr(c, 'min') === '4')!
    expect(Number(attr(d, 'width'))).toBeCloseTo(120 / 7, 2)
    expect(attr(d, 'bestFit')).toBe('1')
    expect(attr(cols[cols.length - 1], 'style')).toBe('2')
  })

  test('a width set over a column with a style keeps the style, also where the width splits it', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', head: COLS, rows: rows([[1]]) }] }))
    await call(o, { op: 'colWidth', c0: 6, c1: 7, px: 90 })
    const cols = colEls(await (await o.save()).text(SHEET))
    const spans = cols.map((c) => [Number(attr(c, 'min')), Number(attr(c, 'max'))])
    // sorted, no overlap
    for (let i = 1; i < spans.length; i++) expect(spans[i][0]).toBeGreaterThan(spans[i - 1][1])
    const hit = cols.filter((c) => Number(attr(c, 'min')) >= 7 && Number(attr(c, 'max')) <= 8)
    expect(hit.length).toBe(1)
    expect(attr(hit[0], 'style')).toBe('2')
    expect(Number(attr(hit[0], 'width'))).toBeCloseTo(90 / 7, 2)
    // what the width did not touch stays with its style
    expect(cols.filter((c) => attr(c, 'style') === '2').length).toBe(3)
  })

  test('hiding and showing a column keeps its width and style', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', head: COLS, rows: rows([[1]]) }] }))
    await call(o, { op: 'hideCols', a: 0, b: 0, hidden: true })
    let cols = colEls(await (await o.save()).text(SHEET))
    expect(attr(cols[0], 'hidden')).toBe('1')
    expect(attr(cols[0], 'style')).toBe('3')
    expect(Number(attr(cols[0], 'width'))).toBeCloseTo(20, 2)
    await call(o, { op: 'hideCols', a: 1, b: 2, hidden: false })
    cols = colEls(await (await o.save()).text(SHEET))
    const shown = cols.find((c) => attr(c, 'min') === '2')!
    expect(attr(shown, 'hidden')).toBeUndefined()
    expect(attr(shown, 'outlineLevel')).toBe('1')
  })

  test('inserting a column shifts the columns with their style and outline', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', head: COLS, rows: rows([[1, 2, 3, 4, 5]]) }] }))
    await call(o, { op: 'insertCols', at: 0, n: 1 })
    const cols = colEls(await (await o.save()).text(SHEET))
    const second = cols.find((c) => attr(c, 'min') === '2')!
    expect(attr(second, 'style')).toBe('3')
    const hidden = cols.find((c) => attr(c, 'min') === '3')!
    expect(attr(hidden, 'outlineLevel')).toBe('1')
    expect(Number(attr(hidden, 'width'))).toBeGreaterThan(9)
  })

  const ROWS =
    '<row r="1" s="3" customFormat="1"><c r="A1"><v>1</v></c></row><row r="2" outlineLevel="1" hidden="1"><c r="A2"><v>2</v></c></row>' +
    '<row r="3" outlineLevel="1" hidden="1"/><row r="4" collapsed="1" ht="30" customHeight="1"><c r="A4"><v>4</v></c></row><row r="5" s="4" customFormat="1" thickBot="1"/>'

  test('inserting rows keeps the style, outline and collapse of rows, also of rows without cells', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: ROWS }] }))
    await call(o, { op: 'insertRows', at: 1, n: 1 })
    const xml = await (await o.save()).text(SHEET)
    expect(rowTag(xml, 1)).toMatch(/\ss="3"/)
    expect(rowTag(xml, 1)).toMatch(/customFormat="1"/)
    expect(rowTag(xml, 3)).toMatch(/outlineLevel="1"/)
    expect(rowTag(xml, 3)).toMatch(/hidden="1"/)
    expect(rowTag(xml, 4)).toMatch(/outlineLevel="1"/)
    expect(rowTag(xml, 5)).toMatch(/collapsed="1"/)
    expect(rowTag(xml, 5)).toMatch(/ht="30"/)
    expect(rowTag(xml, 6)).toMatch(/\ss="4"/)
    expect(rowTag(xml, 6)).toMatch(/thickBot="1"/)
  })

  test('changing the height of a row keeps its style', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: ROWS }] }))
    await call(o, { op: 'rowHeight', r0: 0, r1: 0, px: 40 })
    await call(o, { op: 'rowHeight', r0: 4, r1: 4, px: 40 })
    const xml = await (await o.save()).text(SHEET)
    expect(rowTag(xml, 1)).toMatch(/\ss="3"/)
    expect(rowTag(xml, 1)).toMatch(/ht="30"/)
    expect(rowTag(xml, 5)).toMatch(/\ss="4"/)
    expect(rowTag(xml, 5)).toMatch(/thickBot="1"/)
  })

  test('rows and columns the edit did not touch are not given attributes the file did not have', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1], [2]]) }] }))
    await call(o, { op: 'insertRows', at: 0, n: 1 })
    const xml = await (await o.save()).text(SHEET)
    expect(rowTag(xml, 2)).toBe('<row r="2">')
  })

  test('undo and redo of a structural edit keep them', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', head: COLS, rows: ROWS }] }))
    await call(o, { op: 'deleteRows', at: 0, n: 1 })
    await call(o, { op: 'deleteCols', at: 0, n: 1 })
    await o.engine.call('xlsx.undo', [])
    await o.engine.call('xlsx.undo', [])
    await o.engine.call('xlsx.redo', [])
    await o.engine.call('xlsx.redo', [])
    const xml = await (await o.save()).text(SHEET)
    expect(rowTag(xml, 1)).toMatch(/outlineLevel="1"/)
    expect(attr(colEls(xml)[0], 'outlineLevel')).toBe('1')
  })
})

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NOTE_TYPES = '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>'
const noteBook = () =>
  book({
    sheets: [{ name: 'S', rows: rows([['a', 'b'], ['c', 'd'], ['e', 'f']]), tail: '<legacyDrawing r:id="rId1"/>' }],
    types: NOTE_TYPES,
    parts: {
      'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId2" Type="${REL}/comments" Target="../comments1.xml"/></Relationships>`,
      'xl/comments1.xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Old Author</author><author>Carol</author></authors><commentList>' +
        '<comment ref="B2" authorId="0" shapeId="0"><text><r><rPr><b/></rPr><t>Old Author:</t></r><r><t xml:space="preserve">\nCheck me</t></r></text></comment>' +
        '<comment ref="A3" authorId="1" shapeId="0"><text><t>Carol:\nbottom note</t></text></comment></commentList></comments>',
      'xl/drawings/vmlDrawing1.vml':
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"/>' +
        '<v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>' +
        '<v:shape id="_x0000_s1026" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>2</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>',
    },
  })

describe('P2-013: a note that only moved stays a note', () => {
  test('inserting a row above a note writes it back as the note it was, at its new cell', async () => {
    const o = await openBook(await noteBook())
    await call(o, { op: 'insertRows', at: 0, n: 1 })
    const saved = await o.save()
    const xml = await saved.text('xl/comments1.xml')
    expect(xml).toContain('<author>Old Author</author>')
    expect(xml).toContain('<author>Carol</author>')
    expect(xml).toContain('ref="B3"')
    expect(xml).toContain('ref="A4"')
    expect(xml).not.toContain('[Threaded comment]')
    expect(xml).not.toContain('tc=')
    expect(xml).toContain('<r><rPr><b/></rPr><t>Old Author:</t></r>')
    expect(saved.paths.some((p) => p.includes('threadedComment'))).toBe(false)
    const re = await openBook(saved.blob)
    const list = (await re.engine.call('comments.list', [])) as any[]
    expect(list.map((c) => [c.author.name, c.text, c.at.r, c.at.c]).sort()).toEqual([
      ['Carol', 'bottom note', 3, 0],
      ['Old Author', 'Check me', 2, 1],
    ])
  })

  test('deleting a column left of a note moves it; deleting its row removes it', async () => {
    const o = await openBook(await noteBook())
    await call(o, { op: 'deleteCols', at: 0, n: 1 })
    let xml = await (await o.save()).text('xl/comments1.xml')
    expect(xml).toContain('ref="A2"')
    expect(xml).toContain('<author>Old Author</author>')
    expect(xml).not.toContain('Threaded')
    // Carol's note was on column A: gone with it
    expect(xml).not.toContain('Carol')
    await call(o, { op: 'deleteRows', at: 1, n: 1 })
    const saved = await o.save()
    expect(saved.has('xl/comments1.xml')).toBe(false)
  })

  test('a note whose text was changed becomes a thread, as before', async () => {
    const o = await openBook(await noteBook())
    await call(o, { op: 'insertRows', at: 0, n: 1 })
    const list = (await o.engine.call('comments.list', [])) as any[]
    const note = list.find((c) => c.text === 'Check me')
    await o.engine.call('comments.set', [note.id, { ...note, text: 'Changed' }])
    const xml = await (await o.save()).text('xl/comments1.xml')
    expect(xml).toContain('[Threaded comment]')
  })
})

const SST = (si: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="2">${si}</sst>`
const RICH = '<si><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/></rPr><t>Red bold</t></r><r><rPr><sz val="11"/></rPr><t xml:space="preserve"> and plain</t></r></si><si><t>plain</t></si>'

describe('P2-014: shared strings stay shared strings', () => {
  const src = () =>
    book({
      sheets: [{ name: 'S', rows: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2"><v>5</v></c></row>' }],
      types: '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
      rels: `<Relationship Id="rIdSS" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`,
      parts: { 'xl/sharedStrings.xml': SST(RICH) },
    })

  test('inserting rows or columns keeps the index of a cell nothing edited', async () => {
    const o = await openBook(await src())
    await call(o, { op: 'insertRows', at: 0, n: 1 })
    await call(o, { op: 'insertCols', at: 0, n: 1 })
    const saved = await o.save()
    const xml = await saved.text(SHEET)
    expect(xml).toContain('<c r="B2" t="s"><v>0</v></c>')
    expect(xml).toContain('<c r="C2" t="s"><v>1</v></c>')
    expect(await saved.text('xl/sharedStrings.xml')).toContain('<r><rPr><b/>')
    const re = await openBook(saved.blob)
    expect(await re.shown(0, 1, 1)).toBe('Red bold and plain')
  })

  test('formatting a cell keeps its rich text', async () => {
    const o = await openBook(await src())
    await o.engine.call('xlsx.format', [0, { r0: 0, c0: 0, r1: 0, c1: 0 }, { b: true }])
    const xml = await (await o.save()).text(SHEET)
    expect(xml).toMatch(/<c r="A1" s="\d+" t="s"><v>0<\/v><\/c>/)
  })

  test('a cell edited to a text is written as that text', async () => {
    const o = await openBook(await src())
    await o.set(0, 'A1', 'new')
    await call(o, { op: 'insertRows', at: 5, n: 1 })
    const xml = await (await o.save()).text(SHEET)
    expect(xml).toContain('<c r="A1" t="inlineStr"><is><t xml:space="preserve">new</t></is></c>')
  })
})

describe('P2-016: text is written as ST_Xstring', () => {
  const cases: [string, string][] = [
    ['_x0041_', '_x005F_x0041_'],
    ['a_x000D_b', 'a_x005F_x000D_b'],
    ['bell\u0007x', 'bell_x0007_x'],
    ['vt\u000bx', 'vt_x000B_x'],
    ['nul\u0000x', 'nul_x0000_x'],
    ['cr\rx', 'cr_x000D_x'],
    ['_x00', '_x00'],
    ['_x005F_', '_x005F_x005F_'],
  ]

  test('literal _xHHHH_ is escaped, characters XML cannot hold are written as _xHHHH_, and both read back', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    for (const [i, [text]] of cases.entries()) await o.engine.call('xlsx.set', [0, [{ r: i, c: 1, input: text }]])
    const saved = await o.save()
    const xml = await saved.text(SHEET)
    for (const [i, [, written]] of cases.entries()) expect(xml).toContain(`<c r="B${i + 1}" t="inlineStr"><is><t xml:space="preserve">${written}</t></is></c>`)
    const re = await openBook(saved.blob)
    for (const [i, [text]] of cases.entries()) expect(((await re.engine.call('xlsx.cell', [0, i, 1])) as any).value).toBe(text)
  })

  test('so is the text a formula returns', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'B1', '="_x0041_"&CHAR(7)')
    const saved = await o.save()
    expect(await saved.text(SHEET)).toContain('<v>_x005F_x0041__x0007_</v>')
    const re = await openBook(saved.blob)
    expect(((await re.engine.call('xlsx.cell', [0, 0, 1])) as any).value).toBe('_x0041_\u0007')
  })

  test('a text the file already had in a row left alone is not rewritten', async () => {
    const src = await book({ sheets: [{ name: 'S', rows: '<row r="1"><c r="A1" t="inlineStr"><is><t>_x005F_x0041_</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row>' }] })
    const o = await openBook(src)
    await o.set(0, 'A2', '2')
    const saved = await o.save()
    expect(await saved.text(SHEET)).toContain('<t>_x005F_x0041_</t>')
    const re = await openBook(saved.blob)
    expect(((await re.engine.call('xlsx.cell', [0, 0, 0])) as any).value).toBe('_x0041_')
  })
})

describe('P2-020 / N1-026: a formula that returns a text starting with = keeps its value', () => {
  test('="=abc" is saved with its cached text', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'B1', '="=abc"')
    await o.set(0, 'C1', '="="&"A1"')
    const saved = await o.save()
    const xml = await saved.text(SHEET)
    expect(xml).toContain('<c r="B1" t="str"><f>"=abc"</f><v>=abc</v></c>')
    expect(xml).toContain('t="str"><f>"="&amp;"A1"</f><v>=A1</v></c>')
    const re = await openBook(saved.blob)
    expect(await re.shown(0, 0, 1)).toBe('=abc')
    expect(await re.formula(0, 0, 1)).toBe('="=abc"')
  })

  test('also when the sheet is regenerated', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'B1', '="=abc"')
    await call(o, { op: 'insertRows', at: 0, n: 1 })
    expect(await (await o.save()).text(SHEET)).toContain('<c r="B2" t="str"><f>"=abc"</f><v>=abc</v></c>')
  })
})

const XF_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>' +
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0"/></xf>' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="1" hidden="1"/></xf></cellXfs></styleSheet>'
const grid = '<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row>'

describe('P2-030: protection', () => {
  const sheet = (tail: string, head = '') => book({ styles: XF_STYLES, sheets: [{ name: 'S', head, rows: grid, tail }] })

  test('a locked cell of a protected sheet is refused, an unlocked one is not', async () => {
    const o = await openBook(await sheet('<sheetProtection sheet="1" objects="1" scenarios="1"/>'))
    await expect(o.set(0, 'A1', '9')).rejects.toThrow(/protected/i)
    await expect(o.set(0, 'A3', '9')).rejects.toThrow(/protected/i)
    await o.set(0, 'B1', '9')
    expect(await o.shown(0, 0, 1)).toBe('9')
    expect(await o.shown(0, 0, 0)).toBe('1')
    // one refused cell refuses the whole edit
    await expect(
      o.engine.call('xlsx.set', [0, [{ ...at('B1'), input: '7' }, { ...at('A1'), input: '8' }]]),
    ).rejects.toThrow(/protected/i)
    expect(await o.shown(0, 0, 1)).toBe('9')
  })

  test('so are clearing, filling, pasting and sorting over locked cells', async () => {
    const o = await openBook(await sheet('<sheetProtection sheet="1"/>'))
    await expect(call(o, { op: 'clear', range: { r0: 0, c0: 0, r1: 0, c1: 0 } })).rejects.toThrow(/protected/i)
    await expect(call(o, { op: 'fill', range: { r0: 0, c0: 0, r1: 1, c1: 0 }, dir: 'down' })).rejects.toThrow(/protected/i)
    await expect(call(o, { op: 'sort', range: { r0: 0, c0: 0, r1: 1, c1: 0 }, col: 0 })).rejects.toThrow(/protected/i)
    expect(await o.shown(0, 0, 0)).toBe('1')
    const clip: any = await o.engine.call('xlsx.copy', [0, { r0: 0, c0: 1, r1: 0, c1: 1 }])
    const paste = (r0: number, c0: number) => call(o, { op: 'paste', range: { r0, c0, r1: r0, c1: c0 }, cells: clip.cells, from: { sheet: clip.sheet, r0: 0, c0: 1 } })
    await expect(paste(0, 0)).rejects.toThrow(/protected/i)
    await paste(0, 1)
  })

  test('locked cells are not cut: the paste that moves them is refused, on a sheet that is not protected too', async () => {
    const o = await openBook(await book({ styles: XF_STYLES, sheets: [{ name: 'S', rows: grid, tail: '<sheetProtection sheet="1"/>' }, { name: 'Free', rows: rows([[0]]) }] }))
    await o.load(1)
    const clip: any = await o.engine.call('xlsx.copy', [0, { r0: 0, c0: 0, r1: 0, c1: 0 }])
    const op = { op: 'paste', range: { r0: 4, c0: 0, r1: 4, c1: 0 }, cells: clip.cells, from: { sheet: clip.sheet, r0: 0, c0: 0 } }
    await expect(o.engine.call('xlsx.sheetOp', [1, { ...op, moved: true }])).rejects.toThrow(/protected.*cut/i)
    expect(await o.shown(1, 4, 0)).toBe('')
    // copied, they are pasted
    await o.engine.call('xlsx.sheetOp', [1, op])
    expect(await o.shown(1, 4, 0)).toBe('1')
  })

  test('a cell in a column or row formatted unlocked is unlocked', async () => {
    const o = await openBook(await sheet('<sheetProtection sheet="1"/>', '<cols><col min="4" max="4" width="9" style="1"/></cols>'))
    await o.set(0, 'D5', 'x')
    await expect(o.set(0, 'E5', 'x')).rejects.toThrow(/protected/i)
  })

  test('inserting or deleting rows and columns is refused unless the sheet allows it', async () => {
    const o = await openBook(await sheet('<sheetProtection sheet="1"/>'))
    for (const op of ['insertRows', 'deleteRows', 'insertCols', 'deleteCols']) await expect(call(o, { op, at: 0, n: 1 })).rejects.toThrow(/protected/i)
    const free = await openBook(await sheet('<sheetProtection sheet="1" insertRows="0"/>'))
    await call(free, { op: 'insertRows', at: 0, n: 1 })
    await expect(call(free, { op: 'insertCols', at: 0, n: 1 })).rejects.toThrow(/protected/i)
  })

  test('a sheet without sheet="1" is not protected, whatever its cells say', async () => {
    const o = await openBook(await sheet('<sheetProtection sheet="0"/>'))
    await o.set(0, 'A1', '9')
    await call(o, { op: 'insertRows', at: 0, n: 1 })
    const plain = await openBook(await sheet(''))
    await plain.set(0, 'A1', '9')
  })

  test('resizing and hiding follow formatColumns / formatRows', async () => {
    const o = await openBook(await sheet('<sheetProtection sheet="1"/>'))
    await expect(call(o, { op: 'colWidth', c0: 0, c1: 0, px: 100 })).rejects.toThrow(/protected/i)
    await expect(call(o, { op: 'hideRows', a: 0, b: 0, hidden: true })).rejects.toThrow(/protected/i)
    const free = await openBook(await sheet('<sheetProtection sheet="1" formatColumns="0" formatRows="0"/>'))
    await call(free, { op: 'colWidth', c0: 0, c1: 0, px: 100 })
    await call(free, { op: 'hideRows', a: 0, b: 0, hidden: true })
  })

  test('protection is kept when the workbook is saved', async () => {
    const o = await openBook(await sheet('<sheetProtection sheet="1" password="DAA7"/>'))
    await o.set(0, 'B1', '5')
    const saved = await o.save()
    expect(await saved.text(SHEET)).toContain('<sheetProtection sheet="1" password="DAA7"/>')
    const re = await openBook(saved.blob)
    await expect(re.set(0, 'A1', '1')).rejects.toThrow(/protected/i)
  })

  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  const locked = async (extra = '') => {
    // the workbook part with its structure locked
    return book({
      sheets: [{ name: 'S', rows: grid }, { name: 'T', rows: grid }],
      parts: {
        'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS}" xmlns:r="${REL}"><workbookProtection ${extra || 'lockStructure="1"'}/><sheets><sheet name="S" sheetId="1" r:id="rId1"/><sheet name="T" sheetId="2" r:id="rId2"/></sheets></workbook>`,
      },
    })
  }

  test('a workbook with locked structure refuses adding, renaming, deleting and moving sheets, not editing cells', async () => {
    const o = await openBook(await locked())
    await expect(o.engine.call('xlsx.bookOp', [{ op: 'add', after: 0, name: 'Sneaky' }])).rejects.toThrow(/structure/i)
    await expect(o.engine.call('xlsx.bookOp', [{ op: 'rename', index: 0, name: 'X' }])).rejects.toThrow(/structure/i)
    await expect(o.engine.call('xlsx.bookOp', [{ op: 'delete', index: 1 }])).rejects.toThrow(/structure/i)
    await expect(o.engine.call('xlsx.bookOp', [{ op: 'move', from: 1, to: 0 }])).rejects.toThrow(/structure/i)
    await o.set(0, 'A1', '9')
    expect(await o.shown(0, 0, 0)).toBe('9')
  })

  test('lockStructure="0" or only lockWindows does not lock', async () => {
    for (const extra of ['lockStructure="0"', 'lockWindows="1"']) {
      const o = await openBook(await locked(extra))
      await o.engine.call('xlsx.bookOp', [{ op: 'rename', index: 0, name: 'X' }])
    }
  })
})
