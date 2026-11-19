import { describe, expect, test } from 'bun:test'
import { parseXml } from '../src/core/xml'
import { parseSheetStyles } from '../src/core/xlsx/styles'
import { book, openBook } from './xlsx-builder'

const theme = { colors: {}, major: 'Calibri', minor: 'Calibri', fillStyles: [], bgFillStyles: [] } as never

const styles = (fills: string, borders = '<border><left/><right/><top/><bottom/><diagonal/></border>', xf = '') =>
  parseSheetStyles(
    parseXml(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill>${fills}</fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>${borders}</borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="1" borderId="1">${xf}</xf></cellXfs></styleSheet>`,
    ),
    theme,
  ).xfs[1].css

describe('P2-037 cell styles', () => {
  test('a pattern fill is its colors mixed by the pattern, not the solid foreground', () => {
    const css = styles('<fill><patternFill patternType="gray125"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>')
    // 12.5 % red over white
    expect(css.background).toBe('#ffdfdf')
    expect(styles('<fill><patternFill patternType="darkGray"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>').background).toBe('#404040')
    // a solid fill stays as it is; a pattern without a background color mixes with white
    expect(styles('<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>').background).toBe('#00ff00')
    expect(styles('<fill><patternFill patternType="lightGray"><fgColor rgb="FF0000FF"/></patternFill></fill>').background).toBe('#bfbfff')
  })

  test('diagonal borders are drawn as lines across the cell', () => {
    const css = styles(
      '<fill><patternFill patternType="none"/></fill>',
      '<border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"><color rgb="FFFF0000"/></diagonal></border>',
    )
    expect(String(css.backgroundImage)).toContain('to top right')
    expect(String(css.backgroundImage)).toContain('to bottom right')
    expect(String(css.backgroundImage)).toContain('#ff0000')
    const up = styles('<fill><patternFill patternType="none"/></fill>', '<border diagonalUp="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border>')
    // "up" runs from the bottom left to the top right corner: the 50 % line of "to bottom right"
    expect(String(up.backgroundImage)).toContain('to bottom right')
    expect(String(up.backgroundImage)).not.toContain('to top right')
  })

  test('shrink to fit is carried to the grid', () => {
    const css = styles('<fill><patternFill patternType="none"/></fill>', undefined, '<alignment shrinkToFit="1"/>')
    expect(css.shrinkToFit).toBe(1)
  })
})

describe('P2-037: rich text', () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  const sst =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS}" count="3" uniqueCount="3">` +
    '<si><r><rPr><b/><color rgb="FFFF0000"/><sz val="14"/><rFont val="Arial"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> plain </t></r>' +
    '<r><rPr><i/><vertAlign val="superscript"/></rPr><t>2</t></r><rPh sb="0" eb="1"><t>ph</t></rPh></si>' +
    '<si><t>plain</t></si><si><r><t>one</t></r><r><t>font</t></r></si></sst>'
  const open = async () =>
    openBook(
      await book({
        sheets: [{ name: 'S', rows: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>' }],
        parts: { 'xl/sharedStrings.xml': sst },
        rels: '<Relationship Id="rIdSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
        types: '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
      }),
    )

  test('the parts of a string with fonts of their own reach the view; a string of one font has none', async () => {
    const o = await open()
    const b: any = await o.engine.call('xlsx.block', [0, 0, 1])
    const [a1, b1, c1] = b.rows[0].cells
    expect(a1[1]).toBe('Bold plain 2')
    expect(a1[5]).toEqual([{ t: 'Bold', own: true, b: true, color: '#ff0000', sz: 14, font: 'Arial' }, { t: ' plain ' }, { t: '2', own: true, i: true, va: 'superscript' }])
    expect(b1[1]).toBe('plain')
    expect(b1[5]).toBeUndefined()
    expect(c1[1]).toBe('onefont')
    expect(c1[5]).toBeUndefined()
  })

  test('typing over a rich text leaves a plain one; the cells not typed into keep theirs in the saved file', async () => {
    const o = await open()
    await o.set(0, 'B1', 'typed')
    const b: any = await o.engine.call('xlsx.block', [0, 0, 1])
    expect(b.rows[0].cells[0][5]?.length).toBe(3)
    expect(b.rows[0].cells[1][5]).toBeUndefined()
    const saved = await o.save()
    expect(await saved.text('xl/sharedStrings.xml')).toContain('<rFont val="Arial"/>')
    expect(await saved.text('xl/worksheets/sheet1.xml')).toMatch(/<c r="A1"[^>]*t="s"/)
  })
})
