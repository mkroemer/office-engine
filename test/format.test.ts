import { describe, expect, test } from 'bun:test'
import type { FormatState } from '../src/react/docx/DocxEditor'
import { cellFormat, cellPatch, docxFormat, pptxFormat } from '../src/react/edit/format'

describe('formatting as TextFormat', () => {
  test('docx: marks, font name from the CSS stack, half points, list', () => {
    const s = { b: true, i: false, u: true, strike: false, font: '"Calibri", sans-serif', sz: 11.26, color: '#ff0000', hl: 'yellow', align: 'center', lst: 'number', inTable: false, canUndo: true, canRedo: false } as FormatState
    expect(docxFormat(s)).toEqual({ bold: true, italic: false, underline: true, strike: false, font: 'Calibri', size: 11.5, color: '#ff0000', highlight: 'yellow', align: 'center', list: 'number' })
    expect(docxFormat({ ...s, lst: undefined, align: 'both' }).list).toBeNull()
    expect(docxFormat({ ...s, align: 'both' }).align).toBeUndefined()
  })

  test('pptx: bullets other than bullet / number are no list', () => {
    const f = pptxFormat({ b: false, i: true, u: false, strike: true, sz: 18, algn: 'justify', bullet: 'none', lvl: 0 })
    expect(f).toMatchObject({ italic: true, strike: true, size: 18, align: 'justify', list: null })
  })

  test('cells: CSS of the active cell, the base font where the cell has none, no lists', () => {
    const f = cellFormat({ fontWeight: 700, textDecoration: 'underline line-through', fontSize: '14pt', background: '#ffff00', textAlign: 'right' }, { fontFamily: 'Aptos, sans-serif', fontSize: 11 * (4 / 3) })
    expect(f).toEqual({ bold: true, italic: false, underline: true, strike: true, font: 'Aptos', size: 14, color: undefined, highlight: '#ffff00', align: 'right' })
    expect(cellFormat({}, { fontSize: '11pt' }).size).toBe(11)
    expect(cellFormat({ background: 'linear-gradient(red, blue)' }, {}).highlight).toBeUndefined()
  })

  test('cell patches: a toggle flips the current state; what cells lack is null', () => {
    const now = cellFormat({ fontWeight: 700 }, {})
    expect(cellPatch({ op: 'toggle', mark: 'bold' }, now)).toEqual({ b: false })
    expect(cellPatch({ op: 'toggle', mark: 'italic' }, now)).toEqual({ i: true })
    expect(cellPatch({ op: 'highlight', color: '#ffff00' }, now)).toEqual({ fill: '#ffff00' })
    expect(cellPatch({ op: 'size', pt: 9 }, now)).toEqual({ sz: 9 })
    expect(cellPatch({ op: 'clear' }, now)).toEqual({ clear: true })
    expect(cellPatch({ op: 'align', align: 'justify' }, now)).toBeNull()
    expect(cellPatch({ op: 'list', list: 'bullet' }, now)).toBeNull()
  })
})
