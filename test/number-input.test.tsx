/**
 * Numbers typed in a locale (#17): the parser, a number edited again in a cell, the fields of the toolbar and the
 * dialogs. Cells and collaboration: xlsx-locale.test.ts.
 */
import { describe, expect, test } from 'bun:test'
import type { ReactNode } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { inputLocale, numberSyntax, numberText, parseNumber, readTyped } from '../src/core/number-input'
import { LocaleContext } from '../src/react/common/locale'
import { SizeBox, sizeOf } from '../src/react/edit/Toolbar'
import { PageSetupDialog } from '../src/react/docx/PageSetupDialog'
import { partOf } from '../src/react/docx/HeaderFooterDialog'
import { at, book, openBook, rows, type Opened } from './xlsx-builder'

const NBSP = '\u00a0'
const NNBSP = '\u202f'
const THIN = '\u2009'

describe('the separators of a locale', () => {
  test('without a locale, or with a tag that is not one: a decimal point, commas by thousands', () => {
    for (const tag of [undefined, null, '', 'de_DE', 'not a tag!', 'xx-YY']) {
      const s = numberSyntax(tag)
      expect([s.locale, s.decimal, s.group, s.primary, s.secondary, s.minus, s.percent, s.percentSpace, s.digits]).toEqual([undefined, '.', [','], 3, 3, ['-'], ['%'], false, undefined])
    }
    expect(inputLocale('de-de')).toBe('de-DE')
    expect(inputLocale('de_DE')).toBeUndefined()
    // well formed, but no data: not the browser's locale in its place
    expect(inputLocale('xx-YY')).toBeUndefined()
    expect(inputLocale('de-AT')).toBe('de-AT')
  })

  test('from Intl: German, French, Swiss, Indian, Arabic, Swedish', () => {
    const de = numberSyntax('de-DE')
    expect([de.decimal, de.group, de.percentSpace]).toEqual([',', ['.'], true])
    // a locale grouping with a space reads every space; one with an apostrophe both apostrophes
    expect(numberSyntax('fr-FR').group).toEqual([' ', NBSP, NNBSP, THIN])
    expect(numberSyntax('de-CH').group).toEqual(["'", '’'])
    expect(numberSyntax('de-CH').decimal).toBe('.')
    const inr = numberSyntax('en-IN')
    expect([inr.primary, inr.secondary]).toEqual([3, 2])
    const ar = numberSyntax('ar-EG')
    expect([ar.decimal, ar.group, ar.digits?.join(''), ar.percent]).toEqual(['٫', ['٬'], '٠١٢٣٤٥٦٧٨٩', ['%', '٪']])
    expect(numberSyntax('sv-SE').minus).toEqual(['-', '−'])
    // US English reads as without a locale
    const us = numberSyntax('en-US')
    expect([us.decimal, us.group, us.primary, us.secondary, us.minus, us.percent, us.percentSpace, us.digits]).toEqual(['.', [','], 3, 3, ['-'], ['%'], false, undefined])
  })
})

describe('parseNumber: the whole text, one number', () => {
  const cases: [string | undefined, string, number | null][] = [
    // a decimal point
    [undefined, '1.5', 1.5],
    [undefined, '-1.5', -1.5],
    [undefined, '+.5', 0.5],
    [undefined, '5.', 5],
    [undefined, '1,234', 1234],
    [undefined, '1,234,567.25', 1234567.25],
    [undefined, '1.5e3', 1500],
    [undefined, '-2.5E-2', -0.025],
    // the other convention's separator: refused, never a prefix
    [undefined, '1,5', null],
    [undefined, '12,5', null],
    [undefined, '1.234,5', null],
    [undefined, '1,23', null],
    [undefined, '1,2345', null],
    ['en-US', '1,5', null],
    ['en-US', '1.5', 1.5],
    // a decimal comma
    ['de-DE', '1,5', 1.5],
    ['de-DE', '-1,5', -1.5],
    ['de-DE', ',5', 0.5],
    ['de-DE', '1,5e3', 1500],
    ['de-DE', '1,5E-3', 0.0015],
    ['de-DE', '1.234', 1234],
    ['de-DE', '1.234,5', 1234.5],
    ['de-DE', '1.234.567', 1234567],
    ['de-DE', '1.5', null],
    ['de-DE', '1.23', null],
    ['de-DE', '1.2345', null],
    ['de-DE', '12.34.567', null],
    ['de-DE', '1,234.5', null],
    ['de-DE', '1.234e3', null],
    ['de-DE', '1,5,0', null],
    ['de-DE', '1 234', null],
    // spaces
    ['fr-FR', '1,5', 1.5],
    ['fr-FR', '1 234,5', 1234.5],
    ['fr-FR', `1${NBSP}234,5`, 1234.5],
    ['fr-FR', `1${NNBSP}234${NNBSP}567`, 1234567],
    ['fr-FR', `1${THIN}234`, 1234],
    ['fr-FR', `1 234${NBSP}567`, null],
    ['fr-FR', '1 23', null],
    ['fr-FR', '1.234,5', null],
    ['sv-SE', '−1,5', -1.5],
    ['sv-SE', '-1,5', -1.5],
    ['sv-SE', `1${NBSP}234,5`, 1234.5],
    // apostrophes
    ['de-CH', "1'234.5", 1234.5],
    ['de-CH', '1’234.5', 1234.5],
    ['de-CH', "1'234’567", null],
    ['de-CH', '1,5', null],
    ['de-CH', '1.5', 1.5],
    // India: 3 then 2
    ['en-IN', '12,34,567', 1234567],
    ['en-IN', '1,234', 1234],
    ['en-IN', '12,345.5', 12345.5],
    ['en-IN', '1,234,567', null],
    ['en-IN', '123,456', null],
    // the locale's own digits, or ASCII ones, not mixed
    ['ar-EG', '١٫٥', 1.5],
    ['ar-EG', '١٬٢٣٤٫٥', 1234.5],
    ['ar-EG', '1٫5', 1.5],
    ['ar-EG', '-١٫٥', -1.5],
    ['ar-EG', '؜-١٫٥', -1.5],
    ['ar-EG', '١2', null],
    ['ar-EG', '1.5', null],
    ['hi-IN-u-nu-deva', '१२,३४,५६७', 1234567],
    // not numbers
    [undefined, '', null],
    [undefined, '   ', null],
    [undefined, '-', null],
    [undefined, '.', null],
    [undefined, '1e', null],
    [undefined, '12a', null],
    [undefined, '12pt', null],
    [undefined, 'a12', null],
    [undefined, '0x1F', null],
    [undefined, 'Infinity', null],
    [undefined, 'NaN', null],
    [undefined, '1e999', null],
    [undefined, '1_000', null],
    [undefined, '1..5', null],
    [undefined, '1.2.3', null],
    [undefined, '--1', null],
    [undefined, '1-', null],
    [undefined, '(5)', null],
    [undefined, '1/2', null],
    ['de-DE', '12,5pt', null],
    ['de-DE', '1,,5', null],
    // a field reads no percentage and no amount of money
    [undefined, '50%', null],
    [undefined, '$5', null],
    ['de-DE', '5 €', null],
  ]
  for (const [locale, text, n] of cases) test(`${locale ?? 'no locale'}: ${JSON.stringify(text)} → ${n}`, () => expect(parseNumber(text, locale)).toBe(n))
})

describe('readTyped: what a cell reads', () => {
  test('percentages: the locale writes a space before the sign or not', () => {
    expect(readTyped('12,5%', 'de-DE')).toMatchObject({ kind: 'percent', value: 12.5, decimals: 1 })
    expect(readTyped('50 %', 'de-DE')).toMatchObject({ kind: 'percent', value: 50, decimals: 0 })
    expect(readTyped(`50${NBSP}%`, 'fr-FR')).toMatchObject({ kind: 'percent', value: 50 })
    expect(readTyped('12.5%', 'de-DE')).toBeNull()
    expect(readTyped('50 %', 'en-US')).toBeNull()
    expect(readTyped('50 %')).toBeNull()
    expect(readTyped('٥٠٪', 'ar-EG')).toMatchObject({ kind: 'percent', value: 50 })
    expect(readTyped('1.000%', 'de-DE')).toBeNull()
  })

  test('amounts: grouping and currency, the decimals typed', () => {
    expect(readTyped('1.234,50 €', 'de-DE')).toEqual({ kind: 'amount', value: 1234.5, canon: '1234.50', decimals: 2, grouped: true, currency: '€', after: true })
    expect(readTyped('-€3,75', 'de-DE')).toMatchObject({ kind: 'amount', value: -3.75, currency: '€', after: false })
    expect(readTyped('1.234', 'de-DE')).toMatchObject({ kind: 'amount', value: 1234, grouped: true, decimals: 0 })
    expect(readTyped(`50${NNBSP}000 €`, 'fr-FR')).toMatchObject({ kind: 'amount', value: 50000, currency: '€' })
    expect(readTyped('€1,200.50', 'de-DE')).toBeNull()
    expect(readTyped('$5€', 'de-DE')).toBeNull()
  })

  test('plain numbers keep their decimals and their exponent', () => {
    expect(readTyped('1,50', 'de-DE')).toEqual({ kind: 'plain', value: 1.5, canon: '1.50', decimals: 2, grouped: false })
    expect(readTyped('-1,5e3', 'de-DE')).toEqual({ kind: 'plain', value: -1500, canon: '-1.5e3', decimals: 1, grouped: false })
  })
})

describe('numberText: a number as it is typed, read back the same', () => {
  test('the locale decimal separator, no grouping, ASCII digits', () => {
    expect(numberText(1.5, 'de-DE')).toBe('1,5')
    expect(numberText(1234.5, 'de-DE')).toBe('1234,5')
    expect(numberText(-0.001, 'fr-FR')).toBe('-0,001')
    expect(numberText(1.5e-7, 'de-DE')).toBe('1,5e-7')
    expect(numberText(1e21, 'de-DE')).toBe('1e+21')
    expect(numberText(1.5, 'ar-EG')).toBe('1٫5')
    expect(numberText(1.5)).toBe('1.5')
    expect(numberText(1234.5, 'en-US')).toBe('1234.5')
  })

  test('every number, in every locale, reads back as itself', () => {
    let seed = 17
    const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296)
    const xs = [0, 1, -1, 0.1 + 0.2, 1 / 3, 2 / 3, 1234567.891, -98765.4321, 1e21, 1.5e-7, 5e-324, Number.MAX_VALUE, 123456789012345]
    for (let i = 0; i < 300; i++) xs.push((rand() - 0.5) * 10 ** Math.floor(rand() * 40 - 20))
    for (const locale of [undefined, 'en-US', 'de-DE', 'fr-FR', 'de-CH', 'en-IN', 'ar-EG', 'sv-SE', 'fa-IR', 'hi-IN']) {
      for (const x of xs) {
        const back = readTyped(numberText(x, locale), locale)
        if (back?.kind !== 'plain' || back.value !== (x === 0 ? 0 : x)) throw new Error(`${locale}: ${x} → ${numberText(x, locale)} → ${JSON.stringify(back)}`)
      }
    }
  })
})

// ---------------------------------------------------------------- what a cell edited again shows

const detail = async (o: Opened, ref: string) => (await o.engine.call('xlsx.cell', [0, at(ref).r, at(ref).c])) as { type: string; value: string }
const typeIn = (o: Opened, ref: string, input: string, locale?: string) => o.engine.call('xlsx.set', [0, [{ ...at(ref), input }], locale ? { locale } : undefined])

describe('cells: a number edited again', () => {
  test('is shown in the locale and read back the same (never text, never 15)', async () => {
    const o = await openBook(
      await book({
        sheets: [{ name: 'S', rows: rows([[1.5, -0.001, 1234.5, 1e21, 0.1 + 0.2, { s: 1, v: 0.5 }, 1.5e-7]]) }],
        styles: `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`,
      }),
    )
    for (const locale of ['de-DE', 'fr-FR', 'ar-EG', 'de-CH', undefined]) {
      for (const ref of ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1']) {
        const before = await detail(o, ref)
        // what the formula bar and the cell editor show (XlsxView: numberText of the value)
        const text = numberText(Number(before.value), locale)
        await typeIn(o, ref, text, locale)
        const after = await detail(o, ref)
        expect([locale, ref, text, after.type, after.value]).toEqual([locale, ref, text, 'number', before.value])
      }
    }
    expect(numberText(1.5, 'de-DE')).toBe('1,5')
  })
})

// ---------------------------------------------------------------- fields

describe('fields: the font size, page setup, header and footer', () => {
  test('the font size box reads the whole text in the locale; what is not a size keeps the size it had (null)', () => {
    expect(sizeOf('12,5', 'de-DE')).toBe(12.5)
    expect(sizeOf('12.5', 'de-DE')).toBeNull()
    expect(sizeOf('12,5', 'en-US')).toBeNull()
    expect(sizeOf('12,5pt', 'de-DE')).toBeNull()
    expect(sizeOf('1.000', 'de-DE')).toBeNull()
    expect(sizeOf('10,3', 'fr-FR')).toBe(10.5)
    expect(sizeOf(' 14 ')).toBe(14)
    expect(sizeOf('0')).toBeNull()
    expect(sizeOf('401')).toBeNull()
  })

  test('without a locale the font size box reads a decimal comma too (as since #17), still the whole text only', () => {
    for (const [text, size] of [
      ['12.5', 12.5],
      ['12,5', 12.5],
      [' 12,5 ', 12.5],
      ['10,25', 10.5],
      ['9,0', 9],
      ['14', 14],
    ] as const) {
      expect([text, sizeOf(text)]).toEqual([text, size])
      // a tag that is not one is no locale
      expect([text, sizeOf(text, 'de_DE')]).toEqual([text, size])
    }
    // a prefix, two separators, a comma that could group thousands, nothing: the box keeps its size
    for (const text of ['12pt', '12,5pt', '1.2.5', '1,2,5', '1,234', '0,400', '12,5.5', ',5', '-12,5', '', ' ']) expect([text, sizeOf(text)]).toEqual([text, null])
    // only the size box: a cell or a field without a locale reads no decimal comma
    expect(parseNumber('12,5')).toBeNull()
  })

  const inLocale = (locale: string | undefined, node: ReactNode) => renderToStaticMarkup(<LocaleContext.Provider value={locale}>{node}</LocaleContext.Provider>)

  test('the font size box shows the size in the locale', () => {
    expect(inLocale('de-DE', <SizeBox value={10.5} onChange={() => {}} />)).toContain('value="10,5"')
    expect(inLocale(undefined, <SizeBox value={10.5} onChange={() => {}} />)).toContain('value="10.5"')
    expect(inLocale('de-DE', <SizeBox placeholder={11.5} onChange={() => {}} />)).toContain('placeholder="11,5"')
  })

  test('page setup: text fields in the locale; the browser number fields without one', () => {
    const page = { w: 595.3, h: 841.9, mt: 72, mr: 72, mb: 72, ml: 72 }
    const de = inLocale('de-DE', <PageSetupDialog page={page as never} onOk={() => {}} onClose={() => {}} />)
    expect(de).toMatch(/<input type="text" inputMode="decimal"[^>]*aria-label="Top margin \(cm\)" value="2,54"/)
    const none = inLocale(undefined, <PageSetupDialog page={page as never} onOk={() => {}} onClose={() => {}} />)
    expect(none).toMatch(/<input type="number" inputMode="decimal"[^>]*aria-label="Top margin \(cm\)" value="2.54"/)
  })

  test('header and footer: the size of the text in the locale', () => {
    const blocks = [{ t: 'p', ps: {}, runs: [{ t: 'text', text: 'Draft', s: { sz: 10.5 } }] }] as never
    expect(partOf(blocks, 'de-DE').sz).toBe('10,5')
    expect(partOf(blocks).sz).toBe('10.5')
  })
})
