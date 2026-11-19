/**
 * Regressions from the second QA pass (formula engine and number formats): ids as in the reports
 * (P1-009…P1-026, P2-024, N1-012, N3-006).
 */
import { describe, expect, test } from 'bun:test'
import { colIndex, evaluateFormula, FormulaError, type CellProvider, type Value } from '../src/core/xlsx/formula'
import { formatGeneral, formatValue } from '../src/core/xlsx/numfmt'
import { compare } from '../src/core/xlsx/formula/values'
import { fitNumber } from '../src/react/xlsx/fit'
import { book, openBook } from './xlsx-builder'

type Sheets = Record<string, Record<string, Value>>

const addr = (a: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a)!
  return { r: Number(m[2]) - 1, c: colIndex(m[1]) }
}

/** a provider over in-memory sheets (the first is the formula's own) */
function provider(sheets: Sheets, extra: Partial<CellProvider> = {}): CellProvider {
  const cells = new Map<string, Value>()
  const ext = new Map<string, { rows: number; cols: number }>()
  const first = Object.keys(sheets)[0] ?? 'S'
  for (const [name, data] of Object.entries(sheets)) {
    let rows = 0
    let cols = 0
    for (const [a, v] of Object.entries(data)) {
      const { r, c } = addr(a)
      cells.set(`${name.toLowerCase()}!${r},${c}`, v)
      rows = Math.max(rows, r + 1)
      cols = Math.max(cols, c + 1)
    }
    ext.set(name.toLowerCase(), { rows, cols })
  }
  return {
    value: (s, r, c) => cells.get(`${(s ?? first).toLowerCase()}!${r},${c}`) ?? null,
    bounds: (s) => ext.get((s ?? first).toLowerCase()) ?? { rows: 0, cols: 0 },
    hasSheet: (n) => ext.has(n.toLowerCase()),
    ...extra,
  }
}

const NONE = provider({})
const run = (f: string, p: CellProvider = NONE, sheet?: string): Value => evaluateFormula(f, { provider: p, sheet, r: 40, c: 10 })
const text = (f: string) => run(f)
const err = (v: Value) => (v instanceof FormulaError ? v.code : v)

describe('P1-009: TEXT() formats like the cell does', () => {
  test('scientific', () => {
    expect(text('TEXT(1234.5,"0.00E+00")')).toBe('1.23E+03')
    expect(text('TEXT(0.000123,"0.0E-00")')).toBe('1.2E-04')
    expect(text('TEXT(-1234.5,"0.0E+0")')).toBe('-1.2E+3')
    expect(text('TEXT(12345,"##0.0E+0")')).toBe('12.3E+3')
    expect(text('TEXT(0,"0.00E+00")')).toBe('0.00E+00')
    expect(text('TEXT(99999,"0.0E+0")')).toBe('1.0E+5')
  })
  test('fractions', () => {
    expect(text('TEXT(0.5,"# ?/?")')).toBe(' 1/2')
    expect(text('TEXT(1.5,"# ?/?")')).toBe('1 1/2')
    expect(text('TEXT(2.75,"# ?/?")')).toBe('2 3/4')
    expect(text('TEXT(2.75,"?/?")')).toBe('11/4')
    expect(text('TEXT(1,"# ?/?")')).toBe('1    ')
    expect(text('TEXT(0.5,"# ?/8")')).toBe(' 4/8')
    expect(text('TEXT(0.333,"# ??/??")')).toBe('  1/3 ')
    expect(text('TEXT(0.5,"0 0/0")')).toBe('0 1/2')
    expect(text('TEXT(-1.5,"# ?/?")')).toBe('-1 1/2')
    expect(text('TEXT(0.999,"# ?/?")')).toBe('1    ')
  })
  test('@ with a number and with text', () => {
    expect(text('TEXT(123,"@")')).toBe('123')
    expect(text('TEXT(12.5,"""x""@")')).toBe('x12.5')
    expect(text('TEXT("abc","""x""@")')).toBe('xabc')
  })
  test('characters between the digits', () => {
    expect(text('TEXT(123456789,"000-00-0000")')).toBe('123-45-6789')
    expect(text('TEXT(1234567890,"(###) ###-####")')).toBe('(123) 456-7890')
    expect(text('TEXT(12,"0""-""0")')).toBe('1-2')
    expect(text('TEXT(12,"#-#")')).toBe('1-2')
  })
  test('conditions', () => {
    const f = '"[<1000]0;[<1000000]0.0,""K"";0.0,,""M"""'
    expect(text(`TEXT(500,${f})`)).toBe('500')
    expect(text(`TEXT(1500,${f})`)).toBe('1.5K')
    expect(text(`TEXT(2500000,${f})`)).toBe('2.5M')
    expect(text('TEXT(1,"[=1]""one"";""other""")')).toBe('one')
    expect(text('TEXT(2,"[=1]""one"";""other""")')).toBe('other')
    expect(text('TEXT(150,"[>=100]0.0;0")')).toBe('150.0')
  })
  test('empty sections hide the value', () => {
    expect(text('TEXT(5,";;;")')).toBe('')
    expect(text('TEXT(-5,";;;")')).toBe('')
    expect(text('TEXT(0,";;;")')).toBe('')
    expect(text('TEXT(1,"0;;""z""")')).toBe('1')
    expect(text('TEXT(-1,"0;;""z""")')).toBe('')
    expect(text('TEXT(0,"0;;""z""")')).toBe('z')
  })
  test('? pads with spaces', () => {
    expect(text('TEXT(1.5,"0.0?")')).toBe('1.5 ')
    expect(text('TEXT(5,"??0")')).toBe('  5')
    expect(text('TEXT(5,"?0.0?")')).toBe(' 5.0 ')
  })
  test('currency symbols and locale codes', () => {
    expect(text('TEXT(1234.5,"[$₫-42A]#,##0")')).toBe('₫1,235')
    expect(text('TEXT(5,"[$€-1]0.00")')).toBe('€5.00')
    expect(text('TEXT(45306,"[$-42A]dddd")')).toBe('Thứ Hai')
    expect(text('TEXT(45305,"[$-42A]dddd")')).toBe('Chủ Nhật')
    expect(text('TEXT(45306,"[$-409]dddd")')).toBe('Monday')
  })
  test('fractions of a second', () => {
    expect(text('TEXT(125.46/86400,"mm:ss.0")')).toBe('02:05.5')
    expect(text('TEXT(125.46/86400,"mm:ss.00")')).toBe('02:05.46')
    expect(text('TEXT(59.96/86400,"mm:ss.0")')).toBe('01:00.0')
  })
  test('very large numbers keep their digits', () => {
    expect(text('TEXT(1E+300,"0.00")')).toBe('1' + '0'.repeat(300) + '.00')
    expect(text('TEXT(1E+15,"0")')).toBe('1000000000000000')
    expect(text('TEXT(1E+21,"0")')).toBe('1' + '0'.repeat(21))
    expect(text('TEXT(123456789012345678,"0")')).toBe('123456789012346000')
  })
  test('as before', () => {
    expect(text('TEXT(-5,"0;(0)")')).toBe('(5)')
    expect(text('TEXT(0,"0;-0;;@")')).toBe('')
    expect(text('TEXT("abc","0.00")')).toBe('abc')
    expect(text('TEXT(1.5,"General")')).toBe('1.5')
    expect(text('TEXT(TRUE,"0")')).toBe('TRUE')
    expect(text('TEXT(-1,"yyyy")')).toEqual(new FormulaError('#VALUE!'))
    expect(text('TEXT("1,234.5","0.0")')).toBe('1234.5')
  })
})

describe('P1-010: the cell shows ;;; 0;; locale names, E+0 rounding, early dates', () => {
  const f = (v: number, fmt: string, d1904 = false) => formatValue(v, fmt, d1904).text
  test('empty sections hide', () => {
    for (const v of [1, -1, 0.5, 0, -0.25]) expect(f(v, ';;;')).toBe('')
    expect(f(-1, '0;;"z"')).toBe('')
    expect(f(1, '0;;"z"')).toBe('1')
    expect(f(0, '0;;"z"')).toBe('z')
    expect(f(-1, '0;')).toBe('')
    expect(f(5, ';"neg"')).toBe('')
    expect(f(-5, ';"neg"')).toBe('neg')
    // one empty section is just no format
    expect(f(7, '')).toBe('7')
  })
  test('[$-42A] gives Vietnamese names', () => {
    expect(f(45306, '[$-42A]dddd')).toBe('Thứ Hai')
    expect(f(45305, '[$-42A]dddd')).toBe('Chủ Nhật')
    expect(f(45309, '[$-42A]dddd, d mmmm yyyy')).toBe('Thứ Năm, 18 Tháng Một 2024')
    expect(f(45306, '[$-42A]mmmm')).toBe('Tháng Một')
    expect(f(45306, '[$-42A]mmmm')).not.toBe('January')
    expect(f(45306, '[$-409]mmmm')).toBe('January')
    expect(f(45306, '[$-42A]ddd')).toBe('Th 2')
  })
  test('##0.0E+0 carries when rounding', () => {
    expect(f(0.99999, '##0.0E+0')).toBe('1.0E+0')
    expect(f(999.96, '##0.0E+0')).toBe('1.0E+3')
    expect(f(12345, '##0.0E+0')).toBe('12.3E+3')
    expect(f(0.000123, '##0.0E+0')).toBe('123.0E-6')
    expect(f(9.99999, '0.00E+00')).toBe('1.00E+01')
  })
  test('serial 0 and the weekdays before 1 March 1900', () => {
    expect(f(0, 'dd/mm/yyyy')).toBe('00/01/1900')
    expect(f(1, 'dd/mm/yyyy')).toBe('01/01/1900')
    expect(f(59, 'dd/mm/yyyy')).toBe('28/02/1900')
    expect(f(60, 'dd/mm/yyyy')).toBe('29/02/1900')
    expect(f(61, 'dd/mm/yyyy')).toBe('01/03/1900')
    // Excel: 1/1/1900 is a Sunday (WEEKDAY(1) = 1), as are all the days after it
    const names = [1, 2, 3, 4, 5, 6, 7, 59, 60, 61].map((n) => f(n, 'dddd'))
    expect(names).toEqual(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Tuesday', 'Wednesday', 'Thursday'])
    expect(f(0, 'dddd')).toBe('Saturday')
    expect(f(0.5, 'yyyy-mm-dd hh:mm')).toBe('1900-01-00 12:00')
    // 1904 system: 1/1/1904 is a Friday
    expect(f(0, 'dddd', true)).toBe('Friday')
  })
})

describe('P1-011: the last addition / subtraction near zero is zero', () => {
  test('Excel compensation', () => {
    expect(text('0.1+0.2-0.3')).toBe(0)
    expect(text('1-0.9-0.1')).toBe(0)
    expect(text('0.5-0.4-0.1')).toBe(0)
    expect(text('(0.5-0.4-0.1)')).toBe(0)
    expect(text('0.3-(0.1+0.2)')).toBe(0)
    // not the last operation: Excel keeps the noise (KB 78113)
    expect(Math.abs(text('1*(0.5-0.4-0.1)') as number)).toBeGreaterThan(0)
    // a real small difference is kept
    expect(text('1E-20-2E-20')).toBe(-1e-20)
    expect(text('1+1E-14-1')).toBeCloseTo(1e-14, 15)
    expect(text('0.1+0.2')).toBe(0.30000000000000004)
  })
  test('in a cell: the check cell of a ledger', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: '<row r="1"><c r="A1"><v>0.1</v></c><c r="B1"><v>0.2</v></c><c r="C1"><v>0.3</v></c></row>' }] }))
    await o.set(0, 'D1', '=A1+B1-C1')
    expect(await o.shown(0, 0, 3)).toBe('0')
  })
})

describe('P1-015: ROUND family with extreme digits', () => {
  test('no #NUM!', () => {
    expect(text('ROUND(5,-400)')).toBe(0)
    expect(text('ROUNDDOWN(5,-400)')).toBe(0)
    expect(text('ROUNDUP(5,-400)')).toBe(0)
    expect(text('TRUNC(5,-400)')).toBe(0)
    expect(text('ROUND(1E300,10)')).toBe(1e300)
    expect(text('ROUNDUP(1E300,10)')).toBe(1e300)
    expect(text('ROUNDDOWN(1E300,10)')).toBe(1e300)
    expect(text('TRUNC(1E300,10)')).toBe(1e300)
    expect(text('ROUND(1.5,400)')).toBe(1.5)
    expect(text('ROUND(-5,-400)')).toBe(0)
  })
  test('as before', () => {
    expect(text('ROUND(2.5,0)')).toBe(3)
    expect(text('ROUND(-2.5,0)')).toBe(-3)
    expect(text('ROUND(1234.5678,-2)')).toBe(1200)
    expect(text('ROUNDUP(1.231,2)')).toBe(1.24)
    expect(text('ROUNDDOWN(-1.239,2)')).toBe(-1.23)
    expect(text('TRUNC(8.9)')).toBe(8)
    expect(text('ROUND(2.675,2)')).toBe(2.68)
  })
})

describe('P1-016: COUNTIF / SUMIF with numbers stored as text', () => {
  const p = provider({ S: Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`M${i + 1}`, String(i + 1)])) })
  test('=n and <>n are complements', () => {
    expect(run('COUNTIF(M1:M10,5)', p)).toBe(1)
    expect(run('COUNTIF(M1:M10,"=5")', p)).toBe(1)
    expect(run('COUNTIF(M1:M10,"<>5")', p)).toBe(9)
    expect(run('COUNTIF(M1:M10,"5")', p)).toBe(1)
    expect(run('COUNTIF(M1:M10,5)+COUNTIF(M1:M10,"<>5")', p)).toBe(10)
    expect(run('COUNTIFS(M1:M10,"<>5",M1:M10,"<>6")', p)).toBe(8)
  })
  test('with real numbers and text', () => {
    const q = provider({ S: { A1: 5, A2: 6, A3: 'abc', A4: '5', A5: '' } })
    expect(run('COUNTIF(A1:A5,"<>5")', q)).toBe(3)
    expect(run('SUMIF(A1:A5,"<>5",A1:A5)', q)).toBe(6)
  })
})

describe('P1-017: GCD / LCM with text and empty cells', () => {
  const p = provider({ S: { A1: 5, A2: 'abc', A3: '6', A4: 12 } })
  test('text is #VALUE!, an empty cell is 0', () => {
    expect(err(run('GCD(A1,A2)', p))).toBe('#VALUE!')
    expect(err(run('LCM(A1,A2)', p))).toBe('#VALUE!')
    expect(err(run('GCD(A1:A4)', p))).toBe('#VALUE!')
    expect(err(run('GCD(12,"abc")', p))).toBe('#VALUE!')
    expect(run('GCD(5,B9)', p)).toBe(5)
    expect(run('LCM(5,B9)', p)).toBe(0)
  })
  test('as before', () => {
    expect(run('GCD(12,18)', p)).toBe(6)
    expect(run('GCD(4,"6")', p)).toBe(2)
    expect(run('LCM(4,6)', p)).toBe(12)
    expect(run('GCD(A1,A4)', p)).toBe(1)
    expect(err(run('GCD(-1,2)', p))).toBe('#NUM!')
  })
})

describe('P1-018: text to number', () => {
  test('forms Excel accepts', () => {
    expect(text('"(5)"+0')).toBe(-5)
    expect(text('VALUE("(1,234.50)")')).toBe(-1234.5)
    expect(text('VALUE("Jan 15, 2024")')).toBe(45306)
    expect(text('VALUE("January 15, 2024")')).toBe(45306)
    expect(text('VALUE("15-Jan-2024")')).toBe(45306)
    expect(text('VALUE("15 Jan 2024")')).toBe(45306)
    expect(text('VALUE("15-Jan-24")')).toBe(45306)
    expect(text('VALUE("Jan 2024")')).toBe(45292)
    expect(text('VALUE("1 1/2")')).toBe(1.5)
    expect(text('VALUE("-1 1/2")')).toBe(-1.5)
    expect(text('VALUE("12 3/4")')).toBe(12.75)
    expect(text('DAY("Jan 15, 2024")')).toBe(15)
  })
  test('what is still not a number', () => {
    expect(err(text('VALUE("Foo 15, 2024")'))).toBe('#VALUE!')
    expect(err(text('VALUE("Jan 32, 2024")'))).toBe('#VALUE!')
    expect(err(text('VALUE("1 1/0")'))).toBe('#VALUE!')
    expect(err(text('VALUE("(abc)")'))).toBe('#VALUE!')
    expect(err(text('VALUE("5)")'))).toBe('#VALUE!')
    expect(text('VALUE("1,234.5")')).toBe(1234.5)
    expect(text('VALUE("2024-01-15")')).toBe(45306)
  })
})

describe('P1-019: bounds and error codes', () => {
  test('dates past the last serial are #NUM!', () => {
    expect(err(text('YEAR("20240115")'))).toBe('#NUM!')
    expect(err(text('MONTH("20240115")'))).toBe('#NUM!')
    expect(err(text('DAY(20240115)'))).toBe('#NUM!')
    expect(err(text('WEEKDAY(1E15)'))).toBe('#NUM!')
    expect(text('YEAR(2958465)')).toBe(9999)
    expect(err(text('YEAR(2958466)'))).toBe('#NUM!')
  })
  test('#SPILL!', () => {
    const p = provider({ S: { B2: new FormulaError('#SPILL!' as never), B3: new FormulaError('#N/A') } })
    expect(run('ERROR.TYPE(B2)', p)).toBe(9)
    expect(run('ERROR.TYPE(B3)', p)).toBe(7)
    expect(run('COUNTIF(B2:B3,"#SPILL!")', p)).toBe(1)
    expect(run('ISERROR(B2)', p)).toBe(true)
    expect(run('ERROR.TYPE(#SPILL!)', p)).toBe(9)
  })
  test('INDEX / MID / FACT', () => {
    const p = provider({ S: { C1: 1, C2: 2 } })
    expect(err(run('INDEX(C1:C2,-1)', p))).toBe('#VALUE!')
    expect(err(run('INDEX(C1:C2,3)', p))).toBe('#REF!')
    expect(run('INDEX(C1:C2,2)', p)).toBe(2)
    expect(err(text('MID("abcdef",-3,NA())'))).toBe('#N/A')
    expect(err(text('MID("abcdef",-3,2)'))).toBe('#VALUE!')
    expect(text('MID("abcdef",2,3)')).toBe('bcd')
    expect(err(text('FACT(-0.5)'))).toBe('#NUM!')
    expect(text('FACT(5.9)')).toBe(120)
  })
})

describe('P1-021: text compares by collation', () => {
  const t = (f: string) => text(f)
  test('the cases of the report', () => {
    expect(t('"đ"<"e"')).toBe(true)
    expect(t('"d"<"đ"')).toBe(true)
    expect(t('"á"<"b"')).toBe(true)
    expect(t('"á"<"d"')).toBe(true)
    expect(t('"Ngô"<"Nguyễn"')).toBe(true)
    expect(t('"Nguyễn"<"Ngô"')).toBe(false)
    expect(t('"_"<"1"')).toBe(true)
    expect(t('"a1"<"a_b"')).toBe(false)
    expect(t('"dz"<"đa"')).toBe(true)
  })
  test('case-insensitive, accents count', () => {
    expect(t('"a"="A"')).toBe(true)
    expect(t('"Á"="á"')).toBe(true)
    expect(t('"a"="á"')).toBe(false)
    expect(t('"a"<"B"')).toBe(true)
    expect(t('"B"<"a"')).toBe(false)
    expect(t('"abc"<"abd"')).toBe(true)
    expect(t('"ab"<"abc"')).toBe(true)
    expect(t('"a b"<"ab"')).toBe(true)
    expect(t('"10"<"9"')).toBe(true)
    expect(t('"Z"<"a1"')).toBe(false)
  })
  test('lookups and criteria use it', () => {
    const p = provider({ S: { A1: 'An', A2: 'Đào', A3: 'Em', A4: 'Ngô', A5: 'Nguyễn' } })
    expect(run('MATCH("Ngô",A1:A5,1)', p)).toBe(4)
    expect(run('MATCH("Nguyễn",A1:A5,1)', p)).toBe(5)
    expect(run('MATCH("Đ",A1:A5,1)', p)).toBe(1)
    expect(run('COUNTIF(A1:A5,">Đào")', p)).toBe(3)
    expect(run('COUNTIF(A1:A5,"<Em")', p)).toBe(2)
    expect(run('MATCH("ngô",A1:A5,0)', p)).toBe(4)
  })
  test('plain ASCII stays fast', () => {
    const names = Array.from({ length: 300 }, (_, i) => `Name ${(i * 7919) % 1000} abc`)
    const t0 = performance.now()
    let n = 0
    for (let k = 0; k < 300; k++) for (let j = 0; j < 1000; j++) n += compare(names[k], names[(k + j) % 300])
    expect(n).not.toBe(0.5)
    // 300k comparisons: a few ms here, the bound is generous (the VM is slow)
    expect(performance.now() - t0).toBeLessThan(3000)
    expect(compare('a', 'B')).toBe(-1)
    expect(compare('B', 'b')).toBe(0)
  })
})

describe('P1-022: CHAR / CODE use Windows-1252', () => {
  test('128-159', () => {
    expect(text('CHAR(128)')).toBe('€')
    expect(text('CHAR(149)')).toBe('•')
    expect(text('CHAR(150)')).toBe('–')
    expect(text('CHAR(151)')).toBe('—')
    expect(text('CHAR(153)')).toBe('™')
    expect(text('CHAR(65)')).toBe('A')
    expect(text('CHAR(233)')).toBe('é')
    expect(text('CODE("€")')).toBe(128)
    expect(text('CODE("•")')).toBe(149)
    expect(text('CODE("—")')).toBe(151)
    expect(text('CODE("A")')).toBe(65)
    expect(text('CODE("é")')).toBe(233)
    expect(text('CODE("ồ")')).toBe(7891)
    expect(err(text('CHAR(256)'))).toBe('#VALUE!')
    // every one maps back
    for (let n = 128; n < 160; n++) expect(text(`CODE(CHAR(${n}))`)).toBe(n)
  })
})

describe('P1-023: date cells of type d', () => {
  const styles =
    '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'
  const rowsXml =
    '<row r="1"><c r="A1" t="d"><v>2024-01-15T00:00:00</v></c><c r="B1" t="d" s="1"><v>2024-01-15T00:00:00</v></c></row>' +
    '<row r="2"><c r="A2" t="d"><v>2024-01-16</v></c><c r="B2" t="d"><v>2024-01-16T06:00:00Z</v></c><c r="C2" t="d"><v>not a date</v></c></row>'
  test('formulas see numbers, the cell shows its format', async () => {
    const o = await openBook(await book({ styles, sheets: [{ name: 'S', rows: rowsXml }] }))
    await o.set(0, 'D1', '=SUM(A1:A2)')
    await o.set(0, 'D2', '=ISNUMBER(A1)')
    await o.set(0, 'D3', '=A1+1')
    await o.set(0, 'D4', '=B2-A2')
    await o.set(0, 'D5', '=ISNUMBER(C2)')
    expect(await o.shown(0, 0, 3)).toBe('90613')
    expect(await o.shown(0, 1, 3)).toBe('TRUE')
    expect(await o.shown(0, 2, 3)).toBe('45307')
    expect(await o.shown(0, 3, 3)).toBe('0.25')
    expect(await o.shown(0, 4, 3)).toBe('FALSE')
    // the number format of the cell, else the ISO text as before
    expect(await o.shown(0, 0, 1)).toBe('15/01/2024')
    expect(await o.shown(0, 0, 0)).toBe('2024-01-15')
    expect(await o.shown(0, 1, 0)).toBe('2024-01-16')
    expect(await o.shown(0, 1, 1)).toBe('2024-01-16 06:00:00')
    expect(await o.shown(0, 1, 2)).toBe('not a date')
  })
})

describe('P1-024: whole columns of two sheets', () => {
  const D: Record<string, Value> = {}
  const T: Record<string, Value> = {}
  for (let i = 1; i <= 100; i++) D[`C${i}`] = i
  for (let i = 1; i <= 50; i++) {
    T[`A${i}`] = 2
    T[`Z${i}`] = 1
  }
  const p = provider({ D, T, S: { A1: 1 } })
  const calls: number[] = []
  test('shapes agree', () => {
    expect(run('SUMPRODUCT(D!C:C,D!C:C)', p, 'S')).toBe((100 * 101 * 201) / 6)
    expect(run('SUMPRODUCT(D!C:C,T!A:A)', p, 'S')).toBe(2 * 1275)
    expect(run('SUMPRODUCT(T!A:A,D!C:C)', p, 'S')).toBe(2 * 1275)
    expect(run('SUMIFS(T!Z:Z,D!C:C,">2")', p, 'S')).toBe(48)
    expect(run('COUNTIFS(D!C:C,">2",T!A:A,"<>")', p, 'S')).toBe(48)
    expect(run('SUM(D!C:C*T!A:A)', p, 'S')).toBe(2 * 1275)
    expect(run('SUMIF(D!C:C,">98",T!A:A)', p, 'S')).toBe(0)
    expect(run('SUMPRODUCT((D!C:C>98)*1,T!A:A)', p, 'S')).toBe(0)
  })
  test('a range of a given size is still not a whole column', () => {
    expect(err(run('SUMPRODUCT(D!C1:C100,T!A1:A50)', p, 'S'))).toBe('#VALUE!')
    expect(run('SUMPRODUCT(D!C1:C50,T!A1:A50)', p, 'S')).toBe(2 * 1275)
  })
  test('blank rows are not read', () => {
    let read = 0
    const q = provider({ D, T }, { value: (s, r, c) => (read++, p.value(s, r, c)) })
    expect(run('SUMPRODUCT(D!C:C,T!A:A)', q, 'S')).toBe(2 * 1275)
    expect(read).toBeLessThan(1000)
    calls.push(read)
  })
})

describe('P2-024: a reference to a row past the last of Excel', () => {
  const rowsInSheet = 1_600_001
  const big = (extra: Partial<CellProvider> = {}): CellProvider => ({
    value: (_s, r, c) => (c === 0 ? r + 1 : null),
    bounds: () => ({ rows: rowsInSheet, cols: 1 }),
    ...extra,
  })
  test('is a cell of the sheet', () => {
    expect(run('A1500000*2', big())).toBe(3000000)
    expect(run('A1048577+A1600001', big())).toBe(1048577 + 1600001)
    expect(run('SUM(A1599998:A1600001)', big())).toBe(1599998 + 1599999 + 1600000 + 1600001)
    expect(run('INDEX(A1500000:A1500002,2)', big())).toBe(1500001)
    expect(run('ROW(A1500000)', big())).toBe(1500000)
  })
  test('on a sheet that has no such row it is still not a reference', () => {
    expect(err(run('A1500000*2', provider({ S: { A1: 1 } })))).toBe('#NAME?')
    expect(err(run('A1048577', provider({ S: { A1: 1 } })))).toBe('#NAME?')
    expect(err(run('A1600002', big()))).toBe('#NAME?')
    expect(err(run('A1500000', { value: () => 1 }))).toBe('#NAME?')
  })
  test('whole columns still end where the sheet does', () => {
    let read = 0
    expect(run('SUM(A:A)', big({ value: (_s, r, c) => (read++, c === 0 && (r < 3 || r >= rowsInSheet - 2) ? 1 : null) }))).toBe(5)
    expect(read).toBe(rowsInSheet)
    expect(run('SUM(A:A)', { value: () => 2, bounds: () => ({ rows: 4, cols: 1 }) })).toBe(8)
  })
  test('what a name looked like before is a name still', () => {
    expect(err(run('ZZZZ1+1'))).toBe('#NAME?')
    expect(err(run('A99999999999'))).toBe('#NAME?')
  })
})

describe('N1-012 / N3-006: General fits the column', () => {
  const px = (11 * 4) / 3
  const W = 64 - 6
  test('formatGeneral: 11 characters', () => {
    expect(formatGeneral(1234.5678901234)).toBe('1234.56789')
    expect(formatGeneral(0.123456789012)).toBe('0.123456789')
    expect(formatGeneral(123456.789)).toBe('123456.789')
    expect(formatGeneral(0.1 + 0.2)).toBe('0.3')
    expect(formatGeneral(1e10)).toBe('10000000000')
    expect(formatGeneral(1e11)).toBe('1E+11')
    expect(formatGeneral(123456789012)).toBe('1.23457E+11')
    expect(formatGeneral(1.5e-5)).toBe('1.5E-05')
    expect(formatGeneral(0.00001234)).toBe('1.234E-05')
    expect(formatGeneral(0.0001)).toBe('0.0001')
    expect(formatGeneral(-0.00123)).toBe('-0.00123')
    expect(formatGeneral(1.23456789e-10)).toBe('1.23457E-10')
    expect(formatGeneral(1e100)).toBe('1E+100')
    expect(formatGeneral(0)).toBe('0')
    expect(formatGeneral(-5)).toBe('-5')
    expect(formatGeneral(99999999999)).toBe('99999999999')
  })
  test('formatGeneral shrinks to the room it is given', () => {
    expect(formatGeneral(123456789, 8)).toBe('1.23E+08')
    expect(formatGeneral(1234.5678, 8)).toBe('1234.568')
    expect(formatGeneral(0.123456789, 8)).toBe('0.123457')
    expect(formatGeneral(1e10, 8)).toBe('1E+10')
    expect(formatGeneral(1e10, 4)).toBe('####')
    expect(formatGeneral(9.9999999, 5)).toBe('10')
    expect(formatGeneral(-123456789, 8)).toBe('-1.2E+08')
    expect(formatGeneral(1.6e-5, 5)).toBe('2E-05')
  })
  test('fitNumber in a 64 px column', () => {
    expect(fitNumber('12345678', W, px)).toBe('12345678')
    expect(fitNumber('123456789', W, px)).toBe('1.23E+08')
    expect(fitNumber('-1234567', W, px)).toBe('-1234567')
    expect(fitNumber('10000000000', W, px)).toBe('1E+10')
    expect(fitNumber('1234.5678', W, px)).toBe('1234.568')
    expect(fitNumber('0.123456789', W, px)).toBe('0.123457')
    expect(fitNumber('912345678', W, px)).toBe('9.12E+08')
    expect(fitNumber('1.234E-05', W, px)).toBe('1.23E-05')
    expect(fitNumber('1.23457E+08', W, px)).toBe('1.23E+08')
  })
  test('N3-006: a 7-digit integer in a 61 px column', () => {
    expect(fitNumber('1800000', 61 - 6, px)).toBe('1800000')
  })
  test('### only when not even the scientific form fits', () => {
    expect(fitNumber('123456789', 20, px)).toBe('##')
    expect(fitNumber('123456789', 3, px)).toBe('#')
    expect(fitNumber('1,234,567', W, px)).toBe('########')
  })
})
