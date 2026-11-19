/** The math, text, date, lookup, information, aggregate and Google Sheets functions added for the function coverage. */
import { describe, expect, test } from 'bun:test'
import { colIndex, evaluateArray, evaluateDynamic, evaluateFormula, FormulaError, parseFormula, type CellProvider, type ErrorCode, type PivotInfo, type Value } from '../src/core/xlsx/formula'

type Sheets = Record<string, Record<string, Value>>

const addr = (a: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a)!
  return { r: Number(m[2]) - 1, c: colIndex(m[1]) }
}

function makeProvider(sheets: Sheets, extra: Partial<CellProvider> = {}): CellProvider {
  const cells = new Map<string, Value>()
  const ext = new Map<string, { rows: number; cols: number }>()
  const key = (s: string, r: number, c: number) => `${s.toLowerCase()}!${r},${c}`
  for (const [name, data] of Object.entries(sheets)) {
    let rows = 0
    let cols = 0
    for (const [a, v] of Object.entries(data)) {
      const { r, c } = addr(a)
      cells.set(key(name, r, c), v)
      rows = Math.max(rows, r + 1)
      cols = Math.max(cols, c + 1)
    }
    ext.set(name.toLowerCase(), { rows, cols })
  }
  const provider: CellProvider = {
    value(sheet, r, c) {
      const s = sheet ?? 'Sheet1'
      const v = cells.get(key(s, r, c)) ?? null
      if (typeof v === 'string' && v.startsWith('=')) return evaluateFormula(v.slice(1), { provider, sheet: s, r, c })
      return v
    },
    bounds: (sheet) => ext.get((sheet ?? 'Sheet1').toLowerCase()) ?? { rows: 0, cols: 0 },
    hasSheet: (name) => ext.has(name.toLowerCase()),
    sheetNames: () => Object.keys(sheets),
    now: () => new Date(2024, 0, 15, 13, 30, 0),
    ...extra,
  }
  return provider
}

const E = (code: ErrorCode) => new FormulaError(code)
const DATA: Sheets = {
  Sheet1: {
    A1: 'Name', B1: 'Dept', C1: 'Salary',
    A2: 'Ann', B2: 'Ops', C2: 120,
    A3: 'Bob', B3: 'Dev', C3: 200,
    A4: 'Cid', B4: 'Ops', C4: 80,
    A5: 'Dee', B5: 'Dev', C5: 150,
    E1: 3, E2: 1, E3: E('#DIV/0!'), E4: 2,
  },
  Sheet2: { A1: 10 },
  Sheet3: { A1: 20 },
}
const P = makeProvider(DATA)
const ev = (f: string, r = 9, c = 9, p: CellProvider = P) => evaluateFormula(f, { provider: p, sheet: 'Sheet1', r, c })
const arr = (f: string, p: CellProvider = P) => evaluateDynamic(parseFormula(f), { provider: p, sheet: 'Sheet1', r: 9, c: 9 })
const near = (f: string, x: number, digits = 6) => expect(ev(f) as number).toBeCloseTo(x, digits)

describe('math', () => {
  test('trigonometry', () => {
    near('SIN(PI()/2)', 1)
    near('COS(0)', 1)
    near('TAN(PI()/4)', 1)
    near('ASIN(1)', Math.PI / 2)
    expect(ev('ACOS(2)')).toEqual(E('#NUM!'))
    near('ATAN2(1,1)', Math.PI / 4)
    expect(ev('ATAN2(0,0)')).toEqual(E('#DIV/0!'))
    near('ACOT(2)', 0.463647609)
    expect(ev('COT(0)')).toEqual(E('#DIV/0!'))
    near('SINH(1)', 1.175201194)
    near('ACOSH(1)', 0)
    expect(ev('ATANH(1)')).toEqual(E('#NUM!'))
    near('ACOTH(6)', 0.168236118)
    near('SEC(45)', 1.90359, 4)
    near('CSC(15)', 1.53778, 4)
    near('COTH(2)', 1.03731, 4)
    near('CSCH(1.5)', 0.46964, 4)
    near('DEGREES(PI())', 180)
    near('RADIANS(270)', 4.712389)
    near('SQRTPI(1)', 1.772454)
    expect(ev('SQRTPI(-1)')).toEqual(E('#NUM!'))
    expect(ev('SIN(1E9)')).toEqual(E('#NUM!'))
  })

  test('rounding and counting', () => {
    expect(ev('MROUND(10,3)')).toBe(9)
    expect(ev('MROUND(-10,-3)')).toBe(-9)
    near('MROUND(1.3,0.2)', 1.4)
    expect(ev('MROUND(5,-2)')).toEqual(E('#NUM!'))
    expect(ev('QUOTIENT(5,2)')).toBe(2)
    expect(ev('QUOTIENT(-10,3)')).toBe(-3)
    expect(ev('QUOTIENT(1,0)')).toEqual(E('#DIV/0!'))
    expect(ev('CEILING.PRECISE(4.3)')).toBe(5)
    expect(ev('CEILING.PRECISE(-4.3)')).toBe(-4)
    expect(ev('CEILING.PRECISE(4.3,2)')).toBe(6)
    expect(ev('CEILING.PRECISE(-4.3,-2)')).toBe(-4)
    expect(ev('ISO.CEILING(4.3)')).toBe(5)
    expect(ev('FLOOR.PRECISE(-3.2,-1)')).toBe(-4)
    expect(ev('FLOOR.PRECISE(3.2)')).toBe(3)
    expect(ev('COMBIN(8,2)')).toBe(28)
    expect(ev('COMBINA(4,3)')).toBe(20)
    expect(ev('COMBINA(10,3)')).toBe(220)
    expect(ev('FACTDOUBLE(6)')).toBe(48)
    expect(ev('FACTDOUBLE(7)')).toBe(105)
    expect(ev('MULTINOMIAL(2,3,4)')).toBe(1260)
  })

  test('numerals and bases', () => {
    expect(ev('ROMAN(499,0)')).toBe('CDXCIX')
    expect(ev('ROMAN(499,1)')).toBe('LDVLIV')
    expect(ev('ROMAN(499,2)')).toBe('XDIX')
    expect(ev('ROMAN(499,3)')).toBe('VDIV')
    expect(ev('ROMAN(499,4)')).toBe('ID')
    expect(ev('ROMAN(2024)')).toBe('MMXXIV')
    expect(ev('ROMAN(4000)')).toEqual(E('#VALUE!'))
    expect(ev('ARABIC("LVII")')).toBe(57)
    expect(ev('ARABIC("mcmxii")')).toBe(1912)
    expect(ev('ARABIC("-IV")')).toBe(-4)
    expect(ev('ARABIC("AB")')).toEqual(E('#VALUE!'))
    expect(ev('BASE(7,2)')).toBe('111')
    expect(ev('BASE(100,16)')).toBe('64')
    expect(ev('BASE(15,2,10)')).toBe('0000001111')
    expect(ev('DECIMAL("FF",16)')).toBe(255)
    expect(ev('DECIMAL(111,2)')).toBe(7)
    expect(ev('DECIMAL("zap",36)')).toBe(45745)
    expect(ev('DECIMAL("12",2)')).toEqual(E('#NUM!'))
  })

  test('sums', () => {
    expect(ev('SUMSQ(3,4)')).toBe(25)
    expect(ev('SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})')).toBe(-55)
    expect(ev('SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})')).toBe(521)
    expect(ev('SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})')).toBe(79)
    expect(ev('SUMXMY2({1,2},{1,2,3})')).toEqual(E('#N/A'))
    near('SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666667,-0.001388889})', 0.707103, 5)
    near('PERCENTOF(C2:C3,C2:C5)', 320 / 550)
    expect(ev('COUNTUNIQUE(1,1,2,"a","b","a")')).toBe(4)
  })

  test('SUBTOTAL and AGGREGATE', () => {
    expect(ev('SUBTOTAL(9,C2:C5)')).toBe(550)
    expect(ev('SUBTOTAL(109,C2:C5)')).toBe(550)
    expect(ev('SUBTOTAL(1,C2:C5)')).toBe(137.5)
    expect(ev('SUBTOTAL(3,A1:A5)')).toBe(5)
    expect(ev('SUBTOTAL(2,A1:C5)')).toBe(4)
    expect(ev('SUBTOTAL(4,C2:C5)')).toBe(200)
    expect(ev('SUBTOTAL(12,C2:C5)')).toEqual(E('#VALUE!'))
    expect(ev('SUBTOTAL(9,E1:E4)')).toEqual(E('#DIV/0!'))
    expect(ev('AGGREGATE(9,6,E1:E4)')).toBe(6)
    expect(ev('AGGREGATE(4,6,E1:E4)')).toBe(3)
    expect(ev('AGGREGATE(14,6,E1:E4,2)')).toBe(2)
    expect(ev('AGGREGATE(15,6,E1:E4,1)')).toBe(1)
    expect(ev('AGGREGATE(12,6,E1:E4)')).toBe(2)
    expect(ev('AGGREGATE(9,4,E1:E4)')).toEqual(E('#DIV/0!'))
    near('AGGREGATE(16,6,E1:E4,0.5)', 2)
    // hidden rows are left out by 101-111 and by the options that say so
    const hidden = makeProvider(DATA, { rowHidden: (_s, r, filtered) => r === 2 && !filtered })
    expect(ev('SUBTOTAL(9,C2:C5)', 9, 9, hidden)).toBe(550)
    expect(ev('SUBTOTAL(109,C2:C5)', 9, 9, hidden)).toBe(350)
    expect(ev('AGGREGATE(9,5,C2:C5)', 9, 9, hidden)).toBe(350)
    // the rows a filter hides are left out by both
    const filtered = makeProvider(DATA, { rowHidden: (_s, r) => r === 2 })
    expect(ev('SUBTOTAL(9,C2:C5)', 9, 9, filtered)).toBe(350)
    // a subtotal does not count the subtotals among its cells
    const nested = makeProvider(DATA, { isSubtotal: (_s, r, c) => r === 3 && c === 2 })
    expect(ev('SUBTOTAL(9,C2:C5)', 9, 9, nested)).toBe(470)
  })
})

describe('text', () => {
  test('characters and numbers as text', () => {
    expect(ev('UNICHAR(66)')).toBe('B')
    expect(ev('UNICHAR(128512)')).toBe('😀')
    expect(ev('UNICHAR(0)')).toEqual(E('#VALUE!'))
    expect(ev('UNICODE("B")')).toBe(66)
    expect(ev('UNICODE("😀")')).toBe(128512)
    expect(ev('CLEAN(CHAR(9)&"x"&CHAR(10))')).toBe('x')
    expect(ev('FIXED(1234.567,1)')).toBe('1,234.6')
    expect(ev('FIXED(1234.567,-1)')).toBe('1,230')
    expect(ev('FIXED(-1234.567,-1,TRUE)')).toBe('-1230')
    expect(ev('FIXED(44.332)')).toBe('44.33')
    expect(ev('DOLLAR(1234.567,2)')).toBe('$1,234.57')
    expect(ev('DOLLAR(-1234.567,-2)')).toBe('($1,200)')
    expect(ev('DOLLAR(-0.123,4)')).toBe('($0.1230)')
    expect(ev('DOLLAR(99.888)')).toBe('$99.89')
    near('NUMBERVALUE("2.500,27",",",".")', 2500.27)
    near('NUMBERVALUE("3.5%")', 0.035)
    expect(ev('NUMBERVALUE("")')).toBe(0)
    expect(ev('NUMBERVALUE("1,2,3.4.5")')).toEqual(E('#VALUE!'))
    expect(ev('VALUETOTEXT("a",1)')).toBe('"a"')
    expect(ev('VALUETOTEXT(12)')).toBe('12')
    expect(ev('ARRAYTOTEXT({1,2;"a",TRUE},1)')).toBe('{1,2;"a",TRUE}')
    expect(ev('ARRAYTOTEXT({1,2;"a",TRUE})')).toBe('1, 2, a, TRUE')
    expect(ev('ENCODEURL("a b&c/é")')).toBe('a%20b%26c%2F%C3%A9')
    expect(ev('LENB("abc")')).toBe(3)
    expect(ev('LEFTB("abcdef",2)')).toBe('ab')
    expect(ev('ASC("ＡＢＣ１")')).toBe('ABC1')
    expect(ev('DBCS("ABC")')).toBe('ＡＢＣ')
    expect(ev('JIS("ｶﾞ")')).toBe('ガ')
    expect(ev('ASC("ガ")')).toBe('ｶﾞ')
  })

  test('BAHTTEXT', () => {
    expect(ev('BAHTTEXT(1234)')).toBe('หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน')
    expect(ev('BAHTTEXT(21.5)')).toBe('ยี่สิบเอ็ดบาทห้าสิบสตางค์')
    expect(ev('BAHTTEXT(0)')).toBe('ศูนย์บาทถ้วน')
    expect(ev('BAHTTEXT(11)')).toBe('สิบเอ็ดบาทถ้วน')
    expect(ev('BAHTTEXT(-0.25)')).toBe('ลบยี่สิบห้าสตางค์')
  })

  test('regular expressions', () => {
    expect(ev('REGEXTEST("abc123","\\d+")')).toBe(true)
    expect(ev('REGEXTEST("abc","\\d+")')).toBe(false)
    expect(ev('REGEXTEST("ABC","abc")')).toBe(false)
    expect(ev('REGEXTEST("ABC","abc",1)')).toBe(true)
    expect(ev('REGEXMATCH("abc123","[a-c]+\\d")')).toBe(true)
    expect(ev('REGEXEXTRACT("abc123def45","\\d+")')).toBe('123')
    expect(ev('REGEXEXTRACT("abc","\\d+")')).toEqual(E('#N/A'))
    expect(arr('REGEXEXTRACT("abc123def45","\\d+",1)')).toEqual([['123'], ['45']])
    expect(arr('REGEXEXTRACT("2024-01-15","(\\d+)-(\\d+)-(\\d+)",2)')).toEqual([['2024', '01', '15']])
    expect(ev('REGEXREPLACE("a1b22","\\d+","#")')).toBe('a#b#')
    expect(ev('REGEXREPLACE("a1b22","\\d+","#",2)')).toBe('a1b#')
    expect(ev('REGEXREPLACE("a1b22","\\d+","#",-2)')).toBe('a#b22')
    expect(ev('REGEXREPLACE("john smith","(\\w+) (\\w+)","$2 $1")')).toBe('smith john')
    expect(ev('REGEXTEST("a","(")')).toEqual(E('#VALUE!'))
  })

  test('TEXTBEFORE, TEXTAFTER, TEXTSPLIT, SPLIT, JOIN', () => {
    const t = '"Red riding hood\'s, red hood"'
    expect(ev(`TEXTBEFORE(${t},"hood")`)).toBe('Red riding ')
    expect(ev(`TEXTBEFORE(${t},"hood",2)`)).toBe("Red riding hood's, red ")
    expect(ev(`TEXTBEFORE(${t},"hood",-1)`)).toBe("Red riding hood's, red ")
    expect(ev(`TEXTAFTER(${t},"hood")`)).toBe("'s, red hood")
    expect(ev(`TEXTAFTER(${t},"hood",-1)`)).toBe('')
    expect(ev(`TEXTAFTER(${t},"basket")`)).toEqual(E('#N/A'))
    expect(ev(`TEXTAFTER(${t},"basket",,,,"none")`)).toBe('none')
    expect(ev(`TEXTBEFORE(${t},"RED",1,1)`)).toBe('')
    expect(ev(`TEXTBEFORE(${t},"RED")`)).toEqual(E('#N/A'))
    expect(ev(`TEXTBEFORE(${t},"x",1,0,1)`)).toBe("Red riding hood's, red hood")
    expect(ev(`TEXTAFTER(${t},{"zzz",", "})`)).toBe('red hood')
    expect(ev(`TEXTAFTER(${t},"hood",0)`)).toEqual(E('#VALUE!'))
    expect(arr('TEXTSPLIT("Dakota Lennon Sanchez"," ")')).toEqual([['Dakota', 'Lennon', 'Sanchez']])
    expect(arr('TEXTSPLIT("1,2;3,4",",",";")')).toEqual([
      ['1', '2'],
      ['3', '4'],
    ])
    expect(arr('TEXTSPLIT("a,,b;c",",",";",TRUE)')).toEqual([
      ['a', 'b'],
      ['c', E('#N/A')],
    ])
    expect(arr('TEXTSPLIT("a-b_c",{"-","_"})')).toEqual([['a', 'b', 'c']])
    expect(arr('SPLIT("1,2,,x",",")')).toEqual([[1, 2, 'x']])
    expect(arr('SPLIT("a, b",", ",FALSE)')).toEqual([['a', 'b']])
    expect(ev('JOIN("-",{1,2,3},"a")')).toBe('1-2-3-a')
  })
})

describe('dates', () => {
  test('text to dates, weeks, fractions of a year', () => {
    expect(ev('DATEVALUE("8/22/2011")')).toBe(40777)
    expect(ev('DATEVALUE("22-AUG-2011")')).toBe(40777)
    expect(ev('DATEVALUE("2011/02/23 10:30")')).toBe(40597)
    expect(ev('DATEVALUE("nope")')).toEqual(E('#VALUE!'))
    // without a year: of this one (the provider's clock says 2024)
    expect(ev('DATEVALUE("5-JUL")')).toBe(45478)
    expect(ev('DATEVALUE("7/5")')).toBe(45478)
    near('TIMEVALUE("2:24 AM")', 0.1)
    near('TIMEVALUE("22-Aug-2011 6:35 AM")', 0.2743055556)
    expect(ev('DAYS360(DATE(2011,1,30),DATE(2011,2,1))')).toBe(1)
    expect(ev('DAYS360(DATE(2011,1,1),DATE(2011,12,31))')).toBe(360)
    expect(ev('DAYS360(DATE(2011,1,1),DATE(2011,2,1))')).toBe(30)
    near('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))', 0.58055556)
    near('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)', 0.57650273)
    near('YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)', 0.57808219)
    expect(ev('YEARFRAC(1,2,9)')).toEqual(E('#NUM!'))
    expect(ev('ISOWEEKNUM(DATE(2012,3,9))')).toBe(10)
    expect(ev('ISOWEEKNUM(DATE(2021,1,1))')).toBe(53)
    expect(ev('WEEKNUM(DATE(2012,3,9))')).toBe(10)
    expect(ev('WEEKNUM(DATE(2012,3,9),2)')).toBe(11)
    expect(ev('WEEKNUM(DATE(2021,1,1),21)')).toBe(53)
    expect(ev('WEEKNUM(DATE(2012,3,9),5)')).toEqual(E('#NUM!'))
    near('EPOCHTODATE(1655906568)', 25569 + 1655906568 / 86400)
    near('EPOCHTODATE(1655906568000,2)', 25569 + 1655906568 / 86400)
  })

  test('working days', () => {
    expect(ev('WORKDAY(DATE(2008,10,1),151)')).toBe(39933)
    expect(ev('WORKDAY(DATE(2008,10,1),151,{39778,39786,39834})')).toBe(39938)
    expect(ev('WORKDAY(DATE(2024,1,15),-1)')).toBe(45303)
    expect(ev('NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))')).toBe(22)
    expect(ev('NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))')).toBe(-21)
    expect(ev('NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{"2006/1/2","2006/1/16"})')).toBe(22)
    expect(ev('NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",{"2006/1/2","2006/1/16"})')).toBe(20)
    expect(ev('NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{38719,38733})')).toBe(22)
    expect(ev('NETWORKDAYS(DATE(2006,1,1),DATE(2006,1,31),{"2006/1/2","nope"})')).toEqual(E('#VALUE!'))
    expect(ev('NETWORKDAYS(DATE(2006,1,1),DATE(2006,1,31),"2006/1/2")')).toBe(21)
    expect(ev('NETWORKDAYS.INTL(1,10,"1111111")')).toEqual(E('#VALUE!'))
    expect(ev('WORKDAY.INTL(DATE(2012,1,1),30,0)')).toEqual(E('#NUM!'))
    expect(ev('WORKDAY.INTL(DATE(2012,1,1),90,11)')).toBe(41013)
    expect(ev('WORKDAY.INTL(DATE(2012,1,1),30,17)')).toBe(40944)
  })
})

describe('lookup and information', () => {
  test('ADDRESS, XMATCH, AREAS, TYPE, ISREF, SHEET, SHEETS, CELL', () => {
    expect(ev('ADDRESS(2,3)')).toBe('$C$2')
    expect(ev('ADDRESS(2,3,2)')).toBe('C$2')
    expect(ev('ADDRESS(2,3,2,FALSE)')).toBe('R2C[3]')
    expect(ev('ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")')).toBe("'[Book1]Sheet1'!R2C3")
    expect(ev('ADDRESS(2,3,1,FALSE,"EXCEL SHEET")')).toBe("'EXCEL SHEET'!R2C3")
    expect(ev('ADDRESS(0,3)')).toEqual(E('#VALUE!'))
    expect(ev('XMATCH("Cid",A1:A5)')).toBe(4)
    expect(ev('XMATCH("c*",A1:A5,2)')).toBe(4)
    expect(ev('XMATCH(130,C2:C5,-1)')).toBe(1)
    expect(ev('XMATCH(130,C2:C5,1)')).toBe(4)
    expect(ev('XMATCH("Ops",B1:B5,0,-1)')).toBe(4)
    expect(ev('XMATCH("zzz",A1:A5)')).toEqual(E('#N/A'))
    expect(ev('AREAS((A1:B2,C3))')).toBe(2)
    expect(ev('AREAS(A1:B2)')).toBe(1)
    expect(ev('TYPE(1)')).toBe(1)
    expect(ev('TYPE("a")')).toBe(2)
    expect(ev('TYPE(TRUE)')).toBe(4)
    expect(ev('TYPE(NA())')).toBe(16)
    expect(ev('TYPE({1,2})')).toBe(64)
    expect(ev('TYPE(A1:B2)')).toBe(64)
    expect(ev('ISREF(A1)')).toBe(true)
    expect(ev('ISREF("A1")')).toBe(false)
    expect(ev('SHEET()')).toBe(1)
    expect(ev('SHEET(Sheet3!A1)')).toBe(3)
    expect(ev('SHEET("Sheet2")')).toBe(2)
    expect(ev('SHEET("Nope")')).toEqual(E('#N/A'))
    expect(ev('SHEETS()')).toBe(3)
    expect(ev('SHEETS(Sheet1:Sheet3!A1)')).toBe(3)
    expect(ev('CELL("address",C5)')).toBe('$C$5')
    expect(ev('CELL("address",Sheet2!C5)')).toBe('Sheet2!$C$5')
    expect(ev('CELL("row",A20)')).toBe(20)
    expect(ev('CELL("col",C1)')).toBe(3)
    expect(ev('CELL("contents",C2)')).toBe(120)
    expect(ev('CELL("type",A1)')).toBe('l')
    expect(ev('CELL("type",C2)')).toBe('v')
    expect(ev('CELL("type",Z99)')).toBe('b')
    expect(ev('CELL("nope",A1)')).toEqual(E('#VALUE!'))
    const fmt = makeProvider(DATA, { cellInfo: () => ({ format: '#,##0.00;[Red](#,##0.00)' }) })
    expect(ev('CELL("format",C2)', 9, 9, fmt)).toBe(',2-')
    expect(ev('CELL("format",C2)')).toBe('G')
  })

  test('formulas of cells, 3-D references, names, tables', () => {
    const p = makeProvider(DATA, {
      formula: (_s, r, c) => (r === 0 && c === 4 ? '_xlfn.XLOOKUP(1,A:A,B:B)' : undefined),
      name: (name) => (name.toLowerCase() === 'pay' ? parseFormula('Sheet1!$C$2:$C$5') : name.toLowerCase() === 'double' ? parseFormula('LAMBDA(x,x*2)') : undefined),
      table: (name, _s, r, c) =>
        name === undefined ? (r >= 0 && r <= 4 && c <= 2 ? TABLE : undefined) : name.toLowerCase() === 'staff' ? TABLE : undefined,
    })
    expect(ev('ISFORMULA(E1)', 9, 9, p)).toBe(true)
    expect(ev('ISFORMULA(E2)', 9, 9, p)).toBe(false)
    expect(ev('FORMULATEXT(E1)', 9, 9, p)).toBe('=XLOOKUP(1,A:A,B:B)')
    expect(ev('FORMULATEXT(E2)', 9, 9, p)).toEqual(E('#N/A'))
    expect(ev('SUM(Sheet2:Sheet3!A1)', 9, 9, p)).toBe(30)
    expect(ev("SUM('Sheet1:Sheet3'!A1)", 9, 9, p)).toBe(30)
    expect(ev('SUM(Sheet2:Nope!A1)', 9, 9, p)).toEqual(E('#REF!'))
    expect(ev('SUM(Pay)', 9, 9, p)).toBe(550)
    expect(ev('MAX(pay)*2', 9, 9, p)).toBe(400)
    expect(ev('Double(21)', 9, 9, p)).toBe(42)
    expect(ev('Unknown+1', 9, 9, p)).toEqual(E('#NAME?'))
    expect(ev('SUM(Staff[Salary])', 9, 9, p)).toBe(550)
    expect(ev('ROWS(Staff[#All])', 9, 9, p)).toBe(5)
    expect(ev('Staff[[#Headers],[Dept]]', 9, 9, p)).toBe('Dept')
    expect(ev('COUNTA(Staff[[Name]:[Dept]])', 9, 9, p)).toBe(8)
    expect(ev('Staff[@Salary]*2', 2, 9, p)).toBe(400)
    expect(ev('[@Salary]*2', 3, 2, p)).toBe(160)
    expect(ev('Staff[@Salary]', 9, 9, p)).toEqual(E('#VALUE!'))
    expect(ev('SUM(Staff[Nope])', 9, 9, p)).toEqual(E('#REF!'))
    expect(ev('COUNTA(Staff)', 9, 9, p)).toBe(12)
    expect(ev('Staff[[#Totals],[Salary]]', 9, 9, p)).toEqual(E('#REF!'))
  })

  test('a space between two references is the cells both have', () => {
    const p = makeProvider(DATA, { name: (name) => ({ ops: parseFormula('Sheet1!$A$2:$C$2'), pay: parseFormula('Sheet1!$C$1:$C$5') })[name.toLowerCase()] })
    expect(ev('SUM(A1:C3 C2:C9)')).toBe(320)
    expect(ev('A2:C2 C1:C5')).toBe(120)
    expect(ev('SUM(A2:C4 B3:C9 C1:C3)')).toBe(200)
    expect(ev('A1:A3 C1:C3')).toEqual(E('#NULL!'))
    expect(ev('ROWS(A1:C5 (B2:B9))')).toBe(4)
    expect(ev('AREAS(B2:D4 B2)')).toBe(1)
    expect(ev('Ops Pay', 9, 9, p)).toBe(120)
    expect(ev('SUM( A1:C3 , C4 )')).toBe(400)
    expect(ev('A1:A3 Sheet2!A1:A3')).toEqual(E('#VALUE!'))
    expect(() => parseFormula('1 2')).toThrow()
    expect(() => parseFormula('A1 2')).toThrow()
  })

  test('GETPIVOTDATA', () => {
    // a pivot table at G1: rows by Dept, columns by Year, the sum of Salary
    const cells: Sheets = {
      Sheet1: {
        G1: 'Sum of Salary', H1: 'Year',
        G2: 'Dept', H2: 2023, I2: 2024, J2: 'Grand Total',
        G3: 'Dev', H3: 100, I3: 250, J3: 350,
        G4: 'Ops', H4: 60, I4: 140, J4: 200,
        G5: 'Grand Total', H5: 160, I5: 390, J5: 550,
      },
    }
    const pivot: PivotInfo = {
      sheet: 'Sheet1', r0: 0, c0: 6, r1: 4, c1: 9, firstDataRow: 2, firstDataCol: 1,
      rowFields: ['Dept'], colFields: ['Year'], dataFields: [{ name: 'Sum of Salary', field: 'Salary' }],
      rows: [{ type: 'data', data: 0, items: ['Dev'] }, { type: 'data', data: 0, items: ['Ops'] }, { type: 'grand', data: 0, items: [] }],
      cols: [{ type: 'data', data: 0, items: [2023] }, { type: 'data', data: 0, items: [2024] }, { type: 'grand', data: 0, items: [] }],
    }
    const p = makeProvider(cells, { pivot: (_s, r, c) => (r <= 4 && c >= 6 && c <= 9 ? pivot : undefined) })
    expect(ev('GETPIVOTDATA("Salary",G1)', 9, 0, p)).toBe(550)
    expect(ev('GETPIVOTDATA("Sum of Salary",G1,"Dept","Ops")', 9, 0, p)).toBe(200)
    expect(ev('GETPIVOTDATA("Salary",G1,"dept","dev","Year",2024)', 9, 0, p)).toBe(250)
    expect(ev('GETPIVOTDATA("Salary",G1,"Year",2023)', 9, 0, p)).toBe(160)
    expect(ev('GETPIVOTDATA("Salary",G1,"Dept","HR")', 9, 0, p)).toEqual(E('#REF!'))
    expect(ev('GETPIVOTDATA("Bonus",G1)', 9, 0, p)).toEqual(E('#REF!'))
    expect(ev('GETPIVOTDATA("Salary",A1)', 9, 0, p)).toEqual(E('#REF!'))
  })

  test('ISDATE, ISEMAIL, ISURL, FILTERXML', () => {
    expect(ev('ISDATE("7/20/1969")')).toBe(true)
    expect(ev('ISDATE("hello")')).toBe(false)
    expect(ev('ISDATE(100)')).toBe(false)
    expect(ev('ISEMAIL("ann@example.com")')).toBe(true)
    expect(ev('ISEMAIL("ann@example")')).toBe(false)
    expect(ev('ISURL("https://example.com/x?y=1")')).toBe(true)
    expect(ev('ISURL("example.com")')).toBe(true)
    expect(ev('ISURL("not a url")')).toBe(false)
    const xml = '"<a><b>1</b><b>two</b><c x=\'5\'><d>9</d></c></a>"'
    expect(arr(`FILTERXML(${xml},"//b")`)).toEqual([[1], ['two']])
    expect(arr(`FILTERXML(${xml},"/a/b[2]")`)).toEqual([['two']])
    expect(arr(`FILTERXML(${xml},"//c/@x")`)).toEqual([[5]])
    expect(arr(`FILTERXML(${xml},"/a/c[@x='5']/d")`)).toEqual([[9]])
    expect(arr(`FILTERXML(${xml},"//b[.='two']")`)).toEqual([['two']])
    expect(ev(`FILTERXML(${xml},"//zzz")`)).toEqual(E('#VALUE!'))
  })
})

const TABLE = { name: 'Staff', sheet: 'Sheet1', r0: 0, c0: 0, r1: 4, c1: 2, headers: 1, totals: 0, columns: ['Name', 'Dept', 'Salary'] }

describe('EUROCONVERT', () => {
  test('the examples of its help page', () => {
    expect(ev('EUROCONVERT(1.2,"DEM","EUR")')).toBe(0.61)
    expect(ev('EUROCONVERT(1,"FRF","EUR",TRUE,3)')).toBe(0.152)
    expect(ev('EUROCONVERT(1,"FRF","EUR",FALSE,3)')).toBe(0.15)
    near('EUROCONVERT(1,"FRF","DEM",TRUE,3)', 0.29728616, 8)
    expect(ev('EUROCONVERT(100,"EUR","ITL")')).toBe(193627)
    expect(ev('EUROCONVERT(5,"DEM","dem")')).toBe(5)
    expect(ev('EUROCONVERT(1,"USD","EUR")')).toEqual(E('#VALUE!'))
    expect(ev('EUROCONVERT(1,"FRF","DEM",TRUE,2)')).toEqual(E('#VALUE!'))
  })
})

describe('Google Sheets functions', () => {
  test('the operators as functions, the converters', () => {
    expect(ev('ADD(2,3)')).toBe(5)
    expect(ev('MINUS(2,3)')).toBe(-1)
    expect(ev('MULTIPLY(2,3)')).toBe(6)
    expect(ev('DIVIDE(1,4)')).toBe(0.25)
    expect(ev('DIVIDE(1,0)')).toEqual(E('#DIV/0!'))
    expect(ev('POW(2,10)')).toBe(1024)
    expect(ev('EQ("a","A")')).toBe(true)
    expect(ev('NE(1,2)')).toBe(true)
    expect(ev('GT(2,1)')).toBe(true)
    expect(ev('GTE(2,2)')).toBe(true)
    expect(ev('LT(2,1)')).toBe(false)
    expect(ev('LTE(2,2)')).toBe(true)
    expect(ev('UMINUS(4)')).toBe(-4)
    expect(ev('UPLUS(4)')).toBe(4)
    expect(ev('UNARY_PERCENT(50)')).toBe(0.5)
    expect(ev('ISBETWEEN(5,1,5)')).toBe(true)
    expect(ev('ISBETWEEN(5,1,5,TRUE,FALSE)')).toBe(false)
    expect(ev('ISBETWEEN(0,1,5)')).toBe(false)
    expect(ev('TO_TEXT(1.5)')).toBe('1.5')
    expect(ev('TO_PURE_NUMBER(0.5)')).toBe(0.5)
    expect(ev('TO_DATE(45000)')).toBe(45000)
  })

  test('QUERY', () => {
    expect(arr('QUERY(A1:C5,"select A, C where C > 100 order by C desc")')).toEqual([
      ['Name', 'Salary'],
      ['Bob', 200],
      ['Dee', 150],
      ['Ann', 120],
    ])
    expect(arr('QUERY(A1:C5,"select B, sum(C) group by B")')).toEqual([
      ['Dept', 'sum Salary'],
      ['Dev', 350],
      ['Ops', 200],
    ])
    expect(arr('QUERY(A1:C5,"select count(A), avg(C), max(C) label count(A) \'n\'")')).toEqual([
      ['n', 'avg Salary', 'max Salary'],
      [4, 137.5, 200],
    ])
    expect(arr('QUERY(A1:C5,"select Col1 where Col2 = \'Ops\' and not Col3 < 100",1)')).toEqual([['Name'], ['Ann']])
    expect(arr('QUERY(A2:C5,"select A where B starts with \'D\' limit 1 offset 1",0)')).toEqual([['Dee']])
    expect(arr('QUERY(A1:C5,"select A, C*2 where A like \'_o%\' or A contains \'nn\'")')).toEqual([
      ['Name', ''],
      ['Ann', 240],
      ['Bob', 400],
    ])
    expect(arr('QUERY(A1:C5,"select *")').length).toBe(5)
    expect(ev('QUERY(A1:C5,"select Z")')).toEqual(E('#VALUE!'))
    expect(ev('QUERY(A1:C5,"select A pivot B")')).toEqual(E('#VALUE!'))
    expect(ev('QUERY(A1:C5,"select A, sum(C)")')).toEqual(E('#VALUE!'))
  })

  test('FLATTEN, ARRAY_CONSTRAIN, SORTN, ARRAYFORMULA', () => {
    expect(arr('FLATTEN({1,2;3,4},5)')).toEqual([[1], [2], [3], [4], [5]])
    expect(arr('ARRAY_CONSTRAIN({1,2,3;4,5,6;7,8,9},2,2)')).toEqual([
      [1, 2],
      [4, 5],
    ])
    expect(arr('SORTN({3;1;2;1},2)')).toEqual([[1], [1]])
    expect(arr('SORTN({3;1;2;1},2,2)')).toEqual([[1], [2]])
    expect(arr('SORTN({3;1;2;1},1,1)')).toEqual([[1], [1]])
    expect(arr('SORTN(A2:C5,2,0,3,FALSE)')).toEqual([
      ['Bob', 'Dev', 200],
      ['Dee', 'Dev', 150],
    ])
    expect(arr('ARRAYFORMULA(C2:C5*2)')).toEqual([[240], [400], [160], [300]])
  })
})

describe('FORECAST.ETS', () => {
  const series: Record<string, Value> = {}
  // a line with a pattern of four steps on it
  const season = [3, -1, -4, 2]
  for (let i = 0; i < 24; i++) {
    series[`A${i + 1}`] = i + 1
    series[`B${i + 1}`] = 10 + 2 * i + season[i % 4]
    series[`C${i + 1}`] = 5 + 3 * i
  }
  const p = makeProvider({ Sheet1: series })
  const f = (text: string) => evaluateFormula(text, { provider: p, sheet: 'Sheet1', r: 0, c: 9 })

  test('a line is carried on; a season is found and followed', () => {
    expect(f('FORECAST.ETS(25,C1:C24,A1:A24)') as number).toBeCloseTo(77, 4)
    expect(f('FORECAST.ETS(26,C1:C24,A1:A24)') as number).toBeCloseTo(80, 4)
    expect(f('FORECAST.ETS.SEASONALITY(B1:B24,A1:A24)')).toBe(4)
    expect(f('FORECAST.ETS.SEASONALITY(C1:C24,A1:A24)')).toBe(0)
    expect(f('FORECAST.ETS(25,B1:B24,A1:A24)') as number).toBeCloseTo(10 + 48 + 3, 2)
    expect(f('FORECAST.ETS(26,B1:B24,A1:A24,4)') as number).toBeCloseTo(10 + 50 - 1, 2)
    expect(f('FORECAST.ETS.STAT(B1:B24,A1:A24,8)')).toBe(1)
    expect(f('FORECAST.ETS.STAT(B1:B24,A1:A24,7)') as number).toBeLessThan(0.5)
    expect(f('FORECAST.ETS.CONFINT(25,B1:B24,A1:A24)') as number).toBeGreaterThanOrEqual(0)
    expect(f('FORECAST.ETS(20,B1:B24,A1:A24)')).toEqual(E('#NUM!'))
    expect(f('FORECAST.ETS(25,B1:B24,A1:A20)')).toEqual(E('#N/A'))
    expect(f('FORECAST.ETS.CONFINT(25,B1:B24,A1:A24,1.5)')).toEqual(E('#NUM!'))
  })
})

describe('array results', () => {
  test('evaluateArray repeats and pads as an array formula does', () => {
    expect(evaluateArray(parseFormula('SEQUENCE(2)'), { provider: P, sheet: 'Sheet1', r: 9, c: 9 }, 3, 1)).toEqual([[1], [2], [E('#N/A')]])
  })
})
