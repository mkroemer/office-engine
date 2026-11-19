import { describe, expect, test } from 'bun:test'
import {
  colIndex,
  evaluateFormula,
  FormulaError,
  type CellProvider,
  type ErrorCode,
  type Value,
} from '../src/core/xlsx/formula'

// ---------------------------------------------------------------- in-memory provider

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
    now: () => new Date(2024, 0, 15, 13, 30, 0),
    ...extra,
  }
  return provider
}

const E = (code: ErrorCode) => new FormulaError(code)

const P = makeProvider({
  Sheet1: { A1: '3+4i', A2: '5-3i', A3: null, B1: '1+j', B2: 1, B3: 2, C1: '1+i', C2: 'abc', M1: E('#DIV/0!') },
})

function ev(f: string): Value {
  const { r, c } = addr('J20')
  return evaluateFormula(f, { provider: P, sheet: 'Sheet1', r, c })
}

const txt = (c: { re: number; im: number }) => `${c.re}${c.im < 0 ? '-' : '+'}${Math.abs(c.im)}i`

/** parse "a+bi" text produced by the engine */
function cx(v: Value): { re: number; im: number } {
  const s = String(v)
  const m = /^(?:([+-]?[\d.]+(?:E[+-]\d+)?)(?=[+-]|$))?(?:([+-]?)([\d.]+(?:E[+-]\d+)?)?[ij])?$/.exec(s)
  if (!m) throw new Error(`not complex: ${s}`)
  const im = /[ij]$/.test(s) ? (m[2] === '-' ? -1 : 1) * (m[3] === undefined ? 1 : Number(m[3])) : 0
  return { re: m[1] === undefined ? 0 : Number(m[1]), im }
}

const close = (a: number, b: number, tol = 1e-12) => Math.abs(a - b) <= tol * Math.max(1, Math.abs(b))
function expectCx(f: string, re: number, im: number, tol = 1e-12) {
  const r = cx(ev(f))
  expect(close(r.re, re, tol) && close(r.im, im, tol)).toBe(true)
}
function expectNum(f: string, n: number, tol = 1e-9) {
  const v = ev(f)
  expect(typeof v).toBe('number')
  expect(close(v as number, n, tol)).toBe(true)
}

describe('base conversion (Microsoft examples)', () => {
  test('BIN2DEC', () => {
    expect(ev('BIN2DEC(1100100)')).toBe(100)
    expect(ev('BIN2DEC(1111111111)')).toBe(-1)
    expect(ev('BIN2DEC("1111111111")')).toBe(-1)
    expect(ev('BIN2DEC(12)')).toEqual(E('#NUM!'))
    expect(ev('BIN2DEC(11111111111)')).toEqual(E('#NUM!'))
    expect(ev('BIN2DEC(TRUE)')).toEqual(E('#VALUE!'))
    expect(ev('BIN2DEC(C2)')).toEqual(E('#NUM!'))
  })
  test('BIN2HEX', () => {
    expect(ev('BIN2HEX(11111011,4)')).toBe('00FB')
    expect(ev('BIN2HEX(1110)')).toBe('E')
    expect(ev('BIN2HEX(1111111111)')).toBe('FFFFFFFFFF')
    expect(ev('BIN2HEX(1111111111,2)')).toBe('FFFFFFFFFF')
    expect(ev('BIN2HEX(11111011,1)')).toEqual(E('#NUM!'))
    expect(ev('BIN2HEX(1110,0)')).toEqual(E('#NUM!'))
    expect(ev('BIN2HEX(1110,"x")')).toEqual(E('#VALUE!'))
  })
  test('BIN2OCT', () => {
    expect(ev('BIN2OCT(1001,3)')).toBe('011')
    expect(ev('BIN2OCT(1100100)')).toBe('144')
    expect(ev('BIN2OCT(1111111111)')).toBe('7777777777')
  })
  test('DEC2BIN', () => {
    expect(ev('DEC2BIN(9,4)')).toBe('1001')
    expect(ev('DEC2BIN(9,3)')).toEqual(E('#NUM!'))
    expect(ev('DEC2BIN(-100)')).toBe('1110011100')
    expect(ev('DEC2BIN(511)')).toBe('111111111')
    expect(ev('DEC2BIN(512)')).toEqual(E('#NUM!'))
    expect(ev('DEC2BIN(-512)')).toBe('1000000000')
    expect(ev('DEC2BIN(-513)')).toEqual(E('#NUM!'))
    expect(ev('DEC2BIN(9.9)')).toBe('1001')
    expect(ev('DEC2BIN("x")')).toEqual(E('#VALUE!'))
  })
  test('DEC2HEX', () => {
    expect(ev('DEC2HEX(100,4)')).toBe('0064')
    expect(ev('DEC2HEX(-54)')).toBe('FFFFFFFFCA')
    expect(ev('DEC2HEX(28)')).toBe('1C')
    expect(ev('DEC2HEX(64,1)')).toEqual(E('#NUM!'))
    expect(ev('DEC2HEX(549755813887)')).toBe('7FFFFFFFFF')
    expect(ev('DEC2HEX(549755813888)')).toEqual(E('#NUM!'))
  })
  test('DEC2OCT', () => {
    expect(ev('DEC2OCT(58,3)')).toBe('072')
    expect(ev('DEC2OCT(-100)')).toBe('7777777634')
    expect(ev('DEC2OCT(536870911)')).toBe('3777777777')
    expect(ev('DEC2OCT(536870912)')).toEqual(E('#NUM!'))
  })
  test('HEX2BIN', () => {
    expect(ev('HEX2BIN("F",8)')).toBe('00001111')
    expect(ev('HEX2BIN("B7")')).toBe('10110111')
    expect(ev('HEX2BIN("FFFFFFFFFF")')).toBe('1111111111')
    expect(ev('HEX2BIN("FFFFFFFE00")')).toBe('1000000000')
    expect(ev('HEX2BIN("200")')).toEqual(E('#NUM!'))
    expect(ev('HEX2BIN("G")')).toEqual(E('#NUM!'))
  })
  test('HEX2DEC', () => {
    expect(ev('HEX2DEC("A5")')).toBe(165)
    expect(ev('HEX2DEC("a5")')).toBe(165)
    expect(ev('HEX2DEC("FFFFFFFF5B")')).toBe(-165)
    expect(ev('HEX2DEC("3DA408B9")')).toBe(1034160313)
    expect(ev('HEX2DEC("FFFFFFFFFFF")')).toEqual(E('#NUM!'))
  })
  test('HEX2OCT', () => {
    expect(ev('HEX2OCT("F",3)')).toBe('017')
    expect(ev('HEX2OCT("3B4E")')).toBe('35516')
    expect(ev('HEX2OCT("FFFFFFFF00")')).toBe('7777777400')
    expect(ev('HEX2OCT("FFFFFFFFFF")')).toBe('7777777777')
    expect(ev('HEX2OCT("1FFFFFFF")')).toBe('3777777777')
    expect(ev('HEX2OCT("20000000")')).toEqual(E('#NUM!'))
  })
  test('OCT2BIN / OCT2DEC / OCT2HEX', () => {
    expect(ev('OCT2BIN(3,3)')).toBe('011')
    expect(ev('OCT2BIN(7777777000)')).toBe('1000000000')
    expect(ev('OCT2DEC(54)')).toBe(44)
    expect(ev('OCT2DEC(7777777533)')).toBe(-165)
    expect(ev('OCT2DEC(8)')).toEqual(E('#NUM!'))
    expect(ev('OCT2HEX(100,4)')).toBe('0040')
    expect(ev('OCT2HEX(7777777533)')).toBe('FFFFFFFF5B')
  })
  test('round trips and arrays', () => {
    for (const n of [-512, -100, -1, 0, 1, 77, 511]) {
      expect(ev(`BIN2DEC(DEC2BIN(${n}))`)).toBe(n)
      expect(ev(`HEX2DEC(DEC2HEX(${n}))`)).toBe(n)
      expect(ev(`OCT2DEC(DEC2OCT(${n}))`)).toBe(n)
      expect(ev(`OCT2DEC(HEX2OCT(BIN2HEX(DEC2BIN(${n}))))`)).toBe(n)
    }
    expect(ev('SUM(HEX2DEC({"A","B"}))')).toBe(21)
  })
})

describe('bit functions', () => {
  test('Microsoft examples', () => {
    expect(ev('BITAND(1,5)')).toBe(1)
    expect(ev('BITAND(13,25)')).toBe(9)
    expect(ev('BITOR(23,10)')).toBe(31)
    expect(ev('BITXOR(5,3)')).toBe(6)
    expect(ev('BITLSHIFT(4,2)')).toBe(16)
    expect(ev('BITRSHIFT(13,2)')).toBe(3)
  })
  test('limits', () => {
    expect(ev('BITAND(2^48-1,2^48-1)')).toBe(2 ** 48 - 1)
    expect(ev('BITAND(2^48,1)')).toEqual(E('#NUM!'))
    expect(ev('BITAND(-1,1)')).toEqual(E('#NUM!'))
    expect(ev('BITOR(1.5,1)')).toEqual(E('#NUM!'))
    expect(ev('BITXOR("a",1)')).toEqual(E('#VALUE!'))
    expect(ev('BITLSHIFT(1,47)')).toBe(2 ** 47)
    expect(ev('BITLSHIFT(1,48)')).toEqual(E('#NUM!'))
    expect(ev('BITLSHIFT(1,54)')).toEqual(E('#NUM!'))
    expect(ev('BITLSHIFT(8,-2)')).toBe(2)
    expect(ev('BITRSHIFT(2^47,47)')).toBe(1)
    expect(ev('BITRSHIFT(1,54)')).toEqual(E('#NUM!'))
    expect(ev('BITRSHIFT(8,-1)')).toBe(16)
    expect(ev('BITRSHIFT(1,53)')).toBe(0)
  })
})

describe('complex numbers', () => {
  test('COMPLEX', () => {
    expect(ev('COMPLEX(3,4)')).toBe('3+4i')
    expect(ev('COMPLEX(3,4,"j")')).toBe('3+4j')
    expect(ev('COMPLEX(0,1)')).toBe('i')
    expect(ev('COMPLEX(0,-1)')).toBe('-i')
    expect(ev('COMPLEX(1,0)')).toBe('1')
    expect(ev('COMPLEX(0,0)')).toBe('0')
    expect(ev('COMPLEX(1,-1)')).toBe('1-i')
    expect(ev('COMPLEX(0,2.5)')).toBe('2.5i')
    expect(ev('COMPLEX(-3,-4)')).toBe('-3-4i')
    expect(ev('COMPLEX(3,4,"k")')).toEqual(E('#VALUE!'))
    expect(ev('COMPLEX(3,4,"I")')).toEqual(E('#VALUE!'))
    expect(ev('COMPLEX("a",4)')).toEqual(E('#VALUE!'))
  })
  test('IMABS / IMAGINARY / IMREAL / IMARGUMENT / IMCONJUGATE', () => {
    expect(ev('IMABS("5+12i")')).toBe(13)
    expect(ev('IMAGINARY("3+4i")')).toBe(4)
    expect(ev('IMAGINARY("i")')).toBe(1)
    expect(ev('IMAGINARY("-i")')).toBe(-1)
    expect(ev('IMAGINARY(5)')).toBe(0)
    expect(ev('IMREAL("6-9i")')).toBe(6)
    expect(ev('IMREAL("9i")')).toBe(0)
    expectNum('IMARGUMENT("3+4i")', 0.927295218)
    expect(ev('IMARGUMENT(0)')).toEqual(E('#DIV/0!'))
    expect(ev('IMCONJUGATE("3+4i")')).toBe('3-4i')
    expect(ev('IMCONJUGATE("3-4j")')).toBe('3+4j')
  })
  test('parsing and errors', () => {
    expect(ev('IMABS("abc")')).toEqual(E('#NUM!'))
    expect(ev('IMABS("3+4k")')).toEqual(E('#NUM!'))
    expect(ev('IMABS("3 + 4i")')).toEqual(E('#NUM!'))
    expect(ev('IMABS(TRUE)')).toEqual(E('#VALUE!'))
    expect(ev('IMREAL("1E2+3e-1i")')).toBe(100)
    expect(ev('IMAGINARY("1E2+3e-1i")')).toBe(0.3)
    expect(ev('IMAGINARY("+4i")')).toBe(4)
    expect(ev('IMABS(A1)')).toBe(5)
    expect(ev('IMABS(A3)')).toBe(0)
  })
  test('Microsoft examples (trig, hyperbolic)', () => {
    expectCx('IMCOS("1+i")', 0.833730025131149, -0.988897705762865)
    expectCx('IMCOSH("4+3i")', -27.0349456030742, 3.85115333481178)
    expectCx('IMCOT("4+3i")', 0.00490118239430447, -0.999266927805902)
    expectCx('IMCSC("4+3i")', -0.0754898329158637, 0.0648774713706355)
    expectCx('IMSIN("3+4i")', 3.85373803791938, -27.0168132580039)
    expectCx('IMSINH("3+4i")', -6.548120040911, -7.61923172032141)
    // csch(3+4i) = 1 / sinh(3+4i)
    expectCx('IMCSCH("3+4i")', -0.0648774713706355, 0.0754898329158637)
    expectCx('IMTAN("1+i")', 0.271752585319512, 1.08392332733869)
  })
  test('Microsoft examples (algebra)', () => {
    expect(ev('IMDIV("-238+240i","10+24i")')).toBe('5+12i')
    expectCx('IMEXP("1+i")', 1.4686939399158, 2.28735528717884)
    expectCx('IMLN("3+4i")', 1.6094379124341, 0.927295218001612)
    expectCx('IMLOG10("3+4i")', 0.698970004336019, 0.402719196273373)
    expectCx('IMLOG2("3+4i")', 2.32192809488736, 1.33780421245098)
    expect(ev('IMPOWER("2+3i",3)')).toBe('-46+9i')
    expect(ev('IMPRODUCT("3+4i","5-3i")')).toBe('27+11i')
    expectCx('IMSQRT("1+i")', 1.09868411346781, 0.455089860562227)
    expect(ev('IMSUB("13+4i","5+3i")')).toBe('8+i')
    expect(ev('IMSUM("3+4i","5-3i")')).toBe('8+i')
  })
  test('Google Sheets examples (IMLOG, IMCOTH, IMTANH)', () => {
    expectCx('IMLOG("1+i",3.5)', 0.276647377832556, 0.626932774314643)
    expectCx('IMLOG(COMPLEX(25,34),2.3)', 4.49324546771284, 1.12470086031758)
    expect(ev('IMLOG(100,10)')).toBe('2')
    expect(ev('IMLOG("1+i",1)')).toEqual(E('#NUM!'))
    expect(ev('IMLOG("1+i",0)')).toEqual(E('#NUM!'))
    expectCx('IMCOTH(COMPLEX(4,1))', 0.999720649533931, -0.000609900253822228)
    expectCx('IMCOTH(3.5)', 1.00182542850644, 0)
    expectCx('IMCOTH("3+2i")', 0.996757796569358, 0.00373971037633696)
    expectCx('IMTANH(COMPLEX(4,1))', 1.00027905623447, 0.00061024092137626)
    expectCx('IMTANH(3.5)', 0.998177897611199, 0)
    expectCx('IMTANH("3+2i")', 1.00323862735361, -0.00376402564150425)
  })
  test('identities', () => {
    // sin^2 + cos^2 = 1 for z = 0.7-1.3i
    const s = ev('IMPOWER(IMSIN("0.7-1.3i"),2)')
    const c = ev('IMPOWER(IMCOS("0.7-1.3i"),2)')
    const sum = cx(ev(`IMSUM("${s}","${c}")`))
    expect(close(sum.re, 1) && close(sum.im, 0)).toBe(true)
    // exp(ln z) = z
    expectCx('IMEXP(IMLN("3-2i"))', 3, -2)
    // sqrt(z)^2 = z, also in the other half-planes
    for (const z of ['3+4i', '-3+4i', '-3-4i', '3-4i', '-2', '5']) {
      const r = cx(ev(`IMPOWER(IMSQRT("${z}"),2)`))
      const want = cx(z)
      expect(close(r.re, want.re) && close(r.im, want.im)).toBe(true)
    }
    // sec = 1/cos, csc = 1/sin, cot = cos/sin, sech, tanh
    for (const f of ['COS', 'SIN']) {
      const inv = cx(ev(`IM${f === 'COS' ? 'SEC' : 'CSC'}("1.2+0.4i")`))
      const prod = cx(ev(`IMPRODUCT(IM${f}("1.2+0.4i"),"${txt(inv)}")`))
      expect(close(prod.re, 1) && close(prod.im, 0)).toBe(true)
    }
    const sech = cx(ev('IMSECH("0.5+0.3i")'))
    const cosh = cx(ev('IMCOSH("0.5+0.3i")'))
    const p = cx(ev(`IMPRODUCT("${txt(cosh)}","${txt(sech)}")`))
    expect(close(p.re, 1) && close(p.im, 0)).toBe(true)
    const cs = cx(ev('IMCSCH("0.5+0.3i")'))
    const sh = cx(ev('IMSINH("0.5+0.3i")'))
    const p2 = cx(ev(`IMPRODUCT("${txt(sh)}","${txt(cs)}")`))
    expect(close(p2.re, 1) && close(p2.im, 0)).toBe(true)
    // tanh = sinh/cosh, coth = 1/tanh
    const th = cx(ev('IMTANH("0.5+0.3i")'))
    const ct = cx(ev('IMCOTH("0.5+0.3i")'))
    const p3 = cx(ev(`IMPRODUCT("${txt(th)}","${txt(ct)}")`))
    expect(close(p3.re, 1) && close(p3.im, 0)).toBe(true)
    // log bases
    expectCx('IMLOG10("100")', 2, 0)
    expectCx('IMLOG2("8")', 3, 0)
    // fractional powers
    expectCx('IMPOWER("4",0.5)', 2, 0)
    expectCx('IMPOWER("2+3i",-1)', 2 / 13, -3 / 13)
    expect(ev('IMPOWER("0",2)')).toBe('0')
    expect(ev('IMPOWER("0",0)')).toEqual(E('#NUM!'))
    expect(ev('IMPOWER("1+i","x")')).toEqual(E('#VALUE!'))
  })
  test('suffix handling', () => {
    expect(ev('IMSUM("3+4j","1")')).toBe('4+4j')
    expect(ev('IMSUM("1","3+4j")')).toBe('4+4j')
    expect(ev('IMSUM("3+4i","1+j")')).toEqual(E('#VALUE!'))
    expect(ev('IMPRODUCT("2j","3j")')).toBe('-6')
    expect(ev('IMDIV("4+2j","2")')).toBe('2+j')
    expect(ev('IMDIV("1+i","1+j")')).toEqual(E('#VALUE!'))
    expect(ev('IMSUB("1+i","1+j")')).toEqual(E('#VALUE!'))
    expect(ev('IMSQRT("-4")')).toBe('2i')
    expect(ev('IMSQRT("-4j")')).toMatch(/j$/)
    expect(ev('IMEXP("0")')).toBe('1')
    expect(ev('IMDIV("1","0")')).toEqual(E('#NUM!'))
    expect(ev('IMLN(0)')).toEqual(E('#NUM!'))
    expect(ev('IMCOT(0)')).toEqual(E('#NUM!'))
    expect(ev('IMEXP(1000)')).toEqual(E('#NUM!'))
  })
  test('IMSUM / IMPRODUCT with ranges and arrays', () => {
    expect(ev('IMSUM(A1:A2)')).toBe('8+i')
    expect(ev('IMSUM(A1:A3,"1")')).toBe('9+i')
    expect(ev('IMPRODUCT(A1:A2)')).toBe('27+11i')
    expect(ev('IMSUM(B2:B3)')).toBe('3')
    expect(ev('IMSUM(B2:B3,2)')).toBe('5')
    expect(ev('IMSUM({"1+i","2+i"})')).toBe('3+2i')
    expect(ev('IMSUM(A1:A2,B1)')).toEqual(E('#VALUE!'))
    expect(ev('IMSUM(C1:C2)')).toEqual(E('#NUM!'))
    expect(ev('IMSUM("1+i",M1)')).toBeInstanceOf(FormulaError)
    expect(ev('IMPRODUCT("2")')).toBe('2')
    expect(ev('IMSUM(IMDIV(A1:A2,"1"))')).toBeDefined()
  })
  test('lifted over arrays', () => {
    expect(ev('SUM(IMABS({"3+4i","5+12i"}))')).toBe(18)
  })
})

describe('CONVERT', () => {
  test('Microsoft examples', () => {
    expectNum('CONVERT(1,"lbm","kg")', 0.45359237, 1e-12)
    expect(ev('CONVERT(68,"F","C")')).toBe(20)
    expect(ev('CONVERT(2.5,"ft","sec")')).toEqual(E('#N/A'))
    expectNum('CONVERT(CONVERT(100,"ft","m"),"ft","m")', 9.290304, 1e-12)
    expectNum('CONVERT(6,"C","F")', 42.8, 1e-12)
  })
  test('categories', () => {
    expectNum('CONVERT(1,"mi","km")', 1.609344, 1e-12)
    expectNum('CONVERT(1,"yr","day")', 365.25, 1e-12)
    expectNum('CONVERT(1,"hr","mn")', 60, 1e-12)
    expectNum('CONVERT(1,"atm","Pa")', 101325, 1e-12)
    expectNum('CONVERT(1,"lbf","N")', 4.4482216152605, 1e-12)
    expectNum('CONVERT(1,"cal","J")', 4.1868, 1e-12)
    expectNum('CONVERT(1,"kWh","J")', 3.6e6, 1e-12)
    expectNum('CONVERT(1,"HP","W")', 745.69987158227, 1e-12)
    expectNum('CONVERT(1,"T","ga")', 10000, 1e-12)
    expectNum('CONVERT(0,"C","K")', 273.15, 1e-12)
    expectNum('CONVERT(32,"F","K")', 273.15, 1e-12)
    expectNum('CONVERT(1,"gal","l")', 3.785411784, 1e-12)
    expectNum('CONVERT(1,"uk_gal","l")', 4.54609, 1e-12)
    expectNum('CONVERT(1,"ha","m2")', 10000, 1e-12)
    expectNum('CONVERT(1,"km2","m^2")', 1e6, 1e-12)
    expectNum('CONVERT(1,"m3","l")', 1000, 1e-12)
    expectNum('CONVERT(1,"byte","bit")', 8, 1e-12)
    expectNum('CONVERT(1,"kbyte","byte")', 1000, 1e-12)
    expectNum('CONVERT(1,"kibyte","byte")', 1024, 1e-12)
    expectNum('CONVERT(1,"Mibit","kibit")', 1024, 1e-12)
    expectNum('CONVERT(1,"m/s","m/h")', 3600, 1e-12)
    expectNum('CONVERT(1,"mph","kn")', 0.868976242, 1e-8)
    expectNum('CONVERT(1,"ton","lbm")', 2000, 1e-12)
    expectNum('CONVERT(1,"stone","lbm")', 14, 1e-12)
  })
  test('prefixes', () => {
    expectNum('CONVERT(1,"km","m")', 1000, 1e-12)
    expectNum('CONVERT(1,"mm","m")', 0.001, 1e-12)
    expectNum('CONVERT(1,"um","m")', 1e-6, 1e-12)
    expectNum('CONVERT(1,"dam","m")', 10, 1e-12)
    expectNum('CONVERT(1,"em","m")', 10, 1e-12)
    expectNum('CONVERT(1,"hPa","Pa")', 100, 1e-12)
    expectNum('CONVERT(1,"Gg","kg")', 1e6, 1e-12)
    expectNum('CONVERT(1,"ml","l")', 0.001, 1e-12)
    expectNum('CONVERT(1,"kN","N")', 1000, 1e-12)
    // m, mi, min, mn, cal, c are units before they are prefixes
    expectNum('CONVERT(60,"min","hr")', 1, 1e-12)
  })
  test('errors', () => {
    expect(ev('CONVERT(1,"xx","m")')).toEqual(E('#N/A'))
    expect(ev('CONVERT(1,"m","xx")')).toEqual(E('#N/A'))
    // case-sensitive
    expect(ev('CONVERT(1,"KM","m")')).toEqual(E('#N/A'))
    expect(ev('CONVERT(1,"M","m")')).toEqual(E('#N/A'))
    // binary prefixes only for bit and byte; metric ones only for metric units
    expect(ev('CONVERT(1,"kim","m")')).toEqual(E('#N/A'))
    expect(ev('CONVERT(1,"kft","m")')).toEqual(E('#N/A'))
    expect(ev('CONVERT(1,"kg","m")')).toEqual(E('#N/A'))
    expect(ev('CONVERT("a","m","ft")')).toEqual(E('#VALUE!'))
  })
})

describe('DELTA / GESTEP / ERF', () => {
  test('DELTA and GESTEP', () => {
    expect(ev('DELTA(5,4)')).toBe(0)
    expect(ev('DELTA(5,5)')).toBe(1)
    expect(ev('DELTA(0.5,0)')).toBe(0)
    expect(ev('DELTA(0)')).toBe(1)
    expect(ev('DELTA("a",1)')).toEqual(E('#VALUE!'))
    expect(ev('GESTEP(5,4)')).toBe(1)
    expect(ev('GESTEP(5,5)')).toBe(1)
    expect(ev('GESTEP(-4,-5)')).toBe(1)
    expect(ev('GESTEP(-1)')).toBe(0)
    expect(ev('GESTEP(4,5)')).toBe(0)
  })
  test('ERF family', () => {
    expectNum('ERF(1)', 0.842700793, 1e-9)
    expectNum('ERF(0,1)', 0.842700793, 1e-9)
    expectNum('ERF.PRECISE(1)', 0.842700793, 1e-9)
    expectNum('ERFC(1)', 0.157299207, 1e-9)
    expectNum('ERFC.PRECISE(1)', 0.157299207, 1e-9)
    expectNum('ERF(-1)', -0.842700793, 1e-9)
    expectNum('ERF(1,2)+ERF(0,1)', 0.995322265018953, 1e-12)
    expectNum('ERF(0.4)+ERFC(0.4)', 1, 1e-15)
    expectNum('ERFC(-1)', 1.842700793, 1e-9)
    expect(ev('ERF("a")')).toEqual(E('#VALUE!'))
  })
})

describe('Bessel functions', () => {
  test('Microsoft examples', () => {
    expectNum('BESSELI(1.5,1)', 0.981666428, 1e-9)
    // the page prints 0.329925829; Excel's own routine is off from the true J2(1.9) = 0.32992572769 by 1e-7
    expectNum('BESSELJ(1.9,2)', 0.329925829, 1e-6)
    expectNum('BESSELJ(1.9,2)', 0.3299257276923868, 1e-13)
    expectNum('BESSELK(1.5,1)', 0.277387804, 1e-8)
    expectNum('BESSELY(2.5,1)', 0.145918138, 1e-8)
  })
  test('reference values', () => {
    expectNum('BESSELJ(1,0)', 0.7651976865579666, 1e-13)
    expectNum('BESSELJ(10,0)', -0.2459357644513483, 1e-12)
    expectNum('BESSELJ(100,0)', 0.019985850304223122, 1e-11)
    expectNum('BESSELJ(-1,1)', -0.4400505857449335, 1e-13)
    expectNum('BESSELY(1,0)', 0.08825696421567697, 1e-12)
    expectNum('BESSELY(10,0)', 0.05567116728359939, 1e-11)
    expectNum('BESSELI(1,0)', 1.2660658777520084, 1e-13)
    expectNum('BESSELI(1,1)', 0.5651591039924851, 1e-13)
    expectNum('BESSELK(1,0)', 0.42102443824070834, 1e-12)
    expectNum('BESSELK(1,1)', 0.6019072301972346, 1e-12)
    expectNum('BESSELJ(0,0)', 1)
    expectNum('BESSELJ(0,3)', 0)
    expectNum('BESSELI(0,0)', 1)
  })
  test('Wronskian identities over many arguments', () => {
    for (const x of [0.1, 0.7, 2.9, 3.1, 8, 25, 60, 300]) {
      for (const n of [0, 1, 4]) {
        // J_{n+1} Y_n - J_n Y_{n+1} = 2/(pi x)
        const w = ev(`BESSELJ(${x},${n + 1})*BESSELY(${x},${n})-BESSELJ(${x},${n})*BESSELY(${x},${n + 1})`) as number
        expect(Math.abs(w - 2 / (Math.PI * x)) < 1e-12 + 1e-9 * (2 / (Math.PI * x))).toBe(true)
      }
    }
    for (const x of [0.1, 0.7, 2.9, 8, 25, 60]) {
      for (const n of [0, 1, 4]) {
        // I_n K_{n+1} + I_{n+1} K_n = 1/x
        const w = ev(`BESSELI(${x},${n})*BESSELK(${x},${n + 1})+BESSELI(${x},${n + 1})*BESSELK(${x},${n})`) as number
        expect(Math.abs(w * x - 1) < 1e-9).toBe(true)
      }
    }
  })
  test('large arguments and orders', () => {
    // J_n^2 sum rule: J0^2 + 2 sum Jk^2 = 1
    let s = Number(ev('BESSELJ(40,0)')) ** 2
    for (let k = 1; k < 90; k++) s += 2 * Number(ev(`BESSELJ(40,${k})`)) ** 2
    expect(close(s, 1, 1e-12)).toBe(true)
    // |J0(x)| <= sqrt(2/(pi x)) in the asymptotic range
    expect(Math.abs(Number(ev('BESSELJ(100000,0)'))) <= Math.sqrt(2 / (Math.PI * 1e5)) * 1.0001).toBe(true)
    // Hankel and Miller agree around the switch
    const a = Number(ev('BESSELJ(50000,2)'))
    const b = Number(ev('BESSELJ(50000.5,2)'))
    expect(Math.abs(a - b) < 0.01).toBe(true)
    expectNum('BESSELJ(1,20)', 3.873503008524658e-25, 1e-9)
    expect(ev('BESSELI(1000,0)')).toEqual(E('#NUM!'))
    expect(typeof ev('BESSELI(500,0)')).toBe('number')
    expect(ev('BESSELK(800,0)')).toBe(0)
  })
  test('argument rules', () => {
    expectNum('BESSELJ(1.9,2.9)', 0.329925829, 1e-6)
    expect(ev('BESSELJ(1,-1)')).toEqual(E('#NUM!'))
    expect(ev('BESSELI(1,-1)')).toEqual(E('#NUM!'))
    expect(ev('BESSELK(1,-1)')).toEqual(E('#NUM!'))
    expect(ev('BESSELY(1,-1)')).toEqual(E('#NUM!'))
    expect(ev('BESSELK(0,1)')).toEqual(E('#NUM!'))
    expect(ev('BESSELK(-1,1)')).toEqual(E('#NUM!'))
    expect(ev('BESSELY(0,1)')).toEqual(E('#NUM!'))
    expect(ev('BESSELY(-1,0)')).toEqual(E('#NUM!'))
    expect(ev('BESSELJ("a",1)')).toEqual(E('#VALUE!'))
    expect(ev('BESSELI(1,"a")')).toEqual(E('#VALUE!'))
    expectNum('BESSELI(-1,1)', -0.5651591039924851, 1e-13)
    expect(ev('BESSELY(1e-300,3)')).toEqual(E('#NUM!'))
  })
})
