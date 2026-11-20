import { describe, expect, test } from 'bun:test'
import { colIndex, evaluateFormula, FormulaError, type CellProvider, type Value } from '../src/core/xlsx/formula'

// ---------------------------------------------------------------- in-memory provider

type Sheets = Record<string, Record<string, Value>>

const addr = (a: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a)!
  return { r: Number(m[2]) - 1, c: colIndex(m[1]) }
}

function makeProvider(sheets: Sheets): CellProvider {
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
  return {
    value: (sheet, r, c) => cells.get(key(sheet ?? 'Sheet1', r, c)) ?? null,
    bounds: (sheet) => ext.get((sheet ?? 'Sheet1').toLowerCase()) ?? { rows: 0, cols: 0 },
    hasSheet: (name) => ext.has(name.toLowerCase()),
  }
}

const ev = (text: string, cells: Record<string, Value> = {}): Value =>
  evaluateFormula(text, { provider: makeProvider({ Sheet1: cells }), sheet: 'Sheet1', r: 49, c: 5 })

const val = (text: string, cells?: Record<string, Value>) => ev(text, cells) as number
const err = (text: string, cells?: Record<string, Value>) => {
  const v = ev(text, cells)
  return v instanceof FormulaError ? v.code : v
}
const D = (y: number, m: number, d: number) => `DATE(${y},${m},${d})`

/** formula -> number with `digits` decimals, the way Microsoft prints them */
function close(text: string, expected: number, digits: number, cells?: Record<string, Value>) {
  expect(val(text, cells)).toBeCloseTo(expected, digits)
}

describe('annuities (Microsoft examples)', () => {
  test('FV', () => {
    close('FV(0.06/12,10,-200,-500,1)', 2581.4, 2)
    close('FV(0.12/12,12,-1000)', 12682.5, 2)
    close('FV(0.11/12,35,-2000,,1)', 82846.25, 2)
    close('FV(0.06/12,12,-100,-1000,1)', 2301.4, 2)
    expect(val('FV(0,10,-5,-3)')).toBe(53)
  })
  test('PV', () => {
    close('PV(0.08/12,12*20,500,,0)', -59777.15, 2)
    expect(val('PV(0,10,-5,3)')).toBe(47)
  })
  test('PMT', () => {
    close('PMT(0.08/12,10,10000)', -1037.03, 2)
    close('PMT(0.08/12,10,10000,,1)', -1030.16, 2)
    close('PMT(0.06/12,18*12,0,50000)', -129.08, 2)
    expect(val('PMT(0,10,1000)')).toBe(-100)
    expect(err('PMT(0.1,0,1000)')).toBe('#NUM!')
  })
  test('NPER', () => {
    close('NPER(0.12/12,-100,-1000,10000,1)', 59.6738657, 6)
    close('NPER(0.12/12,-100,-1000,10000)', 60.0821229, 6)
    close('NPER(0.12/12,-100,-1000)', -9.57859404, 6)
    expect(err('NPER(0.01,0,0)')).toBe('#NUM!')
    expect(err('NPER(0,0,-1000)')).toBe('#NUM!')
  })
  test('RATE', () => {
    close('RATE(4*12,-200,8000)*12', 0.0924, 3)
    close('PMT(RATE(48,-200,8000),48,8000)', -200, 6)
    close('FV(RATE(10,-100,-1000,5000,1),10,-100,-1000,1)', 5000, 5)
    close('RATE(48,-200,8000,0,0,0.5)', 0.0077, 3)
    expect(err('RATE(0,-200,8000)')).toBe('#NUM!')
    expect(err('RATE(10,100,100)')).toBe('#NUM!')
  })
  test('IPMT / PPMT', () => {
    close('IPMT(0.1/12,1,3*12,8000)', -66.67, 2)
    close('IPMT(0.1,3,3,8000)', -292.45, 2)
    close('PPMT(0.1/12,1,2*12,2000)', -75.62, 2)
    close('PPMT(0.08,10,10,200000)', -27598.05, 2)
    expect(val('IPMT(0.1,1,3,8000,0,1)')).toBe(0)
    // interest + principal = payment, for both timings
    for (const t of [0, 1]) {
      close(`IPMT(0.05,4,10,1000,0,${t})+PPMT(0.05,4,10,1000,0,${t})-PMT(0.05,10,1000,0,${t})`, 0, 9)
    }
    expect(err('IPMT(0.1,0,3,8000)')).toBe('#NUM!')
    expect(err('PPMT(0.1,4,3,8000)')).toBe('#NUM!')
  })
  test('ISPMT', () => {
    close('ISPMT(0.1,1,3,8000)', 8000 * 0.1 * (1 / 3 - 1), 9)
    expect(err('ISPMT(0.1,1,0,8000)')).toBe('#DIV/0!')
  })
  test('CUMIPMT / CUMPRINC', () => {
    close('CUMIPMT(0.09/12,30*12,125000,13,24,0)', -11135.23213, 5)
    close('CUMIPMT(0.09/12,30*12,125000,1,1,0)', -937.5, 5)
    close('CUMPRINC(0.09/12,30*12,125000,13,24,0)', -934.1071234, 5)
    close('CUMPRINC(0.09/12,30*12,125000,1,1,0)', -68.27827118, 5)
    // whole life: principal repaid = loan
    close('CUMPRINC(0.01,12,1000,1,12,0)', -1000, 6)
    for (const bad of ['0,12,1000,1,2,0', '0.01,0,1000,1,2,0', '0.01,12,0,1,2,0', '0.01,12,1000,0,2,0', '0.01,12,1000,3,2,0', '0.01,12,1000,1,13,0', '0.01,12,1000,1,2,2']) {
      expect(err(`CUMIPMT(${bad})`)).toBe('#NUM!')
      expect(err(`CUMPRINC(${bad})`)).toBe('#NUM!')
    }
  })
})

describe('cash flows', () => {
  const flows = { A2: -70000, A3: 12000, A4: 15000, A5: 18000, A6: 21000, A7: 26000 }
  test('IRR', () => {
    close('IRR(A2:A6)', -0.021, 3, flows)
    close('IRR(A2:A7)', 0.087, 3, flows)
    close('IRR(A2:A4,-10%)', -0.444, 2, flows)
    close('NPV(IRR(A2:A7),A2:A7)', 0, 5, flows)
    expect(err('IRR(A3:A7)', flows)).toBe('#NUM!')
    expect(err('IRR({-1,-2})')).toBe('#NUM!')
  })
  test('NPV', () => {
    close('NPV(0.1,-10000,3000,4200,6800)', 1188.44, 2)
    const c = { A2: 0.08, A3: -40000, A4: 8000, A5: 9200, A6: 10000, A7: 12000, A8: 14500 }
    close('NPV(A2,A4:A8)+A3', 1922.06, 2, c)
    close('NPV(A2,A4:A8,-9000)+A3', -3749.47, 2, c)
    expect(err('NPV(-1,1,2)')).toBe('#DIV/0!')
  })
  test('MIRR', () => {
    const c = { A2: -120000, A3: 39000, A4: 30000, A5: 21000, A6: 37000, A7: 46000, A8: 0.1, A9: 0.12 }
    close('MIRR(A2:A7,A8,A9)', 0.13, 2, c)
    close('MIRR(A2:A5,A8,A9)', -0.05, 2, c)
    close('MIRR(A2:A7,A8,14%)', 0.13, 2, c)
    // (FV of inflows / PV of outflows)^(1/(n-1)) - 1 by hand
    close('MIRR({-100,50,60},0.1,0.1)', Math.sqrt((50 * 1.1 + 60) / 100) - 1, 9)
    expect(err('MIRR({1,2,3},0.1,0.1)')).toBe('#DIV/0!')
  })
  const xs = {
    A3: -10000, A4: 2750, A5: 4250, A6: 3250, A7: 2750,
    B3: 39448, B4: 39508, B5: 39751, B6: 39859, B7: 39904,
  }
  test('XNPV / XIRR', () => {
    close('XNPV(0.09,A3:A7,B3:B7)', 2086.65, 2, xs)
    close('XIRR(A3:A7,B3:B7,0.1)', 0.373362535, 6, xs)
    close('XIRR(A3:A7,B3:B7)', 0.373362535, 6, xs)
    close('XNPV(XIRR(A3:A7,B3:B7),A3:A7,B3:B7)', 0, 4, xs)
    expect(err('XNPV(0.1,A3:A7,B3:B6)', xs)).toBe('#NUM!')
    expect(err('XNPV(0.1,{1,2},{39500,39400})')).toBe('#NUM!')
    expect(err('XNPV(0.1,{1,2},{"a",39400})')).toBe('#VALUE!')
    expect(err('XIRR({1,2},{39400,39500})')).toBe('#NUM!')
    expect(err('XIRR({-1,2},{39400,39300})')).toBe('#NUM!')
  })
  test('FVSCHEDULE', () => {
    close('FVSCHEDULE(1,{0.09,0.11,0.1})', 1.3309, 4)
    close('FVSCHEDULE(1,{0.09,0.11,0.1})', 1.09 * 1.11 * 1.1, 9)
    expect(val('FVSCHEDULE(100,A1:A2)', { A1: 0.5 })).toBe(150)
    expect(err('FVSCHEDULE(1,{0.1,"x"})')).toBe('#VALUE!')
  })
})

describe('rates', () => {
  test('EFFECT / NOMINAL', () => {
    close('EFFECT(0.0525,4)', 0.0535427, 6)
    close('NOMINAL(EFFECT(0.0525,4),4)', 0.0525, 9)
    close('EFFECT(0.0525,4.9)', 0.0535427, 6)
    expect(err('EFFECT(0,4)')).toBe('#NUM!')
    expect(err('EFFECT(0.05,0)')).toBe('#NUM!')
    expect(err('NOMINAL(0.05,0)')).toBe('#NUM!')
    expect(err('NOMINAL(-0.05,2)')).toBe('#NUM!')
    expect(err('EFFECT("a",2)')).toBe('#VALUE!')
  })
  test('RRI / PDURATION', () => {
    close('RRI(2,100,121)', 0.1, 9)
    close('PDURATION(0.1,100,121)', 2, 9)
    close('FV(RRI(5,1000,1600),5,0,-1000)', 1600, 6)
    expect(err('RRI(0,100,121)')).toBe('#NUM!')
    expect(err('PDURATION(0,100,121)')).toBe('#NUM!')
    expect(err('PDURATION(0.1,-100,121)')).toBe('#NUM!')
  })
  test('DOLLARDE / DOLLARFR', () => {
    expect(val('DOLLARDE(1.02,16)')).toBe(1.125)
    expect(val('DOLLARDE(1.1,32)')).toBe(1.3125)
    expect(val('DOLLARFR(1.125,16)')).toBe(1.02)
    expect(val('DOLLARFR(1.125,32)')).toBe(1.04)
    expect(val('DOLLARDE(1.02,16.9)')).toBe(1.125)
    expect(err('DOLLARDE(1.02,-1)')).toBe('#NUM!')
    expect(err('DOLLARDE(1.02,0.5)')).toBe('#DIV/0!')
    expect(err('DOLLARFR(1.125,-1)')).toBe('#NUM!')
    expect(err('DOLLARFR(1.125,0)')).toBe('#DIV/0!')
  })
})

describe('depreciation', () => {
  test('SLN / SYD', () => {
    expect(val('SLN(30000,7500,10)')).toBe(2250)
    expect(err('SLN(30000,7500,0)')).toBe('#DIV/0!')
    close('SYD(30000,7500,10,1)', 4090.91, 2)
    close('SYD(30000,7500,10,10)', 409.09, 2)
    expect(err('SYD(30000,7500,10,11)')).toBe('#NUM!')
    expect(err('SYD(30000,7500,0,1)')).toBe('#NUM!')
  })
  test('DB', () => {
    const [c, s, l] = ['1000000', '100000', '6']
    const exp = [186083.33, 259639.42, 176814.44, 120410.64, 81999.64, 55841.76, 15845.1]
    exp.forEach((e, i) => close(`DB(${c},${s},${l},${i + 1},7)`, e, 2))
    expect(val('DB(1000,100,5,1)')).toBeGreaterThan(0)
    expect(err('DB(1000,100,5,7)')).toBe('#NUM!')
    expect(err('DB(1000,100,0,1)')).toBe('#NUM!')
    expect(err('DB(1000,100,5,1,13)')).toBe('#NUM!')
    expect(err('DB(-1,100,5,1)')).toBe('#NUM!')
  })
  test('DDB', () => {
    close('DDB(2400,300,10*365,1)', 1.32, 2)
    close('DDB(2400,300,10*12,1,2)', 40, 2)
    close('DDB(2400,300,10,1,2)', 480, 2)
    close('DDB(2400,300,10,2,1.5)', 306, 2)
    close('DDB(2400,300,10,10)', 22.12, 2)
    expect(err('DDB(2400,300,10,11)')).toBe('#NUM!')
    expect(err('DDB(2400,300,10,0)')).toBe('#NUM!')
    expect(err('DDB(2400,300,10,1,0)')).toBe('#NUM!')
    // all periods together leave exactly the salvage value
    close('SUMPRODUCT(DDB(2400,300,10,{1,2,3,4,5,6,7,8,9,10}))', 2100, 6)
  })
  test('VDB', () => {
    close('VDB(2400,300,10*365,0,1)', 1.32, 2)
    close('VDB(2400,300,10*12,0,1)', 40, 2)
    close('VDB(2400,300,10,0,1)', 480, 2)
    close('VDB(2400,300,10*12,6,18)', 396.31, 2)
    close('VDB(2400,300,10*12,6,18,1.5)', 311.81, 2)
    close('VDB(2400,300,10,0,0.875,1.5)', 315, 2)
    close('VDB(2400,300,10,0,10)', 2100, 6)
    // no switch: pure declining balance 2400 * (1 - 0.8^3)
    close('VDB(2400,300,10,0,3,2,TRUE)', 2400 * (1 - 0.8 ** 3), 9)
    expect(val('VDB(2400,300,10,6,9)')).toBeGreaterThanOrEqual(val('VDB(2400,300,10,6,9,2,TRUE)'))
    expect(err('VDB(2400,300,10,5,11)')).toBe('#NUM!')
    expect(err('VDB(2400,300,10,5,4)')).toBe('#NUM!')
    expect(err('VDB(2400,300,10,-1,4)')).toBe('#NUM!')
  })
  test('AMORLINC / AMORDEGRC', () => {
    expect(val('AMORLINC(2400,39679,39813,300,1,0.15,1)')).toBe(360)
    expect(val('AMORDEGRC(2400,39679,39813,300,1,0.15,1)')).toBe(776)
    // the first (part) period, and the straight-line periods all together leave the salvage
    close('SUMPRODUCT(AMORLINC(2400,39679,39813,300,{0,1,2,3,4,5,6,7},0.15,1))', 2100, 6)
    expect(val('AMORLINC(2400,39679,39813,300,50,0.15,1)')).toBe(0)
    expect(err('AMORLINC(2400,39679,39813,300,1,0.15,2)')).toBe('#NUM!')
    expect(err('AMORLINC(2400,39679,39813,3000,1,0.15,1)')).toBe('#NUM!')
    expect(err('AMORDEGRC(2400,39679,39813,300,1,0.4,1)')).toBe('#NUM!')
  })
})

describe('coupons', () => {
  const a = `${D(2011, 1, 25)},${D(2011, 11, 15)},2,1`
  test('Microsoft examples', () => {
    expect(val(`COUPDAYBS(${a})`)).toBe(71)
    expect(val(`COUPDAYS(${a})`)).toBe(181)
    expect(val(`COUPDAYSNC(${a})`)).toBe(110)
    expect(val(`COUPNCD(${a})`)).toBe(val(D(2011, 5, 15)))
    expect(val(`COUPPCD(${a})`)).toBe(val(D(2010, 11, 15)))
    expect(val(`COUPNUM(${D(2007, 1, 25)},${D(2008, 11, 15)},2,1)`)).toBe(4)
  })
  test('other bases and edges', () => {
    expect(val(`COUPDAYS(${D(2011, 1, 25)},${D(2011, 11, 15)},2,0)`)).toBe(180)
    expect(val(`COUPDAYS(${D(2011, 1, 25)},${D(2011, 11, 15)},2,3)`)).toBe(182.5)
    // 30/360: 15 Nov -> 25 Jan = 70 days, 110 to go
    expect(val(`COUPDAYBS(${D(2011, 1, 25)},${D(2011, 11, 15)},2,0)`)).toBe(70)
    expect(val(`COUPDAYSNC(${D(2011, 1, 25)},${D(2011, 11, 15)},2,0)`)).toBe(110)
    // settlement on a coupon date: it is the previous one
    expect(val(`COUPPCD(${D(2011, 5, 15)},${D(2011, 11, 15)},2)`)).toBe(val(D(2011, 5, 15)))
    expect(val(`COUPNUM(${D(2011, 5, 15)},${D(2011, 11, 15)},2)`)).toBe(1)
    // month-end maturities keep to month ends
    expect(val(`COUPPCD(${D(2011, 1, 10)},${D(2011, 2, 28)},2)`)).toBe(val(D(2010, 8, 31)))
    expect(val(`COUPNUM(${D(2011, 1, 10)},${D(2013, 2, 28)},4)`)).toBe(9)
  })
  test('errors', () => {
    expect(err(`COUPNUM(${D(2011, 1, 25)},${D(2011, 11, 15)},3)`)).toBe('#NUM!')
    expect(err(`COUPNUM(${D(2011, 1, 25)},${D(2011, 11, 15)},2,5)`)).toBe('#NUM!')
    expect(err(`COUPNUM(${D(2011, 11, 15)},${D(2011, 11, 15)},2)`)).toBe('#NUM!')
    expect(err(`COUPNUM(${D(2012, 1, 1)},${D(2011, 11, 15)},2)`)).toBe('#NUM!')
    expect(err(`COUPNUM("x",${D(2011, 11, 15)},2)`)).toBe('#VALUE!')
    expect(err(`COUPNUM(-5,${D(2011, 11, 15)},2)`)).toBe('#NUM!')
  })
})

describe('bonds with periodic interest', () => {
  const price = `${D(2008, 2, 15)},${D(2017, 11, 15)},0.0575`
  test('PRICE / YIELD', () => {
    close(`PRICE(${price},0.065,100,2,0)`, 94.63, 2)
    close(`YIELD(${D(2008, 2, 15)},${D(2016, 11, 15)},0.0575,95.04287,100,2,0)`, 0.065, 4)
    // YIELD undoes PRICE, on every basis and frequency, long and single-period
    for (const [f, b] of [[1, 0], [2, 1], [4, 2], [2, 3], [1, 4]]) {
      const p = val(`PRICE(${price},0.071,100,${f},${b})`)
      close(`YIELD(${price},${p},100,${f},${b})`, 0.071, 7)
    }
    const p1 = val(`PRICE(${D(2008, 2, 15)},${D(2008, 11, 15)},0.06,0.05,100,2,1)`)
    close(`YIELD(${D(2008, 2, 15)},${D(2008, 11, 15)},0.06,${p1},100,2,1)`, 0.05, 9)
    // zero yield: par + all coupons less the accrued part
    close(`PRICE(${D(2008, 1, 1)},${D(2010, 1, 1)},0.05,0,100,1,0)`, 110, 9)
  })
  test('errors', () => {
    expect(err(`PRICE(${price},-0.01,100,2)`)).toBe('#NUM!')
    expect(err(`PRICE(${price},0.05,0,2)`)).toBe('#NUM!')
    expect(err(`PRICE(${price},0.05,100,3)`)).toBe('#NUM!')
    expect(err(`PRICE(${D(2018, 1, 1)},${D(2017, 11, 15)},0.05,0.05,100,2)`)).toBe('#NUM!')
    expect(err(`PRICE(${price},0.05,100,2,5)`)).toBe('#NUM!')
    expect(err(`YIELD(${price},0,100,2)`)).toBe('#NUM!')
    expect(err(`YIELD(${price},90,0,2)`)).toBe('#NUM!')
  })
  test('DURATION / MDURATION', () => {
    close(`DURATION(${D(2018, 7, 1)},${D(2048, 1, 1)},0.08,0.09,2,1)`, 10.9191453, 5)
    close(`MDURATION(${D(2008, 1, 1)},${D(2016, 1, 1)},0.08,0.09,2,1)`, 5.736, 3)
    // a zero-coupon bond lasts until maturity
    close(`DURATION(${D(2008, 1, 1)},${D(2013, 1, 1)},0,0.05,1,0)`, 5, 9)
    expect(err(`DURATION(${D(2008, 1, 1)},${D(2013, 1, 1)},-0.01,0.05,1)`)).toBe('#NUM!')
    expect(err(`DURATION(${D(2008, 1, 1)},${D(2013, 1, 1)},0.01,-0.05,1)`)).toBe('#NUM!')
    expect(err(`MDURATION(${D(2008, 1, 1)},${D(2013, 1, 1)},0.01,0.05,3)`)).toBe('#NUM!')
  })
})

describe('odd first / last period', () => {
  const f = [D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1)]
  test('ODDFPRICE / ODDFYIELD', () => {
    close(`ODDFPRICE(${f},0.0785,0.0625,100,2,1)`, 113.6, 2)
    close(`ODDFYIELD(${f},0.0575,84.5,100,2,0)`, 0.0772, 4)
    for (const b of [0, 1, 2, 3, 4]) {
      const p = val(`ODDFPRICE(${f},0.0785,0.0625,100,2,${b})`)
      close(`ODDFYIELD(${f},0.0785,${p},100,2,${b})`, 0.0625, 7)
    }
    // a long first period (issue a year before) round-trips too
    const g = [D(2008, 11, 11), D(2021, 3, 1), D(2008, 1, 20), D(2009, 3, 1)]
    const p = val(`ODDFPRICE(${g},0.06,0.05,100,2,1)`)
    close(`ODDFYIELD(${g},0.06,${p},100,2,1)`, 0.05, 7)
  })
  test('ODDLPRICE / ODDLYIELD', () => {
    close(`ODDLPRICE(${D(2008, 2, 7)},${D(2008, 6, 15)},${D(2007, 10, 15)},0.0375,0.0405,100,2,0)`, 99.88, 2)
    close(`ODDLYIELD(${D(2008, 4, 20)},${D(2008, 6, 15)},${D(2007, 12, 24)},0.0375,99.875,100,2,0)`, 0.04519, 5)
    for (const b of [0, 1, 2, 3, 4]) {
      const a = `${D(2008, 2, 7)},${D(2008, 6, 15)},${D(2007, 10, 15)},0.0375`
      const p = val(`ODDLPRICE(${a},0.0405,100,2,${b})`)
      close(`ODDLYIELD(${a},${p},100,2,${b})`, 0.0405, 9)
    }
  })
  test('errors', () => {
    // maturity > first_coupon > settlement > issue
    expect(err(`ODDFPRICE(${D(2008, 11, 11)},${D(2021, 3, 1)},${D(2008, 12, 15)},${D(2009, 3, 1)},0.05,0.05,100,2)`)).toBe('#NUM!')
    expect(err(`ODDFPRICE(${D(2008, 11, 11)},${D(2021, 3, 1)},${D(2008, 10, 15)},${D(2008, 10, 20)},0.05,0.05,100,2)`)).toBe('#NUM!')
    expect(err(`ODDFPRICE(${f},-0.05,0.05,100,2)`)).toBe('#NUM!')
    expect(err(`ODDFPRICE(${f},0.05,-0.05,100,2)`)).toBe('#NUM!')
    expect(err(`ODDFYIELD(${f},0.05,0,100,2)`)).toBe('#NUM!')
    expect(err(`ODDFPRICE(${f},0.05,0.05,100,2,7)`)).toBe('#NUM!')
    expect(err(`ODDFPRICE("x",${D(2021, 3, 1)},${D(2008, 10, 15)},${D(2009, 3, 1)},0.05,0.05,100,2)`)).toBe('#VALUE!')
    // maturity > settlement > last_interest
    expect(err(`ODDLPRICE(${D(2007, 9, 1)},${D(2008, 6, 15)},${D(2007, 10, 15)},0.05,0.05,100,2)`)).toBe('#NUM!')
    expect(err(`ODDLPRICE(${D(2008, 7, 1)},${D(2008, 6, 15)},${D(2007, 10, 15)},0.05,0.05,100,2)`)).toBe('#NUM!')
    expect(err(`ODDLYIELD(${D(2008, 2, 7)},${D(2008, 6, 15)},${D(2007, 10, 15)},0.05,0,100,2)`)).toBe('#NUM!')
    expect(err(`ODDLPRICE(${D(2008, 2, 7)},${D(2008, 6, 15)},${D(2007, 10, 15)},0.05,0.05,100,3)`)).toBe('#NUM!')
  })
})

describe('accrued interest', () => {
  const c = { A2: 39508, A3: 39691, A4: 39569, A5: 0.1, A6: 1000, A7: 2, A8: 0 }
  test('ACCRINT', () => {
    close('ACCRINT(A2,A3,A4,A5,A6,A7,A8)', 16.666667, 6, c)
    close(`ACCRINT(${D(2008, 3, 5)},A3,A4,A5,A6,A7,A8,FALSE)`, 15.555556, 6, c)
    close(`ACCRINT(${D(2008, 4, 5)},A3,A4,A5,A6,A7,A8,TRUE)`, 7.2222222, 6, c)
    // from first interest date: 30/360 from 31 Aug to 31 Oct = 60 days
    close(`ACCRINT(A2,A3,${D(2008, 10, 31)},A5,A6,A7,0,FALSE)`, 1000 * 0.1 * (60 / 360), 9, c)
    close(`ACCRINT(A2,A3,${D(2008, 10, 31)},A5,A6,A7,0,TRUE)`, 1000 * 0.1 * (240 / 360), 9, c)
    expect(err('ACCRINT(A2,A3,A4,0,A6,A7,A8)', c)).toBe('#NUM!')
    expect(err('ACCRINT(A2,A3,A4,A5,0,A7,A8)', c)).toBe('#NUM!')
    expect(err('ACCRINT(A2,A3,A4,A5,A6,3,A8)', c)).toBe('#NUM!')
    expect(err('ACCRINT(A2,A3,A4,A5,A6,A7,5)', c)).toBe('#NUM!')
    expect(err('ACCRINT(A4,A3,A2,A5,A6,A7,A8)', c)).toBe('#NUM!')
    expect(err('ACCRINT("x",A3,A4,A5,A6,A7,A8)', c)).toBe('#VALUE!')
  })
  test('ACCRINTM', () => {
    close('ACCRINTM(39539,39614,0.1,1000,3)', 20.54794521, 8)
    close('ACCRINTM(39539,39614,0.1,,3)', 20.54794521, 8)
    expect(err('ACCRINTM(39539,39614,0,1000,3)')).toBe('#NUM!')
    expect(err('ACCRINTM(39539,39614,0.1,-1,3)')).toBe('#NUM!')
    expect(err('ACCRINTM(39614,39539,0.1,1000,3)')).toBe('#NUM!')
    expect(err('ACCRINTM(39539,39614,0.1,1000,9)')).toBe('#NUM!')
  })
})

describe('discount and maturity securities', () => {
  test('DISC / PRICEDISC / YIELDDISC', () => {
    // the Microsoft page prints 0.001038 for 1 Jul 2018 -> 1 Jan 2048 (a 19.5 year span), which its own YEARFRAC does not give
    close(`DISC(${D(2018, 7, 1)},${D(2038, 1, 1)},97.975,100,1)`, 0.001038, 5)
    close(`DISC(${D(2018, 7, 1)},${D(2048, 1, 1)},97.975,100,1)`, 0.02025 / (29.5 + 1 / 365), 4)
    close(`PRICEDISC(${D(2008, 2, 16)},${D(2008, 3, 1)},0.0525,100,2)`, 99.8, 2)
    close(`YIELDDISC(${D(2008, 2, 16)},${D(2008, 3, 1)},99.795,100,2)`, 0.052823, 5)
    // PRICEDISC and YIELDDISC are inverse
    const p = val(`PRICEDISC(${D(2008, 2, 16)},${D(2008, 3, 1)},0.0525,100,2)`)
    close(`DISC(${D(2008, 2, 16)},${D(2008, 3, 1)},${p},100,2)`, 0.0525, 9)
    expect(err(`DISC(${D(2008, 2, 16)},${D(2008, 3, 1)},0,100)`)).toBe('#NUM!')
    expect(err(`PRICEDISC(${D(2008, 2, 16)},${D(2008, 3, 1)},0.05,0)`)).toBe('#NUM!')
    expect(err(`YIELDDISC(${D(2008, 3, 1)},${D(2008, 2, 16)},99,100)`)).toBe('#NUM!')
    expect(err(`DISC(${D(2008, 2, 16)},${D(2008, 3, 1)},99,100,5)`)).toBe('#NUM!')
  })
  test('INTRATE / RECEIVED', () => {
    close(`INTRATE(${D(2008, 2, 15)},${D(2008, 5, 15)},1000000,1014420,2)`, 0.05768, 5)
    close(`RECEIVED(${D(2008, 2, 15)},${D(2008, 5, 15)},1000000,0.0575,2)`, 1014584.65, 2)
    expect(err(`INTRATE(${D(2008, 2, 15)},${D(2008, 5, 15)},0,1014420,2)`)).toBe('#NUM!')
    expect(err(`RECEIVED(${D(2008, 2, 15)},${D(2008, 5, 15)},1000000,0,2)`)).toBe('#NUM!')
    expect(err(`RECEIVED(${D(2008, 5, 15)},${D(2008, 5, 15)},1000000,0.05,2)`)).toBe('#NUM!')
  })
  test('PRICEMAT / YIELDMAT', () => {
    close(`PRICEMAT(${D(2008, 2, 15)},${D(2008, 4, 13)},${D(2007, 11, 11)},0.061,0.061,0)`, 99.98, 2)
    close(`YIELDMAT(${D(2008, 3, 15)},${D(2008, 11, 3)},${D(2007, 11, 8)},0.0625,100.0123,0)`, 0.060954, 5)
    const a = `${D(2008, 2, 15)},${D(2008, 4, 13)},${D(2007, 11, 11)},0.061`
    const p = val(`PRICEMAT(${a},0.07,1)`)
    close(`YIELDMAT(${a},${p},1)`, 0.07, 9)
    expect(err(`PRICEMAT(${a},-0.01)`)).toBe('#NUM!')
    expect(err(`YIELDMAT(${a},0)`)).toBe('#NUM!')
    expect(err(`PRICEMAT(${D(2008, 4, 13)},${D(2008, 2, 15)},${D(2007, 11, 11)},0.06,0.06)`)).toBe('#NUM!')
  })
  test('TBILL*', () => {
    const d = `${D(2008, 3, 31)},${D(2008, 6, 1)}`
    close(`TBILLEQ(${d},0.0914)`, 0.09415, 4)
    close(`TBILLPRICE(${d},0.09)`, 98.45, 2)
    close(`TBILLYIELD(${d},98.45)`, 0.0914, 4)
    // more than half a year: bond-equivalent yield, bill priced back from it
    const l = `${D(2008, 3, 31)},${D(2008, 12, 31)}`
    const y = val(`TBILLEQ(${l},0.05)`)
    const price = val(`TBILLPRICE(${l},0.05)`)
    expect(price * (1 + y / 2) * (1 + (275 / 365 - 0.5) * y)).toBeCloseTo(100, 6)
    expect(err(`TBILLPRICE(${D(2008, 3, 31)},${D(2009, 4, 2)},0.05)`)).toBe('#NUM!')
    expect(val(`TBILLPRICE(${D(2008, 3, 31)},${D(2009, 3, 31)},0.05)`)).toBeCloseTo(95, 0)
    expect(err(`TBILLEQ(${D(2008, 6, 1)},${D(2008, 3, 31)},0.05)`)).toBe('#NUM!')
    expect(err(`TBILLEQ(${d},0)`)).toBe('#NUM!')
    expect(err(`TBILLYIELD(${d},0)`)).toBe('#NUM!')
    expect(err(`TBILLPRICE(${d},-0.1)`)).toBe('#NUM!')
  })
})

describe('plumbing', () => {
  test('array arguments lift', () => {
    const v = ev('PMT(0.05,{10,20},1000)') as unknown
    expect(v).toBeDefined()
    expect(val('SUM(PMT(0.05,{10,20},1000))')).toBeCloseTo(-129.5046 - 80.2426, 3)
  })
  test('errors in arguments pass through', () => {
    expect(err('FV(#DIV/0!,1,1)')).toBe('#DIV/0!')
    expect(err('NPV(0.1,#N/A)')).toBe('#N/A')
    expect(err('IFERROR(IRR({1,2}),-1)')).toBe(-1)
  })
  test('too few arguments', () => {
    expect(err('FV(0.1,1)')).toBe('#VALUE!')
  })
})
