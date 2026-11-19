import { describe, expect, test } from 'bun:test'
import {
  colIndex,
  evaluateArray,
  evaluateFormula,
  FormulaError,
  parseFormula,
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

/** Raw values only (a text starting with `=` stays text: the database criteria use it). */
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

const E = (code: ErrorCode) => new FormulaError(code)
const LETTERS = 'ABCDEFGHIJ'

/** Columns of values, A, B, C … from row 1 */
function cols(...data: Value[][]): CellProvider {
  const s: Record<string, Value> = {}
  data.forEach((d, j) => d.forEach((v, i) => (s[`${LETTERS[j]}${i + 1}`] = v)))
  return makeProvider({ Sheet1: s })
}

const ctx = (provider: CellProvider) => ({ provider, sheet: 'Sheet1', r: 59, c: 25 })
const ev = (f: string, ...data: Value[][]) => evaluateFormula(f, ctx(cols(...data)))
const evOn = (f: string, p: CellProvider) => evaluateFormula(f, ctx(p))
const arr = (f: string, h: number, w: number, ...data: Value[][]) => evaluateArray(parseFormula(f), ctx(cols(...data)), h, w)

function near(got: Value, want: number, digits = 6) {
  expect(typeof got).toBe('number')
  expect(got as number).toBeCloseTo(want, digits)
}

function nearAll(got: Value[][], want: number[][], digits = 6) {
  expect(got.length).toBe(want.length)
  want.forEach((row, i) => row.forEach((w, j) => near(got[i][j], w, digits)))
}

// ---------------------------------------------------------------- data of the Microsoft pages

const D7 = [4, 5, 8, 7, 11, 4, 3] // DEVSQ, GEOMEAN, HARMEAN
const YX_Y = [2, 3, 9, 1, 8, 7, 5] // RSQ, SLOPE, STEYX
const YX_X = [6, 5, 11, 7, 5, 4, 4]
const STRENGTH = [1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299]
const SK = [3, 4, 5, 2, 3, 4, 5, 6, 4, 7] // KURT, SKEW

describe('single-column statistics', () => {
  test('AVEDEV', () => {
    near(ev('AVEDEV(A1:A7)', [4, 5, 6, 7, 5, 4, 3]), 1.020408)
    near(ev('AVEDEV(1,2,"3")'), 2 / 3)
    expect(ev('AVEDEV(A1:A2)', ['a', 'b'])).toEqual(E('#NUM!'))
    expect(ev('AVEDEV(A1:A2)', [1, E('#N/A')])).toEqual(E('#N/A'))
  })

  test('AVERAGEA counts text as 0 and logicals in ranges', () => {
    near(ev('AVERAGEA(A1:A5)', [10, 7, 9, 2, 'Not available']), 5.6)
    // an empty cell is skipped (the page prints 5.6 here, which would count it)
    near(ev('AVERAGEA(A1:A4,A6)', [10, 7, 9, 2, 'x']), 7)
    near(ev('AVERAGEA(A1:A4)', [1, true, false, 'x']), 0.5)
    expect(ev('AVERAGEA(A1:A2)', [null, null])).toEqual(E('#DIV/0!'))
    expect(ev('AVERAGEA("abc")')).toEqual(E('#VALUE!'))
    near(ev('AVERAGEA(TRUE,3)'), 2)
  })

  test('DEVSQ', () => {
    near(ev('DEVSQ(A1:A7)', D7), 48)
    expect(ev('DEVSQ(A1:A2)', ['a', 'b'])).toEqual(E('#NUM!'))
  })

  test('GEOMEAN / HARMEAN', () => {
    near(ev('GEOMEAN(A1:A7)', D7), 5.476987)
    near(ev('HARMEAN(A1:A7)', D7), 5.028376)
    expect(ev('GEOMEAN(A1:A2)', [4, 0])).toEqual(E('#NUM!'))
    expect(ev('HARMEAN(A1:A2)', [4, -1])).toEqual(E('#NUM!'))
    expect(ev('GEOMEAN(A1:A2)', ['a', 'b'])).toEqual(E('#NUM!'))
  })

  test('KURT / SKEW / SKEW.P', () => {
    near(ev('KURT(A1:A10)', SK), -0.151799637, 8)
    near(ev('SKEW(A1:A10)', SK), 0.359543)
    near(ev('SKEW.P(A1:A10)', SK), 0.303193)
    expect(ev('KURT(A1:A3)', [1, 2, 3])).toEqual(E('#DIV/0!'))
    expect(ev('KURT(A1:A4)', [2, 2, 2, 2])).toEqual(E('#DIV/0!'))
    expect(ev('SKEW(A1:A2)', [1, 2])).toEqual(E('#DIV/0!'))
    expect(ev('SKEW.P(A1:A3)', [5, 5, 5])).toEqual(E('#DIV/0!'))
  })

  test('MAXA / MINA', () => {
    near(ev('MAXA(A1:A5)', [0, 0.2, 0.5, 0.4, true]), 1)
    near(ev('MINA(A1:A5)', [false, 0.2, 0.5, 0.4, 0.8]), 0)
    near(ev('MAXA(A1:A2)', [-5, 'text']), 0)
    expect(ev('MAXA(A1:A2)', [null, null])).toEqual(0)
    expect(ev('MINA(A1:A2)', [E('#REF!'), 1])).toEqual(E('#REF!'))
  })

  test('STDEVA / STDEVPA / VARA / VARPA', () => {
    near(ev('STDEVA(A1:A10)', STRENGTH), 27.46391572, 8)
    near(ev('STDEVPA(A1:A10)', STRENGTH), 26.05455814, 8)
    near(ev('VARA(A1:A10)', STRENGTH), 754.26667, 5)
    near(ev('VARPA(A1:A10)', STRENGTH), 678.84, 5)
    // text counts 0 and TRUE counts 1 in a range
    near(ev('VARA(A1:A3)', [1, 'x', true]), 1 / 3)
    expect(ev('VARA(A1:A1)', [1])).toEqual(E('#DIV/0!'))
    expect(ev('VARPA(A1:A1)', [null])).toEqual(E('#DIV/0!'))
    near(ev('VARPA(A1:A1)', [3]), 0)
  })

  test('TRIMMEAN', () => {
    near(ev('TRIMMEAN(A1:A11,0.2)', [4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3]), 3.778, 3)
    near(ev('TRIMMEAN(A1:A4,0)', [1, 2, 3, 10]), 4)
    expect(ev('TRIMMEAN(A1:A4,1)', [1, 2, 3, 10])).toEqual(E('#NUM!'))
    near(ev('TRIMMEAN(A1:A5,0.8)', [1, 2, 3, 10, 50]), 3)
    expect(ev('TRIMMEAN(A1:A4,1.5)', [1, 2, 3, 10])).toEqual(E('#NUM!'))
    expect(ev('TRIMMEAN(A1:A4,-0.1)', [1, 2, 3, 10])).toEqual(E('#NUM!'))
    expect(ev('TRIMMEAN(A1:A2,0.2)', ['a', 'b'])).toEqual(E('#NUM!'))
    expect(ev('TRIMMEAN(A1:A4,"x")', [1, 2, 3, 10])).toEqual(E('#VALUE!'))
  })

  test('AVERAGE.WEIGHTED', () => {
    near(ev('AVERAGE.WEIGHTED(A1:A3,B1:B3)', [1, 2, 3], [3, 2, 1]), 10 / 6)
    near(ev('AVERAGE.WEIGHTED(A1:A2,B1:B2,C1:C2,D1:D2)', [1, 2], [1, 1], [10, 20], [2, 0]), (1 + 2 + 20) / 4)
    expect(ev('AVERAGE.WEIGHTED(A1:A2,B1:B3)', [1, 2], [1, 1, 1])).toEqual(E('#N/A'))
    expect(ev('AVERAGE.WEIGHTED(A1:A2,B1:B2)', [1, 2], [0, 0])).toEqual(E('#DIV/0!'))
    expect(ev('AVERAGE.WEIGHTED(A1:A2,B1:B2)', [1, 2], [1, -1])).toEqual(E('#VALUE!'))
    expect(ev('AVERAGE.WEIGHTED(A1:A2)', [1, 2])).toEqual(E('#VALUE!'))
  })
})

describe('paired data', () => {
  test('CORREL / PEARSON', () => {
    near(ev('PEARSON(A1:A5,B1:B5)', [9, 7, 5, 3, 1], [10, 6, 1, 5, 3]), 0.699379)
    near(ev('CORREL(A1:A5,B1:B5)', [9, 7, 5, 3, 1], [10, 6, 1, 5, 3]), 0.699379)
    near(ev('CORREL(A1:A3,B1:B3)', [1, 2, 3], [2, 4, 6]), 1)
    near(ev('CORREL(A1:A3,B1:B3)', [1, 2, 3], [6, 4, 2]), -1)
    // a pair with a non-number is dropped
    near(ev('CORREL(A1:A4,B1:B4)', [1, 2, 'x', 3], [2, 4, 100, 6]), 1)
    expect(ev('CORREL(A1:A3,B1:B2)', [1, 2, 3], [1, 2])).toEqual(E('#N/A'))
    expect(ev('CORREL(A1:A3,B1:B3)', [1, 1, 1], [1, 2, 3])).toEqual(E('#DIV/0!'))
    expect(ev('PEARSON(A1:A2,B1:B2)', ['a', 'b'], [1, 2])).toEqual(E('#N/A'))
    expect(ev('CORREL(A1:A2,B1:B2)', [1, E('#NUM!')], [1, 2])).toEqual(E('#NUM!'))
  })

  test('COVAR / COVARIANCE.P / COVARIANCE.S', () => {
    const x = [3, 2, 4, 5, 6]
    const y = [9, 7, 12, 15, 17]
    near(ev('COVAR(A1:A5,B1:B5)', x, y), 5.2)
    near(ev('COVARIANCE.P(A1:A5,B1:B5)', x, y), 5.2)
    near(ev('COVARIANCE.S({2,4,8},{5,11,12})'), 9.666666667, 8)
    near(ev('COVARIANCE.S(A1:A3,B1:B3)', [2, 4, 8], [5, 11, 12]), 9.666666667, 8)
    expect(ev('COVAR(A1:A3,B1:B2)', [1, 2, 3], [1, 2])).toEqual(E('#N/A'))
    expect(ev('COVAR(A1:A2,B1:B2)', ['a', 'b'], ['c', 'd'])).toEqual(E('#DIV/0!'))
    expect(ev('COVARIANCE.S(A1:A1,B1:B1)', [1], [2])).toEqual(E('#DIV/0!'))
    near(ev('COVARIANCE.P(A1:A1,B1:B1)', [1], [2]), 0)
  })

  test('SLOPE / INTERCEPT / RSQ / STEYX', () => {
    near(ev('SLOPE(A1:A7,B1:B7)', YX_Y, YX_X), 0.305556)
    near(ev('RSQ(A1:A7,B1:B7)', YX_Y, YX_X), 0.05795, 4)
    near(ev('STEYX(A1:A7,B1:B7)', YX_Y, YX_X), 3.305719)
    near(ev('INTERCEPT(A1:A5,B1:B5)', [2, 3, 9, 1, 8], [6, 5, 11, 7, 5]), 0.0483871, 7)
    expect(ev('SLOPE(A1:A3,B1:B2)', [1, 2, 3], [1, 2])).toEqual(E('#N/A'))
    expect(ev('SLOPE(A1:A2,B1:B2)', ['a', 'b'], ['c', 'd'])).toEqual(E('#N/A'))
    expect(ev('SLOPE(A1:A3,B1:B3)', [1, 2, 3], [4, 4, 4])).toEqual(E('#DIV/0!'))
    expect(ev('INTERCEPT(A1:A3,B1:B3)', [1, 2, 3], [4, 4, 4])).toEqual(E('#DIV/0!'))
    expect(ev('RSQ(A1:A1,B1:B1)', [1], [2])).toEqual(E('#DIV/0!'))
    expect(ev('STEYX(A1:A2,B1:B2)', [1, 2], [1, 2])).toEqual(E('#DIV/0!'))
    near(ev('RSQ(A1:A3,B1:B3)', [1, 2, 3], [2, 4, 6]), 1)
  })

  test('FORECAST / FORECAST.LINEAR', () => {
    const y = [6, 7, 9, 15, 21]
    const x = [20, 28, 31, 38, 40]
    near(ev('FORECAST.LINEAR(30,A1:A5,B1:B5)', y, x), 10.607253)
    near(ev('FORECAST(30,A1:A5,B1:B5)', y, x), 10.607253)
    near(ev('FORECAST(4,A1:A3,B1:B3)', [2, 4, 6], [1, 2, 3]), 8)
    expect(ev('FORECAST("a",A1:A3,B1:B3)', [2, 4, 6], [1, 2, 3])).toEqual(E('#VALUE!'))
    expect(ev('FORECAST(1,A1:A3,B1:B2)', [2, 4, 6], [1, 2])).toEqual(E('#N/A'))
    expect(ev('FORECAST(1,A1:A3,B1:B3)', [2, 4, 6], [5, 5, 5])).toEqual(E('#DIV/0!'))
  })
})

describe('ranking and percentiles', () => {
  const A9 = [1, 2, 3, 6, 6, 6, 7, 8, 9]
  const A10 = [13, 12, 11, 8, 4, 3, 2, 1, 1, 1]

  test('PERCENTILE family', () => {
    near(ev('PERCENTILE.INC(A1:A4,0.3)', [1, 3, 2, 4]), 1.9)
    near(ev('PERCENTILE(A1:A4,0.3)', [1, 3, 2, 4]), 1.9)
    near(ev('PERCENTILE.INC(A1:A4,0)', [1, 3, 2, 4]), 1)
    near(ev('PERCENTILE.INC(A1:A4,1)', [1, 3, 2, 4]), 4)
    near(ev('PERCENTILE.EXC(A1:A4,0.5)', [1, 3, 2, 4]), 2.5)
    near(ev('PERCENTILE.EXC(A1:A4,0.2)', [1, 3, 2, 4]), 1)
    expect(ev('PERCENTILE.EXC(A1:A4,0.1)', [1, 3, 2, 4])).toEqual(E('#NUM!'))
    expect(ev('PERCENTILE.EXC(A1:A4,0.9)', [1, 3, 2, 4])).toEqual(E('#NUM!'))
    expect(ev('PERCENTILE.EXC(A1:A4,0)', [1, 3, 2, 4])).toEqual(E('#NUM!'))
    expect(ev('PERCENTILE.EXC(A1:A4,1)', [1, 3, 2, 4])).toEqual(E('#NUM!'))
    expect(ev('PERCENTILE.INC(A1:A4,1.1)', [1, 3, 2, 4])).toEqual(E('#NUM!'))
    expect(ev('PERCENTILE.INC(A1:A4,-0.1)', [1, 3, 2, 4])).toEqual(E('#NUM!'))
    expect(ev('PERCENTILE(A1:A4,"x")', [1, 3, 2, 4])).toEqual(E('#VALUE!'))
    expect(ev('PERCENTILE(A1:A2,0.5)', ['a', 'b'])).toEqual(E('#NUM!'))
    expect(ev('PERCENTILE.EXC(A1:A2,0.5)', [null, null])).toEqual(E('#NUM!'))
  })

  test('PERCENTILE with an array of k', () => {
    nearAll(arr('PERCENTILE.INC(A1:A4,{0;1})', 2, 1, [1, 3, 2, 4]), [[1], [4]])
  })

  test('QUARTILE family', () => {
    near(ev('QUARTILE(A1:A8,1)', [1, 2, 4, 7, 8, 9, 10, 12]), 3.5)
    near(ev('QUARTILE.INC(A1:A8,1)', [1, 2, 4, 7, 8, 9, 10, 12]), 3.5)
    const d = [6, 7, 15, 36, 39, 40, 41, 42, 43, 47, 49]
    near(ev('QUARTILE.EXC(A1:A11,1)', d), 15)
    near(ev('QUARTILE.EXC(A1:A11,3)', d), 43)
    near(ev('QUARTILE.EXC(A1:A11,2)', d), 40)
    near(ev('QUARTILE.INC(A1:A11,0)', d), 6)
    near(ev('QUARTILE.INC(A1:A11,4)', d), 49)
    near(ev('QUARTILE.INC(A1:A11,2.9)', d), 40)
    expect(ev('QUARTILE.INC(A1:A11,5)', d)).toEqual(E('#NUM!'))
    expect(ev('QUARTILE.INC(A1:A11,-1)', d)).toEqual(E('#NUM!'))
    expect(ev('QUARTILE.EXC(A1:A11,0)', d)).toEqual(E('#NUM!'))
    expect(ev('QUARTILE.EXC(A1:A11,4)', d)).toEqual(E('#NUM!'))
    expect(ev('QUARTILE(A1:A2,1)', ['a', 'b'])).toEqual(E('#NUM!'))
    expect(ev('QUARTILE(A1:A3,"q")', [1, 2, 3])).toEqual(E('#VALUE!'))
  })

  test('PERCENTRANK.EXC', () => {
    near(ev('PERCENTRANK.EXC(A1:A9,7)', A9), 0.7)
    near(ev('PERCENTRANK.EXC(A1:A9,5.43)', A9), 0.381)
    near(ev('PERCENTRANK.EXC(A1:A9,5.43,1)', A9), 0.3)
    expect(ev('PERCENTRANK.EXC(A1:A9,5.43,0)', A9)).toEqual(E('#NUM!'))
    expect(ev('PERCENTRANK.EXC(A1:A9,10)', A9)).toEqual(E('#N/A'))
    expect(ev('PERCENTRANK.EXC(A1:A2,1)', ['a', 'b'])).toEqual(E('#NUM!'))
  })

  test('PERCENTRANK / PERCENTRANK.INC', () => {
    for (const fn of ['PERCENTRANK', 'PERCENTRANK.INC']) {
      near(ev(`${fn}(A1:A10,2)`, A10), 0.333)
      near(ev(`${fn}(A1:A10,4)`, A10), 0.555)
      near(ev(`${fn}(A1:A10,8)`, A10), 0.666)
      near(ev(`${fn}(A1:A10,5)`, A10), 0.583)
    }
    // the significance truncates (5/9 = 0.5555…), it does not round
    near(ev('PERCENTRANK.INC(A1:A10,4,2)', A10), 0.55)
    near(ev('PERCENTRANK.INC(A1:A10,4,5)', A10), 0.55555)
    expect(ev('PERCENTRANK.INC(A1:A10,4,0)', A10)).toEqual(E('#NUM!'))
    expect(ev('PERCENTRANK(A1:A10,14)', A10)).toEqual(E('#N/A'))
    expect(ev('PERCENTRANK(A1:A10,0)', A10)).toEqual(E('#N/A'))
    near(ev('PERCENTRANK(A1:A3,1)', [1, 2, 3]), 0)
    near(ev('PERCENTRANK(A1:A3,3)', [1, 2, 3]), 1)
  })

  test('RANK.AVG', () => {
    const temp = [89, 88, 92, 101, 94, 97, 95]
    near(ev('RANK.AVG(94,A1:A7)', temp), 4)
    near(ev('RANK.AVG(94,A1:A7,1)', temp), 4)
    near(ev('RANK.AVG(2,A1:A4)', [1, 2, 2, 3]), 2.5)
    near(ev('RANK.AVG(2,A1:A4,1)', [1, 2, 2, 3]), 2.5)
    near(ev('RANK.AVG(2,A1:A4,1)', [1, 2, 2, 2]), 3)
    expect(ev('RANK.AVG(50,A1:A4)', [1, 2, 2, 3])).toEqual(E('#N/A'))
    expect(ev('RANK.AVG("x",A1:A4)', [1, 2, 2, 3])).toEqual(E('#VALUE!'))
  })
})

describe('array results', () => {
  test('FREQUENCY', () => {
    const data = [79, 85, 78, 85, 50, 81, 95, 88, 97]
    expect(arr('FREQUENCY(A1:A9,B1:B3)', 4, 1, data, [70, 79, 89])).toEqual([[1], [2], [4], [2]])
    // no bins: all the data in the one cell
    expect(arr('FREQUENCY(A1:A3,{1,2})', 3, 1, [1, 2, 3])).toEqual([[1], [1], [1]])
    // text and blanks are ignored; data with no numbers: zeros
    expect(arr('FREQUENCY(A1:A3,{5})', 2, 1, ['a', null, 'c'])).toEqual([[0], [0]])
    expect(arr('FREQUENCY(A1:A2,{1,2})', 3, 1, [E('#NUM!'), 1])).toEqual([[E('#NUM!')], [E('#NUM!')], [E('#NUM!')]])
  })

  test('FREQUENCY with no bins counts the data', () => {
    expect(arr('FREQUENCY(A1:A3,B1:B2)', 1, 1, [1, 2, 3], ['x', 'y'])).toEqual([[3]])
  })

  test('MODE.MULT', () => {
    const d = [1, 2, 3, 4, 3, 2, 1, 2, 3, 5, 6, 1]
    expect(arr('MODE.MULT(A1:A12)', 8, 1, d)).toEqual([[1], [2], [3], [E('#N/A')], [E('#N/A')], [E('#N/A')], [E('#N/A')], [E('#N/A')]])
    expect(ev('MODE.MULT(A1:A12)', d)).toBe(1)
    expect(ev('MODE.MULT(A1:A3)', [1, 2, 3])).toEqual(E('#N/A'))
    expect(ev('MODE.MULT(A1:A3)', [E('#DIV/0!'), 2, 2])).toEqual(E('#DIV/0!'))
  })

  test('LINEST: slope and intercept, with const FALSE', () => {
    nearAll(arr('LINEST(A1:A4,B1:B4,,FALSE)', 1, 2, [1, 9, 5, 7], [0, 4, 2, 3]), [[2, 1]])
    // through the origin: y = 2x
    nearAll(arr('LINEST(A1:A3,B1:B3,FALSE)', 1, 2, [2, 4, 6], [1, 2, 3]), [[2, 0]])
    // default x = 1,2,3,…
    nearAll(arr('LINEST(A1:A3)', 1, 2, [3, 5, 7]), [[2, 1]])
  })

  test('LINEST: sum of the page (sales in month 9)', () => {
    const p = cols([1, 2, 3, 4, 5, 6], [3100, 4500, 4400, 5400, 7500, 8100])
    near(evOn('SUM(LINEST(B1:B6,A1:A6)*{9,1})', p), 11000, 6)
  })

  test('LINEST: multiple regression with statistics', () => {
    const x1 = [2310, 2333, 2356, 2379, 2402, 2425, 2448, 2471, 2494, 2517, 2540]
    const x2 = [2, 2, 3, 3, 2, 4, 2, 2, 3, 4, 2]
    const x3 = [2, 2, 1.5, 2, 3, 2, 1.5, 2, 3, 4, 3]
    const x4 = [20, 12, 33, 43, 53, 23, 99, 34, 23, 55, 22]
    const y = [142000, 144000, 151000, 150000, 139000, 169000, 126000, 142900, 163000, 169000, 149000]
    const got = arr('LINEST(E1:E11,A1:D11,TRUE,TRUE)', 5, 5, x1, x2, x3, x4, y)
    // the first column of the page's table: m4, se4, r2, F, ssreg
    near(got[0][0], -234.2371645, 6)
    near(got[1][0], 13.26801148, 6)
    near(got[2][0], 0.996747993, 8)
    near(got[3][0], 459.7536742, 5)
    expect(Math.abs((got[4][0] as number) - 1732393319)).toBeLessThan(1)
    // df, and the cells without a value
    near(got[3][1], 6)
    expect(got[2][2]).toEqual(E('#N/A'))
    expect(got[4][4]).toEqual(E('#N/A'))
    // ssreg / ssresid are consistent with r2
    const ssreg = got[4][0] as number
    const ssresid = got[4][1] as number
    near(got[2][0], ssreg / (ssreg + ssresid), 10)
    near(got[2][1], Math.sqrt(ssresid / 6), 6)
  })

  test('LINEST: exact multiple regression, coefficients in reverse order', () => {
    const x1 = [1, 2, 3, 4, 5, 6]
    const x2 = [2, 1, 5, 3, 9, 4]
    const y = x1.map((v, i) => 1 + 2 * v + 3 * x2[i])
    nearAll(arr('LINEST(A1:A6,B1:C6)', 1, 3, y, x1, x2), [[3, 2, 1]], 8)
    // variables in rows when y is a row
    const p = makeProvider({ Sheet1: { A1: y[0], B1: y[1], C1: y[2], D1: y[3], E1: y[4], F1: y[5], A2: 1, B2: 2, C2: 3, D2: 4, E2: 5, F2: 6, A3: 2, B3: 1, C3: 5, D3: 3, E3: 9, F3: 4 } })
    nearAll(evaluateArray(parseFormula('LINEST(A1:F1,A2:F3)'), ctx(p), 1, 3), [[3, 2, 1]], 8)
  })

  test('LINEST: a collinear column gets 0', () => {
    const x1 = [1, 2, 3, 4, 5]
    const x2 = x1.map((v) => 2 * v)
    const y = x1.map((v) => 1 + 2 * v)
    nearAll(arr('LINEST(A1:A5,B1:C5)', 1, 3, y, x1, x2), [[0, 2, 1]], 8)
  })

  test('LINEST: statistics without a constant', () => {
    const got = arr('LINEST(A1:A4,B1:B4,FALSE,TRUE)', 5, 2, [2.1, 3.9, 6.2, 7.8], [1, 2, 3, 4])
    expect(got[1][1]).toEqual(E('#N/A'))
    expect(got[0][1]).toBe(0)
    near(got[3][1], 3)
  })

  test('LINEST errors', () => {
    expect(ev('LINEST(A1:A3,B1:B2)', [1, 2, 3], [1, 2])).toEqual(E('#REF!'))
    expect(ev('LINEST(A1:A3,B1:B3)', [1, 'x', 3], [1, 2, 3])).toEqual(E('#VALUE!'))
    expect(ev('LINEST(A1:A3,B1:B3)', [1, 2, 3], [1, null, 3])).toEqual(E('#VALUE!'))
  })

  test('LOGEST / GROWTH: y = 2 * 3^x', () => {
    const y = [6, 18, 54, 162]
    nearAll(arr('LOGEST(A1:A4)', 1, 2, y), [[3, 2]], 8)
    nearAll(arr('LOGEST(A1:A4,B1:B4)', 1, 2, y, [1, 2, 3, 4]), [[3, 2]], 8)
    // b forced to 1
    nearAll(arr('LOGEST(A1:A3,B1:B3,FALSE)', 1, 2, [3, 9, 27], [1, 2, 3]), [[3, 1]], 8)
    const st = arr('LOGEST(A1:A4,,TRUE,TRUE)', 5, 2, y)
    near(st[2][0], 1, 8)
    near(st[3][1], 2)
    expect(arr('LOGEST(A1:A3)', 1, 2, [1, 0, 2])).toEqual([[E('#NUM!'), E('#NUM!')]])
    expect(arr('LOGEST(A1:A3)', 1, 2, [1, -1, 2])).toEqual([[E('#NUM!'), E('#NUM!')]])
  })

  test('GROWTH: the page example', () => {
    const month = [11, 12, 13, 14, 15, 16]
    const units = [33100, 47300, 69000, 102000, 150000, 220000]
    const fitted = arr('GROWTH(B1:B6,A1:A6)', 6, 1, month, units).map((r) => Math.round(r[0] as number))
    expect(fitted).toEqual([32618, 47729, 69841, 102197, 149542, 218822])
    const next = arr('GROWTH(B1:B6,A1:A6,C1:C2)', 2, 1, month, units, [17, 18]).map((r) => Math.round(r[0] as number))
    expect(next).toEqual([320197, 468536])
  })

  test('GROWTH: shapes and defaults', () => {
    nearAll(arr('GROWTH(A1:A4,,{5;6})', 2, 1, [6, 18, 54, 162]), [[486], [1458]], 6)
    nearAll(arr('GROWTH(A1:A4)', 4, 1, [6, 18, 54, 162]), [[6], [18], [54], [162]], 6)
    // new_x as a row gives a row
    nearAll(arr('GROWTH(A1:A4,B1:B4,{5,6})', 1, 2, [6, 18, 54, 162], [1, 2, 3, 4]), [[486, 1458]], 6)
    expect(arr('GROWTH(A1:A3)', 1, 1, [1, 0, 2])).toEqual([[E('#NUM!')]])
  })

  test('TREND', () => {
    nearAll(arr('TREND(A1:A4,B1:B4,{5;6})', 2, 1, [3, 5, 7, 9], [1, 2, 3, 4]), [[11], [13]], 8)
    // x omitted: 1,2,3…; new_x omitted: the fitted values
    nearAll(arr('TREND(A1:A4)', 4, 1, [3, 5, 7, 9]), [[3], [5], [7], [9]], 8)
    nearAll(arr('TREND(A1:A3,,{4})', 1, 1, [3, 5, 7]), [[9]], 8)
    // through the origin
    nearAll(arr('TREND(A1:A3,B1:B3,{4},FALSE)', 1, 1, [2, 4, 6], [1, 2, 3]), [[8]], 8)
    // a row of new x gives a row
    nearAll(arr('TREND(A1:A4,B1:B4,{5,6})', 1, 2, [3, 5, 7, 9], [1, 2, 3, 4]), [[11, 13]], 8)
    // two variables: y = 1 + 2*x1 + 3*x2
    const x1 = [1, 2, 3, 4, 5, 6]
    const x2 = [2, 1, 5, 3, 9, 4]
    const y = x1.map((v, i) => 1 + 2 * v + 3 * x2[i])
    nearAll(arr('TREND(A1:A6,B1:C6,{10,10;0,0})', 2, 1, y, x1, x2), [[51], [1]], 8)
    expect(arr('TREND(A1:A6,B1:C6,{1,2,3})', 1, 1, y, x1, x2)).toEqual([[E('#REF!')]])
    expect(arr('TREND(A1:A3,B1:B2)', 1, 1, [1, 2, 3], [1, 2])).toEqual([[E('#REF!')]])
  })
})

// ---------------------------------------------------------------- database functions

const TREES: Record<string, Value> = {
  A1: 'Tree', B1: 'Height', F1: 'Height',
  A2: '=Apple', B2: '>10', F2: '<16',
  A3: '=Pear',
  A5: 'Tree', B5: 'Height', C5: 'Age', D5: 'Yield', E5: 'Profit',
  A6: 'Apple', B6: 18, C6: 20, D6: 14, E6: 105,
  A7: 'Pear', B7: 12, C7: 12, D7: 10, E7: 96,
  A8: 'Cherry', B8: 13, C8: 14, D8: 9, E8: 105,
  A9: 'Apple', B9: 14, C9: 15, D9: 10, E9: 75,
  A10: 'Pear', B10: 9, C10: 8, D10: 8, E10: 77,
  A11: 'Apple', B11: 8, C11: 9, D11: 6, E11: 45,
}
const T = makeProvider({ Sheet1: TREES })
const d = (f: string) => evOn(f, T)

describe('database functions', () => {
  test('DSUM: the page examples', () => {
    expect(d('DSUM(A5:E11,"Profit",A1:A2)')).toBe(225)
    expect(d('DSUM(A5:E11,"Profit",A1:F3)')).toBe(248)
  })

  test('field by label (any case) or position', () => {
    expect(d('DSUM(A5:E11,"profit",A1:A2)')).toBe(225)
    expect(d('DSUM(A5:E11,5,A1:A2)')).toBe(225)
    expect(d('DSUM(A5:E11,"Nope",A1:A2)')).toEqual(E('#VALUE!'))
    expect(d('DSUM(A5:E11,6,A1:A2)')).toEqual(E('#VALUE!'))
    expect(d('DSUM(A5:E11,0,A1:A2)')).toEqual(E('#VALUE!'))
    expect(d('DSUM(A5:E11,TRUE,A1:A2)')).toEqual(E('#VALUE!'))
    expect(d('DSUM(A5:E11,,A1:A2)')).toEqual(E('#VALUE!'))
  })

  test('the other aggregates, on the apples', () => {
    expect(d('DAVERAGE(A5:E11,"Profit",A1:A2)')).toBe(75)
    expect(d('DMAX(A5:E11,"Yield",A1:A2)')).toBe(14)
    expect(d('DMIN(A5:E11,"Yield",A1:A2)')).toBe(6)
    expect(d('DPRODUCT(A5:E11,"Yield",A1:A2)')).toBe(840)
    near(d('DSTDEV(A5:E11,"Yield",A1:A2)'), 4)
    near(d('DSTDEVP(A5:E11,"Yield",A1:A2)'), Math.sqrt(32 / 3))
    near(d('DVAR(A5:E11,"Yield",A1:A2)'), 16)
    near(d('DVARP(A5:E11,"Yield",A1:A2)'), 32 / 3)
    expect(d('DCOUNT(A5:E11,"Age",A1:A2)')).toBe(3)
    expect(d('DCOUNTA(A5:E11,"Tree",A1:A2)')).toBe(3)
    // rows ORed: apples of 10-16 feet and all pears
    expect(d('DCOUNT(A5:E11,"Age",A1:F3)')).toBe(3)
  })

  test('DCOUNT / DCOUNTA with a blank field count the records', () => {
    expect(d('DCOUNT(A5:E11,,A1:A2)')).toBe(3)
    expect(d('DCOUNTA(A5:E11,,A1:A2)')).toBe(3)
  })

  test('DCOUNT counts numbers, DCOUNTA non-empty cells', () => {
    const p = makeProvider({
      Sheet1: { A1: 'k', B1: 'v', A2: 'x', B2: 1, A3: 'x', B3: 'text', A4: 'x', A5: 'x', B5: true, D1: 'k', D2: 'x' },
    })
    expect(evOn('DCOUNT(A1:B5,"v",D1:D2)', p)).toBe(1)
    expect(evOn('DCOUNTA(A1:B5,"v",D1:D2)', p)).toBe(3)
  })

  test('DGET', () => {
    const p = makeProvider({ Sheet1: { ...TREES, H1: 'Tree', H2: '=Cherry', J1: 'Tree', J2: '=Plum' } })
    expect(evOn('DGET(A5:E11,"Profit",H1:H2)', p)).toBe(105)
    expect(evOn('DGET(A5:E11,"Profit",J1:J2)', p)).toEqual(E('#VALUE!'))
    expect(evOn('DGET(A5:E11,"Profit",A1:A2)', p)).toEqual(E('#NUM!'))
    expect(evOn('DGET(A5:E11,"Tree",H1:H2)', p)).toBe('Cherry')
  })

  test('text without an operator matches the start; ="=text" is exact', () => {
    const p = makeProvider({ Sheet1: { ...TREES, H1: 'Tree', H2: 'App', J1: 'Tree', J2: 'a', L1: 'Tree', L2: '=App', N1: 'Tree', N2: 'a*e' } })
    expect(evOn('DSUM(A5:E11,"Profit",H1:H2)', p)).toBe(225)
    expect(evOn('DSUM(A5:E11,"Profit",J1:J2)', p)).toBe(225)
    expect(evOn('DSUM(A5:E11,"Profit",L1:L2)', p)).toBe(0)
    expect(evOn('DSUM(A5:E11,"Profit",N1:N2)', p)).toBe(225)
  })

  test('operators, equal numbers, AND within a row', () => {
    const p = makeProvider({
      Sheet1: { ...TREES, H1: 'Age', H2: '>=14', J1: 'Age', K1: 'Yield', J2: '>=14', K2: '<10', L1: 'Height', L2: 13, N1: 'Tree', N2: '<>Apple' },
    })
    expect(evOn('DSUM(A5:E11,"Profit",H1:H2)', p)).toBe(105 + 105 + 75)
    expect(evOn('DSUM(A5:E11,"Profit",J1:K2)', p)).toBe(105)
    expect(evOn('DSUM(A5:E11,"Profit",L1:L2)', p)).toBe(105)
    expect(evOn('DCOUNT(A5:E11,"Profit",N1:N2)', p)).toBe(3)
  })

  test('an empty condition does not restrict', () => {
    const p = makeProvider({ Sheet1: { ...TREES, H1: 'Tree', H2: null, H3: 'Pear' } })
    expect(evOn('DSUM(A5:E11,"Profit",H1:H2)', p)).toBe(105 + 96 + 105 + 75 + 77 + 45)
    const q = makeProvider({ Sheet1: { ...TREES, H1: 'Tree', J1: 'Age', H2: 'Pear', J2: null } })
    expect(evOn('DSUM(A5:E11,"Profit",H1:J2)', q)).toBe(96 + 77)
  })

  test('criteria labels are matched without regard to case; unknown labels match nothing', () => {
    const p = makeProvider({ Sheet1: { ...TREES, H1: 'TREE', H2: '=Pear', J1: 'Colour', J2: 'red' } })
    expect(evOn('DSUM(A5:E11,"Profit",H1:H2)', p)).toBe(96 + 77)
    expect(evOn('DSUM(A5:E11,"Profit",J1:J2)', p)).toBe(0)
  })

  test('errors and empty results', () => {
    near(d('DAVERAGE(A5:E11,"Profit",A1:A3)'), (105 + 75 + 45 + 96 + 77) / 5)
    const none = makeProvider({ Sheet1: { ...TREES, H1: 'Tree', H2: '=Plum' } })
    expect(evOn('DAVERAGE(A5:E11,"Profit",H1:H2)', none)).toEqual(E('#DIV/0!'))
    expect(evOn('DSTDEV(A5:E11,"Profit",H1:H2)', none)).toEqual(E('#DIV/0!'))
    expect(evOn('DSTDEVP(A5:E11,"Profit",H1:H2)', none)).toEqual(E('#DIV/0!'))
    expect(evOn('DVAR(A5:E11,"Profit",H1:H2)', none)).toEqual(E('#DIV/0!'))
    expect(evOn('DMAX(A5:E11,"Profit",H1:H2)', none)).toBe(0)
    expect(evOn('DMIN(A5:E11,"Profit",H1:H2)', none)).toBe(0)
    expect(evOn('DSUM(A5:E11,"Profit",H1:H2)', none)).toBe(0)
    expect(evOn('DCOUNT(A5:E11,"Profit",H1:H2)', none)).toBe(0)
    // one record: the sample statistics need two
    const one = makeProvider({ Sheet1: { ...TREES, H1: 'Tree', H2: '=Cherry' } })
    expect(evOn('DSTDEV(A5:E11,"Profit",H1:H2)', one)).toEqual(E('#DIV/0!'))
    expect(evOn('DVAR(A5:E11,"Profit",H1:H2)', one)).toEqual(E('#DIV/0!'))
    expect(evOn('DVARP(A5:E11,"Profit",H1:H2)', one)).toBe(0)
    // a criteria range with no condition row
    expect(d('DSUM(A5:E11,"Profit",A1:A1)')).toEqual(E('#VALUE!'))
    // an error in the field
    const bad = makeProvider({ Sheet1: { ...TREES, E6: E('#N/A') } })
    expect(evOn('DSUM(A5:E11,"Profit",A1:A2)', bad)).toEqual(E('#N/A'))
  })
})
