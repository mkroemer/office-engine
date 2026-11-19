import { describe, expect, test } from 'bun:test'
import { colIndex, evaluateFormula, FormulaError, type CellProvider, type ErrorCode, type Value } from '../src/core/xlsx/formula'

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

// data of the worked examples of the Microsoft pages
const DATA: Sheets = {
  Sheet1: {
    // Z.TEST
    A1: 3, A2: 6, A3: 7, A4: 8, A5: 6, A6: 5, A7: 4, A8: 2, A9: 1, A10: 9,
    // T.TEST
    C1: 3, C2: 4, C3: 5, C4: 8, C5: 9, C6: 1, C7: 2, C8: 4, C9: 5,
    D1: 6, D2: 19, D3: 3, D4: 2, D5: 14, D6: 4, D7: 5, D8: 17, D9: 1,
    // F.TEST
    F1: 6, F2: 7, F3: 9, F4: 15, F5: 21,
    G1: 20, G2: 28, G3: 31, G4: 38, G5: 40,
    // CHISQ.TEST: actual, expected
    H1: 58, I1: 35, H2: 11, I2: 25, H3: 10, I3: 23,
    K1: 45.35, L1: 47.65, K2: 17.56, L2: 18.44, K3: 16.09, L3: 16.91,
    // PROB
    M1: 0, M2: 1, M3: 2, M4: 3,
    N1: 0.2, N2: 0.3, N3: 0.1, N4: 0.4,
    // text and blanks among numbers (skipped)
    P1: 3, P2: 'x', P3: 4, /* P4 empty */ P5: 8, P6: 9,
    Q1: 1, Q2: 'y', Q3: 2, Q4: 3,
    R1: 4, R2: 8, R3: 6, R4: 9,
    // a short column, a constant column
    S1: 1, S2: 2, S3: 3,
    T1: 5, T2: 5, T3: 5,
  },
}
const P = makeProvider(DATA)

function ev(f: string): Value {
  return evaluateFormula(f, { provider: P, sheet: 'Sheet1', r: 30, c: 9 })
}

/** [formula, value printed on the Microsoft page, decimals of that value] */
type Ex = [string, number, number]

function examples(cases: Ex[]) {
  for (const [f, want, digits] of cases) {
    const got = ev(f)
    expect(typeof got, f).toBe('number')
    expect(got as number, f).toBeCloseTo(want, digits)
  }
}

function errors(cases: [string, ErrorCode][]) {
  for (const [f, code] of cases) expect(ev(f), f).toEqual(E(code))
}

/** relative closeness */
function rel(f: string, want: number, tol = 1e-11) {
  const got = ev(f)
  expect(typeof got, f).toBe('number')
  expect(Math.abs((got as number) - want) / Math.max(Math.abs(want), 1e-300), f).toBeLessThan(tol)
}

// ---------------------------------------------------------------- normal

describe('normal family', () => {
  test('Microsoft examples', () =>
    examples([
      ['NORM.DIST(42,40,1.5,TRUE)', 0.9087888, 7],
      ['NORM.DIST(42,40,1.5,FALSE)', 0.10934, 5],
      ['NORMDIST(42,40,1.5,TRUE)', 0.9087888, 7],
      ['NORMDIST(42,40,1.5,FALSE)', 0.10934, 5],
      ['NORM.INV(0.908789,40,1.5)', 42.000002, 6],
      ['NORMINV(0.908789,40,1.5)', 42.000002, 6],
      ['NORM.S.DIST(1.333333,TRUE)', 0.908788726, 9],
      ['NORM.S.DIST(1.333333,FALSE)', 0.164010148, 9],
      ['NORMSDIST(1.333333)', 0.908788726, 9],
      ['NORM.S.INV(0.908789)', 1.3333347, 7],
      ['NORMSINV(0.9088)', 1.3334, 4],
      ['GAUSS(2)', 0.47725, 5],
      ['PHI(0.75)', 0.301137432, 9],
      ['STANDARDIZE(42,40,1.5)', 1.33333333, 8],
      ['LOGNORM.DIST(4,3.5,1.2,TRUE)', 0.0390836, 7],
      ['LOGNORM.DIST(4,3.5,1.2,FALSE)', 0.0176176, 7],
      ['LOGNORMDIST(4,3.5,1.2)', 0.0390836, 7],
      ['LOGNORM.INV(0.039084,3.5,1.2)', 4.0000252, 6],
      ['LOGINV(0.039084,3.5,1.2)', 4.0000252, 6],
      ['CONFIDENCE(0.05,2.5,50)', 0.692951912, 8],
      ['CONFIDENCE.NORM(0.05,2.5,50)', 0.692952, 6],
      ['CONFIDENCE.T(0.05,1,50)', 0.284196855, 8],
      ['FISHER(0.75)', 0.9729551, 7],
      ['FISHERINV(0.972955)', 0.75, 6],
    ]))

  test('Z.TEST / ZTEST', () => {
    examples([
      ['Z.TEST(A1:A10,4)', 0.090574, 6],
      ['2 * MIN(Z.TEST(A1:A10,4), 1 - Z.TEST(A1:A10,4))', 0.181148, 6],
      ['Z.TEST(A1:A10,6)', 0.863043, 6],
      ['2 * MIN(Z.TEST(A1:A10,6), 1 - Z.TEST(A1:A10,6))', 0.273913, 6],
      ['ZTEST(A1:A10,4)', 0.090574, 6],
      ['ZTEST(A1:A10,6)', 0.863043, 6],
    ])
    // with a known sigma: P(mean > observed) = 1 - Φ((mean - x) / (sigma / √n)); mean of A1:A10 is 5.1
    rel('Z.TEST(A1:A10,4,2)', 1 - (ev('NORM.S.DIST((5.1-4)/(2/SQRT(10)),TRUE)') as number), 1e-9)
    // text and blanks are skipped: P1:P6 holds 3, 4, 8, 9
    expect(ev('Z.TEST(P1:P6,5)')).toBeCloseTo(ev('Z.TEST({3,4,8,9},5)') as number, 12)
  })

  test('identities and accuracy', () => {
    for (const p of [1e-10, 0.001, 0.3, 0.5, 0.9, 0.999999]) {
      rel(`NORM.S.DIST(NORM.S.INV(${p}),TRUE)`, p, 1e-11)
      rel(`NORM.DIST(NORM.INV(${p},10,3),10,3,TRUE)`, p, 1e-11)
      rel(`LOGNORM.DIST(LOGNORM.INV(${p},1,0.5),1,0.5,TRUE)`, p, 1e-11)
    }
    expect(ev('GAUSS(0)')).toBe(0)
    expect(ev('NORM.S.INV(0.5)')).toBe(0)
    rel('FISHERINV(FISHER(0.3))', 0.3, 1e-14)
    rel('GAUSS(1.5)', (ev('NORM.S.DIST(1.5,TRUE)') as number) - 0.5, 1e-12)
    // the standard normal as a special case of NORM.DIST
    expect(ev('NORM.DIST(0.7,0,1,TRUE)')).toBe(ev('NORM.S.DIST(0.7,TRUE)'))
    // far tail keeps its digits
    expect(ev('NORM.S.DIST(-10,TRUE)') as number).toBeGreaterThan(7.6e-24)
    expect(ev('NORM.S.DIST(-10,TRUE)') as number).toBeLessThan(7.7e-24)
  })

  test('errors', () =>
    errors([
      ['NORM.DIST(1,0,0,TRUE)', '#NUM!'],
      ['NORM.DIST(1,0,-1,FALSE)', '#NUM!'],
      ['NORM.DIST("a",0,1,TRUE)', '#VALUE!'],
      ['NORM.DIST(1,0,1,"maybe")', '#VALUE!'],
      ['NORMDIST(1,"b",1,TRUE)', '#VALUE!'],
      ['NORM.INV(0,0,1)', '#NUM!'],
      ['NORM.INV(1,0,1)', '#NUM!'],
      ['NORM.INV(0.5,0,0)', '#NUM!'],
      ['NORMINV(1.5,0,1)', '#NUM!'],
      ['NORM.S.INV(0)', '#NUM!'],
      ['NORMSINV(1)', '#NUM!'],
      ['NORMSDIST("x")', '#VALUE!'],
      ['GAUSS("x")', '#VALUE!'],
      ['PHI("x")', '#VALUE!'],
      ['STANDARDIZE(1,2,0)', '#NUM!'],
      ['STANDARDIZE(1,2,-3)', '#NUM!'],
      ['LOGNORM.DIST(0,0,1,TRUE)', '#NUM!'],
      ['LOGNORM.DIST(-1,0,1,TRUE)', '#NUM!'],
      ['LOGNORM.DIST(1,0,0,TRUE)', '#NUM!'],
      ['LOGNORMDIST(0,0,1)', '#NUM!'],
      ['LOGNORM.INV(0,0,1)', '#NUM!'],
      ['LOGINV(1,0,1)', '#NUM!'],
      ['LOGINV(0.5,0,0)', '#NUM!'],
      ['CONFIDENCE(0,1,10)', '#NUM!'],
      ['CONFIDENCE(1,1,10)', '#NUM!'],
      ['CONFIDENCE(0.05,0,10)', '#NUM!'],
      ['CONFIDENCE.NORM(0.05,1,0)', '#NUM!'],
      ['CONFIDENCE.T(0.05,1,1)', '#DIV/0!'],
      ['CONFIDENCE.T(0.05,1,0)', '#NUM!'],
      ['CONFIDENCE.T(0,1,10)', '#NUM!'],
      ['CONFIDENCE.T(0.05,-1,10)', '#NUM!'],
      ['Z.TEST(A20:A25,4)', '#N/A'],
      ['Z.TEST(S1:S1,4)', '#DIV/0!'],
      ['Z.TEST(T1:T3,4)', '#DIV/0!'],
      ['Z.TEST(A1:A10,4,0)', '#NUM!'],
      ['FISHER(1)', '#NUM!'],
      ['FISHER(-1)', '#NUM!'],
      ['FISHER(2)', '#NUM!'],
      ['FISHER("x")', '#VALUE!'],
      ['FISHERINV("x")', '#VALUE!'],
    ]))

  test('truncation and coercion', () => {
    // the size of CONFIDENCE is truncated
    expect(ev('CONFIDENCE(0.05,2.5,50.9)')).toBe(ev('CONFIDENCE(0.05,2.5,50)'))
    expect(ev('NORM.S.DIST(1,1)')).toBe(ev('NORM.S.DIST(1,TRUE)'))
    expect(ev('NORM.S.DIST(1,0)')).toBe(ev('NORM.S.DIST(1,FALSE)'))
    // arrays lift
    expect(ev('SUM(NORM.S.DIST({0,0},TRUE))')).toBe(1)
  })
})

// ---------------------------------------------------------------- Student t

describe('Student t family', () => {
  test('Microsoft examples', () =>
    examples([
      ['T.DIST(60,1,TRUE)', 0.99469533, 8],
      ['T.DIST(8,3,FALSE)', 0.00073691, 8],
      ['T.DIST.2T(1.959999998,60)', 0.054645, 6],
      ['T.DIST.RT(1.959999998,60)', 0.027322, 6],
      ['T.INV(0.75,2)', 0.8164966, 7],
      ['T.INV.2T(0.546449,60)', 0.606533, 6],
      // the page prints 2.28139 in a remark, a typo for the known 2.228139
      ['T.INV.2T(0.05,10)', 2.228139, 6],
      ['TDIST(1.959999998,60,2)', 0.05464493, 8],
      ['TDIST(1.959999998,60,1)', 0.027322465, 9],
      ['TINV(0.05464,60)', 1.96, 2],
      ['TINV(0.05,10)', 2.228139, 6],
      ['T.TEST(C1:C9,D1:D9,2,1)', 0.196016, 6],
      ['TTEST(C1:C9,D1:D9,2,1)', 0.19602, 5],
    ]))

  test('identities', () => {
    // df = 1 is the Cauchy law: right tail = 1/2 - atan(x)/π
    rel('T.DIST.RT(3,1)', 0.5 - Math.atan(3) / Math.PI, 1e-13)
    rel('T.DIST(-3,1,TRUE)', 0.5 - Math.atan(3) / Math.PI, 1e-13)
    // df = 2: cdf = 1/2 + x / (2 √(2 + x²))
    rel('T.DIST(1.5,2,TRUE)', 0.5 + 1.5 / (2 * Math.sqrt(2 + 2.25)), 1e-13)
    for (const [x, d] of [
      [0.3, 5],
      [2.5, 12],
      [-1.7, 3],
      [8, 40],
    ]) {
      rel(`T.DIST.2T(${Math.abs(x)},${d})`, 2 * (ev(`T.DIST.RT(${Math.abs(x)},${d})`) as number), 1e-13)
      expect((ev(`T.DIST(${x},${d},TRUE)`) as number) + (ev(`T.DIST.RT(${x},${d})`) as number)).toBeCloseTo(1, 14)
    }
    for (const [p, d] of [
      [0.75, 2],
      [0.01, 7],
      [0.999, 30],
      [1e-8, 4],
    ]) {
      rel(`T.DIST(T.INV(${p},${d}),${d},TRUE)`, p, 1e-10)
      rel(`T.DIST.2T(T.INV.2T(${p},${d}),${d})`, p, 1e-10)
    }
    expect(ev('T.INV(0.5,9)')).toBe(0)
    expect(ev('T.INV.2T(1,9)')).toBe(0)
    // symmetric
    expect(ev('T.INV(0.1,6)')).toBe(-(ev('T.INV(0.9,6)') as number))
    // TINV is two-tailed: TINV(2p) is the one-tailed value T.INV(1 - p)
    rel('TINV(0.1,8)', ev('T.INV(0.95,8)') as number, 1e-12)
  })

  test('degrees of freedom are truncated', () => {
    expect(ev('T.DIST(1,5.9,TRUE)')).toBe(ev('T.DIST(1,5,TRUE)'))
    expect(ev('T.INV.2T(0.05,10.7)')).toBe(ev('T.INV.2T(0.05,10)'))
    expect(ev('TDIST(1.5,4.8,1.9)')).toBe(ev('TDIST(1.5,4,1)'))
  })

  test('T.TEST types', () => {
    // two-sample, equal variance: C1:C9 mean 4.5556, D1:D9 mean 7.8889
    const two = ev('T.TEST(C1:C9,D1:D9,2,2)') as number
    expect(two).toBeGreaterThan(0)
    expect(two).toBeLessThan(1)
    expect(ev('T.TEST(C1:C9,D1:D9,1,2)')).toBeCloseTo(two / 2, 14)
    const welch = ev('T.TEST(C1:C9,D1:D9,2,3)') as number
    expect(welch).toBeGreaterThan(0)
    expect(welch).toBeLessThan(1)
    // swapping the arrays gives the same probability
    expect(ev('T.TEST(D1:D9,C1:C9,2,3)')).toBeCloseTo(welch, 14)
    // same sizes and variances: Welch = pooled
    expect(ev('T.TEST({1,2,3,4},{3,4,5,6},2,3)')).toBeCloseTo(ev('T.TEST({1,2,3,4},{3,4,5,6},2,2)') as number, 12)
    // identical samples
    expect(ev('T.TEST({1,2,3,4},{1,2,3,5},2,2)')).toBeGreaterThan(0.5)
    // paired t on the differences = one-sample t against 0 (tdist of t = mean/(sd/√n))
    const d = [3 - 6, 4 - 19, 5 - 3, 8 - 2, 9 - 14, 1 - 4, 2 - 5, 4 - 17, 5 - 1]
    const m = d.reduce((a, b) => a + b) / d.length
    const sd = Math.sqrt(d.reduce((a, b) => a + (b - m) ** 2, 0) / (d.length - 1))
    rel(`T.DIST.2T(${Math.abs(m) / (sd / 3)},8)`, ev('T.TEST(C1:C9,D1:D9,2,1)') as number, 1e-12)
    // text and blanks are skipped
    expect(ev('T.TEST(P1:P6,{3,4,8,9},2,2)')).toBeCloseTo(1, 12)
  })

  test('errors', () =>
    errors([
      ['T.DIST(1,0,TRUE)', '#NUM!'],
      ['T.DIST(1,0.9,FALSE)', '#NUM!'],
      ['T.DIST("a",5,TRUE)', '#VALUE!'],
      ['T.DIST.2T(-1,5)', '#NUM!'],
      ['T.DIST.2T(1,0)', '#NUM!'],
      ['T.DIST.RT(1,0)', '#NUM!'],
      ['T.INV(0,5)', '#NUM!'],
      ['T.INV(1,5)', '#NUM!'],
      ['T.INV(0.5,0)', '#NUM!'],
      ['T.INV.2T(0,5)', '#NUM!'],
      ['T.INV.2T(1.1,5)', '#NUM!'],
      ['T.INV.2T(0.5,0)', '#NUM!'],
      ['TINV(0,5)', '#NUM!'],
      ['TINV(0.5,0)', '#NUM!'],
      ['TDIST(-1,5,1)', '#NUM!'],
      ['TDIST(1,5,3)', '#NUM!'],
      ['TDIST(1,5,0)', '#NUM!'],
      ['TDIST(1,0,1)', '#NUM!'],
      ['TDIST(1,5,"a")', '#VALUE!'],
      ['T.TEST(C1:C9,D1:D9,3,1)', '#NUM!'],
      ['T.TEST(C1:C9,D1:D9,2,4)', '#NUM!'],
      ['T.TEST(C1:C9,D1:D9,"a",1)', '#VALUE!'],
      ['T.TEST(C1:C9,D1:D5,2,1)', '#N/A'],
      ['TTEST(C1:C9,D1:D5,2,1)', '#N/A'],
      ['T.TEST(S1:S1,S1:S1,2,2)', '#DIV/0!'],
      ['T.TEST(T1:T3,T1:T3,2,2)', '#DIV/0!'],
      ['T.TEST(T1:T3,T1:T3,2,1)', '#DIV/0!'],
    ]))

  test('unpaired tests accept different sizes', () => {
    expect(typeof ev('T.TEST(C1:C9,D1:D5,2,2)')).toBe('number')
    expect(typeof ev('T.TEST(C1:C9,D1:D5,2,3)')).toBe('number')
  })
})

// ---------------------------------------------------------------- chi-square

describe('chi-square family', () => {
  test('Microsoft examples', () =>
    examples([
      ['CHISQ.DIST(0.5,1,TRUE)', 0.52049988, 8],
      ['CHISQ.DIST(2,3,FALSE)', 0.20755375, 8],
      ['CHISQ.DIST.RT(18.307,10)', 0.0500006, 7],
      ['CHIDIST(18.307,10)', 0.0500006, 7],
      ['CHISQ.INV(0.93,1)', 3.283020286, 8],
      ['CHISQ.INV(0.6,2)', 1.832581464, 9],
      ['CHISQ.INV.RT(0.050001,10)', 18.306973, 6],
      ['CHIINV(0.050001,10)', 18.306973, 6],
      ['CHISQ.TEST(H1:I3,K1:L3)', 0.0003082, 7],
      ['CHITEST(H1:I3,K1:L3)', 0.0003082, 7],
    ]))

  test('identities', () => {
    // df = 2: cdf = 1 - e^(-x/2)
    rel('CHISQ.DIST(3,2,TRUE)', 1 - Math.exp(-1.5), 1e-14)
    rel('CHISQ.DIST.RT(40,2)', Math.exp(-20), 1e-13)
    rel('CHISQ.DIST(3,2,FALSE)', Math.exp(-1.5) / 2, 1e-13)
    // df = 1: P(χ² < x) = erf(√(x/2))
    rel('CHISQ.DIST(2,1,TRUE)', ev('2*NORM.S.DIST(SQRT(2),TRUE)-1') as number, 1e-13)
    expect((ev('CHISQ.DIST(7,5,TRUE)') as number) + (ev('CHISQ.DIST.RT(7,5)') as number)).toBeCloseTo(1, 14)
    for (const [p, k] of [
      [0.93, 1],
      [0.001, 20],
      [1e-12, 3],
      [0.999999, 8],
      [0.5, 100],
    ]) {
      rel(`CHISQ.DIST(CHISQ.INV(${p},${k}),${k},TRUE)`, p, 1e-10)
      rel(`CHISQ.DIST.RT(CHISQ.INV.RT(${p},${k}),${k})`, p, 1e-10)
    }
    expect(ev('CHISQ.INV(0,4)')).toBe(0)
    expect(ev('CHISQ.INV.RT(1,4)')).toBe(0)
    // a gamma law in disguise
    rel('GAMMA.DIST(7,2.5,2,TRUE)', ev('CHISQ.DIST(7,5,TRUE)') as number, 1e-13)
    // CHISQ.TEST of a column against its mean: 1 degree of freedom per row - 1
    rel('CHISQ.TEST({10;20},{15;15})', ev('CHISQ.DIST.RT(2*25/15,1)') as number, 1e-13)
    // CHISQ.TEST on a row
    rel('CHISQ.TEST({10,20,30},{20,20,20})', ev('CHISQ.DIST.RT(10,2)') as number, 1e-13)
  })

  test('degrees of freedom are truncated', () => {
    expect(ev('CHISQ.DIST(2,3.9,TRUE)')).toBe(ev('CHISQ.DIST(2,3,TRUE)'))
    expect(ev('CHIINV(0.1,4.5)')).toBe(ev('CHIINV(0.1,4)'))
  })

  test('errors', () =>
    errors([
      ['CHISQ.DIST(-1,3,TRUE)', '#NUM!'],
      ['CHISQ.DIST(1,0,TRUE)', '#NUM!'],
      ['CHISQ.DIST(1,10000000001,TRUE)', '#NUM!'],
      ['CHISQ.DIST("a",3,TRUE)', '#VALUE!'],
      ['CHISQ.DIST.RT(-1,3)', '#NUM!'],
      ['CHISQ.DIST.RT(1,0)', '#NUM!'],
      ['CHIDIST(-1,3)', '#NUM!'],
      ['CHIDIST(1,0)', '#NUM!'],
      ['CHISQ.INV(-0.1,3)', '#NUM!'],
      ['CHISQ.INV(1.1,3)', '#NUM!'],
      ['CHISQ.INV(0.5,0)', '#NUM!'],
      ['CHISQ.INV.RT(-0.1,3)', '#NUM!'],
      ['CHISQ.INV.RT(1.1,3)', '#NUM!'],
      ['CHISQ.INV.RT(0.5,0)', '#NUM!'],
      ['CHIINV(2,3)', '#NUM!'],
      ['CHISQ.TEST(H1:I3,K1:K3)', '#N/A'],
      ['CHITEST(H1:I3,K1:L2)', '#N/A'],
      ['CHISQ.TEST(S1:S1,S1:S1)', '#N/A'],
      ['CHISQ.TEST({1,2},{0,1})', '#DIV/0!'],
      ['CHISQ.TEST({1,"a"},{1,2})', '#VALUE!'],
    ]))
})

// ---------------------------------------------------------------- F

describe('F family', () => {
  test('Microsoft examples', () =>
    examples([
      ['F.DIST(15.2069,6,4,TRUE)', 0.99, 5],
      ['F.DIST(15.2069,6,4,FALSE)', 0.0012238, 7],
      ['F.DIST.RT(15.2068649,6,4)', 0.01, 7],
      ['F.INV(0.01,6,4)', 0.10930991, 8],
      ['F.INV.RT(0.01,6,4)', 15.20686, 5],
      ['F.TEST(F1:F5,G1:G5)', 0.64831785, 8],
      ['FDIST(15.20686486,6,4)', 0.01, 7],
      ['FINV(0.01,6,4)', 15.206865, 6],
      ['FTEST(F1:F5,G1:G5)', 0.64831785, 8],
    ]))

  test('identities', () => {
    expect((ev('F.DIST(2.5,6,4,TRUE)') as number) + (ev('F.DIST.RT(2.5,6,4)') as number)).toBeCloseTo(1, 14)
    // F(1, n) is the square of a t(n): P(F < x) = 1 - T.DIST.2T(√x, n)
    rel('F.DIST.RT(4,1,9)', ev('T.DIST.2T(2,9)') as number, 1e-13)
    // F(2, 2): cdf = x / (1 + x)
    rel('F.DIST(3,2,2,TRUE)', 0.75, 1e-14)
    for (const [p, a, b] of [
      [0.01, 6, 4],
      [0.9, 3, 12],
      [1e-9, 5, 5],
      [0.999999, 2, 30],
    ]) {
      rel(`F.DIST(F.INV(${p},${a},${b}),${a},${b},TRUE)`, p, 1e-10)
      rel(`F.DIST.RT(F.INV.RT(${p},${a},${b}),${a},${b})`, p, 1e-10)
    }
    // the right tail of F.INV.RT is the left of F.INV
    rel('F.INV.RT(0.1,5,7)', ev('F.INV(0.9,5,7)') as number, 1e-11)
    expect(ev('F.INV(0,3,4)')).toBe(0)
    expect(ev('F.INV.RT(1,3,4)')).toBe(0)
    // F.TEST of a sample with itself: ratio 1, the probability is 1
    expect(ev('F.TEST({1,2,4,8},{1,2,4,8})')).toBeCloseTo(1, 12)
    // text and blanks are skipped
    expect(ev('F.TEST(P1:P6,{3,4,8,9})')).toBeCloseTo(1, 12)
  })

  test('degrees of freedom are truncated', () => {
    expect(ev('F.DIST(2,3.9,4.2,TRUE)')).toBe(ev('F.DIST(2,3,4,TRUE)'))
    expect(ev('FINV(0.1,3.5,4.5)')).toBe(ev('FINV(0.1,3,4)'))
  })

  test('errors', () =>
    errors([
      ['F.DIST(-1,3,4,TRUE)', '#NUM!'],
      ['F.DIST(1,0,4,TRUE)', '#NUM!'],
      ['F.DIST(1,3,0,FALSE)', '#NUM!'],
      ['F.DIST("a",3,4,TRUE)', '#VALUE!'],
      ['F.DIST.RT(-1,3,4)', '#NUM!'],
      ['F.DIST.RT(1,0,4)', '#NUM!'],
      ['FDIST(1,3,10000000000)', '#NUM!'],
      ['F.INV(-0.1,3,4)', '#NUM!'],
      ['F.INV(1.1,3,4)', '#NUM!'],
      ['F.INV(0.5,0,4)', '#NUM!'],
      ['F.INV.RT(2,3,4)', '#NUM!'],
      ['F.INV.RT(0.5,3,0)', '#NUM!'],
      ['FINV(0.5,0,4)', '#NUM!'],
      ['F.TEST(S1:S1,F1:F5)', '#DIV/0!'],
      ['F.TEST(T1:T3,F1:F5)', '#DIV/0!'],
      ['FTEST(F1:F5,T1:T3)', '#DIV/0!'],
    ]))
})

// ---------------------------------------------------------------- beta, gamma

describe('beta and gamma family', () => {
  test('Microsoft examples', () =>
    examples([
      ['BETA.DIST(2,8,10,TRUE,1,3)', 0.6854706, 7],
      ['BETA.DIST(2,8,10,FALSE,1,3)', 1.4837646, 7],
      ['BETA.INV(0.685470581,8,10,1,3)', 2, 8],
      ['BETADIST(2,8,10,1,3)', 0.6854706, 7],
      ['BETAINV(0.685470581,8,10,1,3)', 2, 8],
      ['GAMMA(2.5)', 1.329, 3],
      ['GAMMA(-3.75)', 0.268, 3],
      ['GAMMA.DIST(10.00001131,9,2,FALSE)', 0.032639, 6],
      ['GAMMA.DIST(10.00001131,9,2,TRUE)', 0.068094, 6],
      ['GAMMADIST(10.00001131,9,2,FALSE)', 0.032639, 6],
      ['GAMMADIST(10.00001131,9,2,TRUE)', 0.068094, 6],
      ['GAMMA.INV(0.068094,9,2)', 10.0000112, 7],
      ['GAMMAINV(0.068094,9,2)', 10.0000112, 7],
      ['GAMMALN(4)', 1.7917595, 7],
      ['GAMMALN.PRECISE(4)', 1.7917595, 7],
    ]))

  test('identities', () => {
    // defaults A = 0, B = 1
    expect(ev('BETA.DIST(0.3,2,5,TRUE)')).toBe(ev('BETA.DIST(0.3,2,5,TRUE,0,1)'))
    expect(ev('BETADIST(0.3,2,5)')).toBe(ev('BETA.DIST(0.3,2,5,TRUE,0,1)'))
    expect(ev('BETAINV(0.3,2,5)')).toBe(ev('BETA.INV(0.3,2,5,0,1)'))
    // uniform
    rel('BETA.DIST(0.37,1,1,TRUE)', 0.37, 1e-14)
    rel('BETA.DIST(0.37,1,1,FALSE)', 1, 1e-14)
    // a beta(2, 2): cdf = 3x² - 2x³
    rel('BETA.DIST(0.3,2,2,TRUE)', 3 * 0.09 - 2 * 0.027, 1e-13)
    rel('BETA.DIST(0.3,2,2,FALSE)', 6 * 0.3 * 0.7, 1e-13)
    // edges of the interval
    expect(ev('BETA.DIST(1,2,3,TRUE,1,5)')).toBe(0)
    expect(ev('BETA.DIST(5,2,3,TRUE,1,5)')).toBe(1)
    expect(ev('BETA.DIST(0,1,3,FALSE)')).toBe(3)
    for (const p of [1e-9, 0.01, 0.5, 0.99, 0.999999]) {
      rel(`BETA.DIST(BETA.INV(${p},2.5,4),2.5,4,TRUE)`, p, 1e-10)
      rel(`GAMMA.DIST(GAMMA.INV(${p},3.5,2),3.5,2,TRUE)`, p, 1e-10)
    }
    rel('BETA.INV(0.25,3,4,10,20)', 10 + 10 * (ev('BETA.INV(0.25,3,4)') as number), 1e-12)
    expect(ev('BETA.INV(1,3,4,10,20)')).toBe(20)
    expect(ev('GAMMA.INV(0,3,2)')).toBe(0)
    // alpha = 1: exponential with mean beta
    rel('GAMMA.DIST(3,1,2,TRUE)', 1 - Math.exp(-1.5), 1e-14)
    rel('GAMMA.DIST(3,1,2,FALSE)', Math.exp(-1.5) / 2, 1e-14)
    expect(ev('GAMMA.DIST(0,1,2,FALSE)')).toBe(0.5)
    expect(ev('GAMMA.DIST(0,3,2,FALSE)')).toBe(0)
    // gamma function
    rel('GAMMA(5)', 24, 1e-14)
    rel('GAMMA(0.5)', Math.sqrt(Math.PI), 1e-13)
    rel('GAMMA(1.5)', Math.sqrt(Math.PI) / 2, 1e-13)
    rel('GAMMALN(10)', Math.log(362880), 1e-13)
    rel('EXP(GAMMALN(6))', 120, 1e-12)
    rel('GAMMALN(0.5)', Math.log(Math.sqrt(Math.PI)), 1e-13)
  })

  test('errors', () =>
    errors([
      ['BETA.DIST(0.5,0,2,TRUE)', '#NUM!'],
      ['BETA.DIST(0.5,2,-1,TRUE)', '#NUM!'],
      ['BETA.DIST(2,2,2,TRUE)', '#NUM!'],
      ['BETA.DIST(-1,2,2,TRUE)', '#NUM!'],
      ['BETA.DIST(0.5,2,2,TRUE,1,3)', '#NUM!'],
      ['BETA.DIST(2,2,2,TRUE,2,2)', '#NUM!'],
      ['BETA.DIST(0.5,2,2,"x")', '#VALUE!'],
      ['BETA.DIST(0,0.5,2,FALSE)', '#NUM!'],
      ['BETADIST(5,2,2,1,3)', '#NUM!'],
      ['BETADIST(2,2,0)', '#NUM!'],
      ['BETA.INV(0,2,2)', '#NUM!'],
      ['BETA.INV(1.1,2,2)', '#NUM!'],
      ['BETA.INV(0.5,0,2)', '#NUM!'],
      ['BETA.INV(0.5,2,2,3,3)', '#NUM!'],
      ['BETAINV(0.5,2,-2)', '#NUM!'],
      ['GAMMA(0)', '#NUM!'],
      ['GAMMA(-2)', '#NUM!'],
      ['GAMMA("x")', '#VALUE!'],
      ['GAMMA(200)', '#NUM!'],
      ['GAMMA.DIST(-1,2,2,TRUE)', '#NUM!'],
      ['GAMMA.DIST(1,0,2,TRUE)', '#NUM!'],
      ['GAMMA.DIST(1,2,0,FALSE)', '#NUM!'],
      ['GAMMA.DIST(0,0.5,1,FALSE)', '#NUM!'],
      ['GAMMADIST(1,2,-1,TRUE)', '#NUM!'],
      ['GAMMA.INV(-0.1,2,2)', '#NUM!'],
      ['GAMMA.INV(1,2,2)', '#NUM!'],
      ['GAMMA.INV(0.5,0,2)', '#NUM!'],
      ['GAMMAINV(0.5,2,0)', '#NUM!'],
      ['GAMMALN(0)', '#NUM!'],
      ['GAMMALN(-1)', '#NUM!'],
      ['GAMMALN.PRECISE(-1)', '#NUM!'],
      ['GAMMALN("x")', '#VALUE!'],
    ]))
})

// ---------------------------------------------------------------- discrete

describe('discrete distributions', () => {
  test('Microsoft examples', () =>
    examples([
      ['BINOM.DIST(6,10,0.5,FALSE)', 0.2050781, 7],
      ['BINOMDIST(6,10,0.5,FALSE)', 0.2050781, 7],
      ['BINOM.DIST.RANGE(60,0.75,48)', 0.084, 3],
      ['BINOM.DIST.RANGE(60,0.75,45,50)', 0.524, 3],
      ['BINOM.INV(6,0.5,0.75)', 4, 12],
      ['CRITBINOM(6,0.5,0.75)', 4, 12],
      ['HYPGEOM.DIST(1,4,8,20,TRUE)', 0.4654, 4],
      ['HYPGEOM.DIST(1,4,8,20,FALSE)', 0.3633, 4],
      ['HYPGEOMDIST(1,4,8,20)', 0.3633, 4],
      ['NEGBINOM.DIST(10,5,0.25,TRUE)', 0.3135141, 7],
      ['NEGBINOM.DIST(10,5,0.25,FALSE)', 0.0550487, 7],
      ['NEGBINOMDIST(10,5,0.25)', 0.05504866, 8],
      ['POISSON(2,5,TRUE)', 0.124652, 6],
      ['POISSON(2,5,FALSE)', 0.084224, 6],
      ['POISSON.DIST(2,5,TRUE)', 0.124652, 6],
      ['POISSON.DIST(2,5,FALSE)', 0.084224, 6],
      ['EXPON.DIST(0.2,10,TRUE)', 0.86466472, 8],
      ['EXPON.DIST(0.2,10,FALSE)', 1.35335283, 8],
      ['EXPONDIST(0.2,10,TRUE)', 0.86466472, 8],
      ['EXPONDIST(0.2,10,FALSE)', 1.35335283, 8],
      ['WEIBULL(105,20,100,TRUE)', 0.929581, 6],
      ['WEIBULL(105,20,100,FALSE)', 0.035589, 6],
      ['WEIBULL.DIST(105,20,100,TRUE)', 0.929581, 6],
      ['WEIBULL.DIST(105,20,100,FALSE)', 0.035589, 6],
      ['PERMUT(100,3)', 970200, 9],
      ['PERMUT(3,2)', 6, 12],
      ['PERMUTATIONA(3,2)', 9, 12],
      ['PERMUTATIONA(2,2)', 4, 12],
    ]))

  test('PROB', () => {
    examples([
      ['PROB(M1:M4,N1:N4,2)', 0.1, 12],
      ['PROB(M1:M4,N1:N4,1,3)', 0.8, 12],
    ])
    expect(ev('PROB(M1:M4,N1:N4,5)')).toBe(0)
    expect(ev('PROB(M1:M4,N1:N4,-1,10)')).toBeCloseTo(1, 12)
    expect(ev('PROB({1,2,3},{0.2,0.5,0.3},2,3)')).toBeCloseTo(0.8, 12)
    // a probability may be 0
    expect(ev('PROB({1,2},{0,1},2)')).toBe(1)
    errors([
      ['PROB(M1:M4,N1:N3,2)', '#N/A'],
      ['PROB(M1:M4,{0.2,0.3,0.1,0.5},2)', '#NUM!'],
      ['PROB(M1:M4,{0.2,0.3,-0.1,0.6},2)', '#NUM!'],
      ['PROB(M1:M4,{0.2,0.3,1.1,0.4},2)', '#NUM!'],
      ['PROB(M1:M4,N1:N4,"a")', '#VALUE!'],
    ])
  })

  test('binomial identities', () => {
    // pmf sums to the cdf
    let sum = 0
    for (let k = 0; k <= 7; k++) sum += ev(`BINOM.DIST(${k},12,0.3,FALSE)`) as number
    expect(ev('BINOM.DIST(7,12,0.3,TRUE)') as number).toBeCloseTo(sum, 13)
    expect(ev('BINOM.DIST(12,12,0.3,TRUE)')).toBe(1)
    expect(ev('BINOM.DIST(0,12,0,FALSE)')).toBe(1)
    expect(ev('BINOM.DIST(12,12,1,FALSE)')).toBe(1)
    expect(ev('BINOM.DIST(3,12,1,TRUE)')).toBe(0)
    // closed form: C(5,2) / 2^5
    rel('BINOM.DIST(2,5,0.5,FALSE)', 10 / 32, 1e-14)
    // large n goes through the logarithm
    rel('BINOM.DIST(501,1000,0.5,FALSE)', (ev('BINOM.DIST(500,1000,0.5,FALSE)') as number) * (500 / 501), 1e-11)
    // the number of successes and trials are truncated
    expect(ev('BINOM.DIST(6.9,10.9,0.5,FALSE)')).toBe(ev('BINOM.DIST(6,10,0.5,FALSE)'))
    // range
    rel('BINOM.DIST.RANGE(10,0.4,2,5)', (ev('BINOM.DIST(5,10,0.4,TRUE)') as number) - (ev('BINOM.DIST(1,10,0.4,TRUE)') as number), 1e-12)
    expect(ev('BINOM.DIST.RANGE(10,0.4,0,10)')).toBeCloseTo(1, 13)
    expect(ev('BINOM.DIST.RANGE(10,0.4,3)')).toBe(ev('BINOM.DIST(3,10,0.4,FALSE)'))
    // inverse: the smallest k reaching alpha
    for (const [n, p, a] of [
      [20, 0.3, 0.5],
      [100, 0.1, 0.9],
      [10, 0.5, 0.05],
    ]) {
      const k = ev(`BINOM.INV(${n},${p},${a})`) as number
      expect(ev(`BINOM.DIST(${k},${n},${p},TRUE)`) as number).toBeGreaterThanOrEqual(a)
      if (k > 0) expect(ev(`BINOM.DIST(${k - 1},${n},${p},TRUE)`) as number).toBeLessThan(a)
    }
    // alpha equal to a cdf value hits that value (cdf(0) = 0.5 for 1 trial)
    expect(ev('BINOM.INV(1,0.5,0.5)')).toBe(0)
    expect(ev('BINOM.INV(6.9,0.5,0.75)')).toBe(4)
  })

  test('other discrete laws', () => {
    // hypergeometric cdf is the sum of the pmf
    let sum = 0
    for (let k = 0; k <= 3; k++) sum += ev(`HYPGEOM.DIST(${k},6,10,25,FALSE)`) as number
    expect(ev('HYPGEOM.DIST(3,6,10,25,TRUE)') as number).toBeCloseTo(sum, 13)
    // C(3,1) C(7,1) / C(10,2) = 21 / 45
    rel('HYPGEOM.DIST(1,2,3,10,FALSE)', 21 / 45, 1e-14)
    // large population goes through the logarithm
    rel('HYPGEOM.DIST(2,4,600,2000,FALSE)', ev('BINOM.DIST(2,4,0.3,FALSE)') as number, 2e-2)
    // sample_s below the lower bound n - (N - K) is an error even if >= 0
    expect(ev('HYPGEOM.DIST(0,5,8,10,FALSE)')).toEqual(E('#NUM!'))
    // negative binomial: p^r when no failure
    rel('NEGBINOM.DIST(0,3,0.4,FALSE)', 0.064, 1e-14)
    // geometric (1 success): cdf = 1 - (1-p)^(f+1)
    rel('NEGBINOM.DIST(4,1,0.3,TRUE)', 1 - 0.7 ** 5, 1e-13)
    expect(ev('NEGBINOM.DIST(0,3,1,FALSE)')).toBe(1)
    expect(ev('NEGBINOM.DIST(2,3,1,FALSE)')).toBe(0)
    expect(ev('NEGBINOM.DIST(2,3,0,FALSE)')).toBe(0)
    // truncation
    expect(ev('NEGBINOM.DIST(10.9,5.9,0.25,FALSE)')).toBe(ev('NEGBINOM.DIST(10,5,0.25,FALSE)'))
    // Poisson
    rel('POISSON.DIST(0,3,FALSE)', Math.exp(-3), 1e-14)
    rel('POISSON.DIST(3,3,TRUE)', Math.exp(-3) * (1 + 3 + 4.5 + 4.5), 1e-13)
    expect(ev('POISSON.DIST(2.9,5,FALSE)')).toBe(ev('POISSON.DIST(2,5,FALSE)'))
    expect(ev('POISSON.DIST(0,0,FALSE)')).toBe(1)
    expect(ev('POISSON.DIST(3,0,FALSE)')).toBe(0)
    expect(ev('POISSON.DIST(3,0,TRUE)')).toBe(1)
    rel('POISSON.DIST(200,150,FALSE)', ev('EXP(-150+200*LN(150)-GAMMALN(201))') as number, 1e-12)
    // exponential and Weibull
    rel('EXPON.DIST(2,0.5,TRUE)', 1 - Math.exp(-1), 1e-14)
    rel('WEIBULL.DIST(3,1,2,TRUE)', 1 - Math.exp(-1.5), 1e-14)
    rel('WEIBULL.DIST(3,1,2,FALSE)', Math.exp(-1.5) / 2, 1e-14)
    expect(ev('WEIBULL.DIST(0,1,2,FALSE)')).toBe(0.5)
    expect(ev('WEIBULL.DIST(0,2,2,FALSE)')).toBe(0)
    expect(ev('EXPON.DIST(0,2,FALSE)')).toBe(2)
    // permutations
    expect(ev('PERMUT(5,0)')).toBe(1)
    expect(ev('PERMUT(5.9,2.9)')).toBe(20)
    expect(ev('PERMUTATIONA(0,0)')).toBe(1)
    expect(ev('PERMUTATIONA(0,3)')).toBe(0)
    expect(ev('PERMUTATIONA(2.9,3.9)')).toBe(8)
  })

  test('errors', () =>
    errors([
      ['BINOM.DIST(-1,10,0.5,TRUE)', '#NUM!'],
      ['BINOM.DIST(11,10,0.5,TRUE)', '#NUM!'],
      ['BINOM.DIST(1,10,1.5,FALSE)', '#NUM!'],
      ['BINOM.DIST(1,10,-0.5,FALSE)', '#NUM!'],
      ['BINOM.DIST(1,-1,0.5,FALSE)', '#NUM!'],
      ['BINOM.DIST("a",10,0.5,TRUE)', '#VALUE!'],
      ['BINOMDIST(11,10,0.5,TRUE)', '#NUM!'],
      ['BINOM.DIST.RANGE(10,0.5,11)', '#NUM!'],
      ['BINOM.DIST.RANGE(10,0.5,-1)', '#NUM!'],
      ['BINOM.DIST.RANGE(10,1.5,2)', '#NUM!'],
      ['BINOM.DIST.RANGE(10,0.5,5,4)', '#NUM!'],
      ['BINOM.DIST.RANGE(10,0.5,5,11)', '#NUM!'],
      ['BINOM.INV(-1,0.5,0.5)', '#NUM!'],
      ['BINOM.INV(10,0,0.5)', '#NUM!'],
      ['BINOM.INV(10,1,0.5)', '#NUM!'],
      ['BINOM.INV(10,0.5,0)', '#NUM!'],
      ['BINOM.INV(10,0.5,1)', '#NUM!'],
      ['CRITBINOM(10,0.5,1)', '#NUM!'],
      ['HYPGEOM.DIST(-1,4,8,20,TRUE)', '#NUM!'],
      ['HYPGEOM.DIST(5,4,8,20,TRUE)', '#NUM!'],
      ['HYPGEOM.DIST(9,10,8,20,TRUE)', '#NUM!'],
      ['HYPGEOM.DIST(1,0,8,20,TRUE)', '#NUM!'],
      ['HYPGEOM.DIST(1,21,8,20,TRUE)', '#NUM!'],
      ['HYPGEOM.DIST(0,4,0,20,TRUE)', '#NUM!'],
      ['HYPGEOM.DIST(1,4,21,20,TRUE)', '#NUM!'],
      ['HYPGEOM.DIST(1,4,8,0,TRUE)', '#NUM!'],
      ['HYPGEOMDIST(5,4,8,20)', '#NUM!'],
      ['NEGBINOM.DIST(10,5,1.5,TRUE)', '#NUM!'],
      ['NEGBINOM.DIST(10,5,-0.5,TRUE)', '#NUM!'],
      ['NEGBINOM.DIST(-1,5,0.5,TRUE)', '#NUM!'],
      ['NEGBINOM.DIST(1,0,0.5,TRUE)', '#NUM!'],
      ['NEGBINOMDIST(1,0.5,0.5)', '#NUM!'],
      ['POISSON(-1,5,TRUE)', '#NUM!'],
      ['POISSON(2,-5,TRUE)', '#NUM!'],
      ['POISSON.DIST(2,"a",TRUE)', '#VALUE!'],
      ['EXPON.DIST(-1,2,TRUE)', '#NUM!'],
      ['EXPON.DIST(1,0,TRUE)', '#NUM!'],
      ['EXPONDIST(1,-1,FALSE)', '#NUM!'],
      ['WEIBULL(-1,2,2,TRUE)', '#NUM!'],
      ['WEIBULL(1,0,2,TRUE)', '#NUM!'],
      ['WEIBULL.DIST(1,2,0,FALSE)', '#NUM!'],
      ['WEIBULL.DIST(0,0.5,2,FALSE)', '#NUM!'],
      ['PERMUT(0,0)', '#NUM!'],
      ['PERMUT(-1,2)', '#NUM!'],
      ['PERMUT(3,-1)', '#NUM!'],
      ['PERMUT(3,4)', '#NUM!'],
      ['PERMUT("a",1)', '#VALUE!'],
      ['PERMUTATIONA(-1,2)', '#NUM!'],
      ['PERMUTATIONA(2,-1)', '#NUM!'],
      ['PERMUTATIONA(10,400)', '#NUM!'],
    ]))
})

// ---------------------------------------------------------------- Google Sheets

describe('MARGINOFERROR', () => {
  test('z · s / √n', () => {
    // R1:R4 = 4, 8, 6, 9: mean 6.75, sample sd √(14.75/3)
    const sd = Math.sqrt(14.75 / 3)
    rel('MARGINOFERROR(R1:R4,0.95)', (ev('NORM.S.INV(0.975)') as number) * sd / 2, 1e-12)
    // it is CONFIDENCE.NORM for the sample sd
    rel('MARGINOFERROR(R1:R4,0.9)', ev(`CONFIDENCE.NORM(0.1,${sd},4)`) as number, 1e-12)
    // text and blanks are skipped
    expect(ev('MARGINOFERROR(Q1:Q4,0.95)')).toBe(ev('MARGINOFERROR({1,2,3},0.95)'))
  })

  test('errors', () =>
    errors([
      ['MARGINOFERROR(R1:R4,0)', '#NUM!'],
      ['MARGINOFERROR(R1:R4,1)', '#NUM!'],
      ['MARGINOFERROR(S1:S1,0.95)', '#DIV/0!'],
      ['MARGINOFERROR(R1:R4,"x")', '#VALUE!'],
    ]))
})

// ---------------------------------------------------------------- plumbing

describe('plumbing', () => {
  test('errors in arguments propagate', () => {
    expect(ev('NORM.DIST(1/0,0,1,TRUE)')).toEqual(E('#DIV/0!'))
    expect(ev('T.DIST(#N/A,5,TRUE)')).toEqual(E('#N/A'))
  })

  test('array arguments lift element-wise', () => {
    expect(ev('SUM(BINOM.DIST({0,1,2},2,0.5,FALSE))')).toBeCloseTo(1, 14)
    expect(ev('SUM(CHISQ.DIST.RT({1,2,3},2))')).toBeCloseTo(Math.exp(-0.5) + Math.exp(-1) + Math.exp(-1.5), 13)
  })

  test('a non-finite result is #NUM!', () => {
    expect(ev('PERMUTATIONA(10,400)')).toEqual(E('#NUM!'))
    expect(ev('GAMMA(172)')).toEqual(E('#NUM!'))
  })

  test('range arguments: the numbers only, direct values count', () => {
    expect(ev('PROB({1,2,3},{0.2,0.5,0.3},1)')).toBeCloseTo(0.2, 14)
    expect(typeof ev('F.TEST(P1:P6,R1:R4)')).toBe('number')
  })
})
