import { describe, expect, test } from 'bun:test'
import { ymdToSerial } from '../src/core/xlsx/formula/dates'
import { days360, yearfrac } from '../src/core/xlsx/formula/daycount'
import { betaI, erf, erfc, gamma, gammaln, gammaP, gammaQ, invert, normCdf, normInv } from '../src/core/xlsx/formula/special'

const near = (a: number, b: number, rel = 1e-13) => expect(Math.abs(a - b)).toBeLessThanOrEqual(Math.abs(b) * rel + 1e-300)

describe('special functions', () => {
  test('erf / erfc', () => {
    near(erf(0.5), 0.5204998778130465)
    near(erf(1), 0.8427007929497149)
    near(erf(2), 0.9953222650189527)
    near(erf(3), 0.9999779095030014)
    near(erf(-1), -0.8427007929497149)
    near(erfc(0.25), 0.7236736098317631)
    near(erfc(3), 2.209049699858544e-5)
    near(erfc(5), 1.5374597944280347e-12)
    near(erfc(10), 2.088487583762545e-45)
    near(erfc(-1), 1.8427007929497148)
  })
  test('gamma', () => {
    near(gammaln(0.5), 0.5723649429247001)
    near(gammaln(10), 12.801827480081469)
    near(gamma(0.5), 1.772453850905516)
    near(gamma(5.5), 52.34277778455352)
    near(gamma(-1.5), 2.363271801207355)
    expect(gamma(5)).toBe(24)
    near(gamma(170.5) / gamma(169.5), 169.5, 1e-12)
  })
  test('incomplete gamma / beta', () => {
    near(gammaP(1, 1), 1 - Math.exp(-1))
    near(gammaP(2.5, 3), 0.6937810815867216, 1e-12)
    near(gammaQ(0.5, 50), erfc(Math.sqrt(50)))
    near(betaI(0.5, 2, 3), 0.6875)
    near(betaI(0.2, 0.5, 0.5), (2 / Math.PI) * Math.asin(Math.sqrt(0.2)))
    near(betaI(0.9, 30, 2), 0.9 ** 30 * 4, 1e-12)
  })
  test('normal law', () => {
    near(normCdf(1.96), 0.9750021048517795)
    near(normInv(0.975), 1.959963984540054)
    near(normInv(1e-10), -6.361340902404056)
    expect(Math.abs(normInv(0.5))).toBeLessThan(1e-16)
    near(invert(normCdf, 0.975, -Infinity, Infinity), 1.959963984540054, 1e-12)
  })
  test('day counts', () => {
    const d = (y: number, m: number, dd: number) => ymdToSerial(y, m, dd)
    near(yearfrac(d(2012, 1, 1), d(2012, 7, 30), 0), 0.58055556, 1e-8)
    near(yearfrac(d(2012, 1, 1), d(2012, 7, 30), 1), 0.57650273, 1e-8)
    near(yearfrac(d(2012, 1, 1), d(2012, 7, 30), 3), 0.57808219, 1e-8)
    expect(days360(d(2011, 1, 30), d(2011, 2, 1), false)).toBe(1)
    expect(days360(d(2011, 1, 1), d(2011, 12, 31), false)).toBe(360)
    expect(days360(d(2011, 1, 1), d(2011, 2, 28), false)).toBe(57)
    expect(days360(d(2011, 2, 28), d(2011, 3, 31), false)).toBe(30)
    expect(days360(d(2011, 1, 31), d(2011, 3, 31), true)).toBe(60)
  })
})
