/** Math functions. */
import { A, INF, opt, S, visit, type Ev, type FnDef } from './fn'
import { roundHalf } from './format'
import { ERR, FormulaError, type Value } from './types'
import { int, num, toNum, type Val } from './values'

/** Remove binary noise (15 significant digits). */
const clean = (x: number) => Number(x.toPrecision(15))

function roundDir(x: number, d: number, up: boolean): number {
  if (x === 0) return 0
  const ax = Math.abs(x)
  const to = up ? Math.ceil : Math.floor
  let r: number
  if (d >= 0) {
    const y = ax * Math.pow(10, d)
    // no digit left to round away (1E300 to 10 decimals)
    if (!Number.isFinite(y) || y >= 2 ** 52) return x
    r = to(clean(y)) / Math.pow(10, d)
  } else {
    const f = Math.pow(10, -d)
    // 10^400 is not a number: what is left at that scale is 0
    if (!Number.isFinite(f)) return 0
    r = to(clean(ax / f)) * f
    if (!Number.isFinite(r)) return 0
  }
  return x < 0 ? -r : r
}

function finite(x: number): number {
  if (!Number.isFinite(x)) throw ERR.NUM
  return x
}

/** multiple of `sig` via ceil/floor of the (noise-free) quotient */
const multiple = (x: number, sig: number, up: boolean) => clean((up ? Math.ceil : Math.floor)(clean(x / sig)) * sig)

function gcd2(a: number, b: number): number {
  while (b) [a, b] = [b, a % b]
  return a
}

/** the integers of GCD / LCM: an empty cell is 0, text that is not a number (even in a range) is #VALUE! */
function intArgs(args: Val[], ev: Ev): number[] {
  const out: number[] = []
  const add = (v: Value) => {
    const n = toNum(v)
    if (n instanceof FormulaError) throw n
    if (n < 0) throw ERR.NUM
    out.push(Math.trunc(n))
  }
  for (const a of args) visit(a, ev, (v) => (typeof v === 'boolean' ? add(ERR.VALUE) : add(v)), add)
  return out
}

export const MATH: Record<string, FnDef> = {
  ABS: S(1, 1, ([x]) => Math.abs(num(x))),
  ROUND: S(2, 2, ([x, d]) => roundHalf(num(x), int(d))),
  ROUNDUP: S(2, 2, ([x, d]) => roundDir(num(x), int(d), true)),
  ROUNDDOWN: S(2, 2, ([x, d]) => roundDir(num(x), int(d), false)),
  TRUNC: S(1, 2, (a) => roundDir(num(a[0]), Math.trunc(opt(a, 1, 0)), false)),
  INT: S(1, 1, ([x]) => Math.floor(num(x))),
  MOD: S(2, 2, ([x, y]) => {
    const n = num(x)
    const d = num(y)
    if (d === 0) throw ERR.DIV0
    return clean(n - d * Math.floor(n / d))
  }),
  POWER: S(2, 2, ([x, y]) => {
    const b = num(x)
    const e = num(y)
    if (b === 0 && e === 0) throw ERR.NUM
    if (b === 0 && e < 0) throw ERR.DIV0
    const r = Math.pow(b, e)
    if (Number.isNaN(r)) throw ERR.NUM
    return finite(r)
  }),
  SQRT: S(1, 1, ([x]) => {
    const n = num(x)
    if (n < 0) throw ERR.NUM
    return Math.sqrt(n)
  }),
  EXP: S(1, 1, ([x]) => finite(Math.exp(num(x)))),
  LN: S(1, 1, ([x]) => {
    const n = num(x)
    if (n <= 0) throw ERR.NUM
    return Math.log(n)
  }),
  LOG: S(1, 2, (a) => {
    const n = num(a[0])
    const b = opt(a, 1, 10)
    if (n <= 0 || b <= 0) throw ERR.NUM
    if (b === 1) throw ERR.DIV0
    return clean(Math.log(n) / Math.log(b))
  }),
  LOG10: S(1, 1, ([x]) => {
    const n = num(x)
    if (n <= 0) throw ERR.NUM
    return Math.log10(n)
  }),
  PI: A(0, 0, () => Math.PI),
  SIGN: S(1, 1, ([x]) => Math.sign(num(x))),
  CEILING: S(1, 2, (a) => {
    const x = num(a[0])
    const sig = opt(a, 1, 1)
    if (sig === 0 || x === 0) return 0
    if (x > 0 && sig < 0) throw ERR.NUM
    return multiple(x, sig, true)
  }),
  FLOOR: S(1, 2, (a) => {
    const x = num(a[0])
    const sig = opt(a, 1, 1)
    if (sig === 0) throw ERR.DIV0
    if (x > 0 && sig < 0) throw ERR.NUM
    return multiple(x, sig, false)
  }),
  'CEILING.MATH': S(1, 3, (a) => {
    const x = num(a[0])
    const sig = Math.abs(opt(a, 1, 1))
    if (sig === 0 || x === 0) return 0
    if (x < 0 && opt(a, 2, 0) !== 0) return -multiple(-x, sig, true)
    return multiple(x, sig, true)
  }),
  'FLOOR.MATH': S(1, 3, (a) => {
    const x = num(a[0])
    const sig = Math.abs(opt(a, 1, 1))
    if (sig === 0 || x === 0) return 0
    if (x < 0 && opt(a, 2, 0) !== 0) return -multiple(-x, sig, false)
    return multiple(x, sig, false)
  }),
  RAND: A(0, 0, () => Math.random()),
  RANDBETWEEN: S(2, 2, ([a, b]) => {
    const lo = Math.ceil(num(a))
    const hi = Math.floor(num(b))
    if (lo > hi) throw ERR.NUM
    return lo + Math.floor(Math.random() * (hi - lo + 1))
  }),
  FACT: S(1, 1, ([x]) => {
    const v = num(x)
    // 171! is past what a number holds: no need to count up to it (nor to 1E15)
    if (v < 0 || v >= 171) throw ERR.NUM
    const n = Math.trunc(v)
    let r = 1
    for (let i = 2; i <= n; i++) r *= i
    return finite(r)
  }),
  EVEN: S(1, 1, ([x]) => {
    const n = num(x)
    const v = Math.ceil(clean(Math.abs(n) / 2)) * 2
    return n < 0 ? -v : v
  }),
  ODD: S(1, 1, ([x]) => {
    const n = num(x)
    let v = Math.ceil(clean(Math.abs(n)))
    if (v % 2 === 0) v++
    return n < 0 ? -v : v
  }),
  GCD: A(1, INF, (a, ev) => intArgs(a, ev).reduce((g, n) => gcd2(g, n), 0)),
  LCM: A(1, INF, (a, ev) =>
    finite(intArgs(a, ev).reduce((l, n) => (l === 0 || n === 0 ? 0 : (l / gcd2(l, n)) * n), 1)),
  ),
}
