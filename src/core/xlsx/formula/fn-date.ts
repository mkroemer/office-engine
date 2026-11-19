/** Date & time functions (serial numbers; 1900 system with Lotus bug or 1904). */
import { dateToSerial, daysInMonth, serialSeconds, serialToYmd, ymdToSerial } from './dates'
import { A, nums, opt, S, type FnDef } from './fn'
import { ERR, type Value } from './types'
import { num, str } from './values'

/** 31/12/9999: past it Excel has no date (a code like 20240115 is not one) */
const MAX_SERIAL = 2958466

function serial(v: Value): number {
  const n = num(v)
  if (n < 0 || n >= MAX_SERIAL) throw ERR.NUM
  return n
}
const day = (v: Value) => Math.floor(serial(v))

function addMonths(start: number, months: number, eom: boolean): number {
  const p = serialToYmd(start)
  const total = p.y * 12 + (p.m - 1) + Math.trunc(months)
  const y = Math.floor(total / 12)
  const m = (total % 12) + 1
  const d = eom ? daysInMonth(y, m) : Math.min(p.d, daysInMonth(y, m))
  const s = ymdToSerial(y, m, d)
  if (s < 0) throw ERR.NUM
  return s
}

const WEEKDAY_START: Record<number, [number, number]> = {
  // type: [day the week starts on (0=Sun), number of the first day]
  1: [0, 1],
  2: [1, 1],
  3: [1, 0],
  11: [1, 1],
  12: [2, 1],
  13: [3, 1],
  14: [4, 1],
  15: [5, 1],
  16: [6, 1],
  17: [0, 1],
}

function datedif(s: number, e: number, unit: string): number {
  if (s > e) throw ERR.NUM
  const a = serialToYmd(s)
  const b = serialToYmd(e)
  const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0)
  switch (unit.toUpperCase()) {
    case 'D':
      return e - s
    case 'M':
      return months
    case 'Y':
      return Math.floor(months / 12)
    case 'YM':
      return months % 12
    case 'MD': {
      if (b.d >= a.d) return b.d - a.d
      const pm = b.m === 1 ? 12 : b.m - 1
      const py = b.m === 1 ? b.y - 1 : b.y
      return daysInMonth(py, pm) - a.d + b.d
    }
    case 'YD': {
      // the anniversary of 29 Feb in a year without one is 1 March (DATE overflows the day), not 28 Feb
      let anniv = ymdToSerial(b.y, a.m, a.d)
      if (anniv > e) anniv = ymdToSerial(b.y - 1, a.m, a.d)
      return e - anniv
    }
  }
  throw ERR.NUM
}

function networkdays(s: number, e: number, holidays: Set<number>): number {
  const sign = s <= e ? 1 : -1
  const [lo, hi] = s <= e ? [s, e] : [e, s]
  // whole weeks by count, the days left one by one (never the span day by day: it can be as long as the numbers go)
  const days = hi - lo + 1
  let n = Math.floor(days / 7) * 5
  for (let d = hi - (days % 7) + 1; d <= hi; d++) {
    const dow = serialToYmd(d).dow
    if (dow !== 0 && dow !== 6) n++
  }
  for (const h of holidays) {
    if (h < lo || h > hi) continue
    const dow = serialToYmd(h).dow
    if (dow !== 0 && dow !== 6) n--
  }
  return sign * n
}

export const DATE: Record<string, FnDef> = {
  TODAY: A(0, 0, (_a, ev) => Math.floor(dateToSerial(ev.now()))),
  NOW: A(0, 0, (_a, ev) => dateToSerial(ev.now())),
  DATE: S(3, 3, ([y, m, d]) => {
    let yy = Math.trunc(num(y))
    if (yy >= 0 && yy < 1900) yy += 1900
    if (yy < 0 || yy > 9999) throw ERR.NUM
    const s = ymdToSerial(yy, Math.trunc(num(m)), Math.trunc(num(d)))
    if (s < 0) throw ERR.NUM
    return s
  }),
  TIME: S(3, 3, ([h, m, s]) => {
    const t = Math.trunc(num(h)) * 3600 + Math.trunc(num(m)) * 60 + Math.trunc(num(s))
    if (t < 0) throw ERR.NUM
    return (t % 86400) / 86400
  }),
  YEAR: S(1, 1, ([x]) => serialToYmd(day(x)).y),
  MONTH: S(1, 1, ([x]) => serialToYmd(day(x)).m),
  DAY: S(1, 1, ([x]) => serialToYmd(day(x)).d),
  HOUR: S(1, 1, ([x]) => Math.floor(serialSeconds(serial(x)) / 3600)),
  MINUTE: S(1, 1, ([x]) => Math.floor((serialSeconds(serial(x)) % 3600) / 60)),
  SECOND: S(1, 1, ([x]) => serialSeconds(serial(x)) % 60),
  WEEKDAY: S(1, 2, (a) => {
    const dow = serialToYmd(day(a[0])).dow
    const spec = WEEKDAY_START[Math.trunc(opt(a, 1, 1))]
    if (!spec) throw ERR.NUM
    return ((dow - spec[0] + 7) % 7) + spec[1]
  }),
  EDATE: S(2, 2, ([s, m]) => addMonths(day(s), num(m), false)),
  EOMONTH: S(2, 2, ([s, m]) => addMonths(day(s), num(m), true)),
  DATEDIF: S(3, 3, ([s, e, u]) => datedif(day(s), day(e), str(u))),
  DAYS: S(2, 2, ([e, s]) => Math.floor(num(e)) - Math.floor(num(s))),
  NETWORKDAYS: A(2, 3, (a, ev) => {
    const hol = new Set(a.length > 2 ? nums([a[2]], ev).map(Math.floor) : [])
    return networkdays(day(ev.scalar(a[0])), day(ev.scalar(a[1])), hol)
  }),
}
