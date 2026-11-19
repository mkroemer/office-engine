/** Function registry (upper-case names, `_xlfn.` prefixes already stripped by the lexer). */
import { DATE } from './fn-date'
import { DIST } from './fn-dist'
import { ENGINEERING } from './fn-engineering'
import { FINANCIAL } from './fn-financial'
import { LOGIC } from './fn-logic'
import { LOOKUP } from './fn-lookup'
import { MATH } from './fn-math'
import { STATS } from './fn-stats'
import { TEXT } from './fn-text'
import type { FnDef } from './fn'

export const FUNCTIONS: Readonly<Record<string, FnDef>> = {
  ...MATH,
  ...STATS,
  ...LOGIC,
  ...TEXT,
  ...DATE,
  ...LOOKUP,
  ...DIST,
  ...FINANCIAL,
  ...ENGINEERING,
  SINGLE: { kind: 'eager', min: 1, max: 1, scalarArgs: false, fn: ([v], ev) => ev.scalar(v) },
}

/** Names of all supported functions. */
export function supportedFunctions(): string[] {
  return Object.keys(FUNCTIONS).sort()
}
