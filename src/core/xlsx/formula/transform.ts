/** Formula text rewriting (fill/copy, row/col insert/delete, sheet rename) and dependency extraction. */
import { tokenize, type Token } from './lexer'
import type { Ast } from './parser'
import { fmtRef, normalizeRef, quoteSheet, refInBounds, sameSheet, type RefAddr } from './refs'
import { MAX_COL, MAX_REF_ROW, MAX_ROW, type CellRef, type RangeRef } from './types'

/** Re-emit `text`, replacing tokens for which `map` returns a string. Unlexable text is returned unchanged. */
function rewrite(text: string, map: (t: Token) => string | undefined): string {
  const lead = text.startsWith('=') ? '=' : ''
  const src = text.slice(lead.length)
  let toks: Token[]
  try {
    toks = tokenize(src)
  } catch {
    return text
  }
  let out = lead
  let pos = 0
  for (const t of toks) {
    const rep = map(t)
    if (rep === undefined) continue
    out += src.slice(pos, t.s) + rep
    pos = t.e
  }
  return out + src.slice(pos)
}

const refError = (t: Token & { t: 'ref' }) => (t.ref.sheet === undefined ? '' : `${t.sheetText ?? quoteSheet(t.ref.sheet)}!`) + '#REF!'

function translateRef(t: Token & { t: 'ref' }, dr: number, dc: number): string {
  const a: RefAddr = { ...t.ref }
  if (a.kind !== 'cols') {
    if (!a.ar0) a.r0 += dr
    if (!a.ar1) a.r1 += dr
  }
  if (a.kind !== 'rows') {
    if (!a.ac0) a.c0 += dc
    if (!a.ac1) a.c1 += dc
  }
  // (a reference past Excel's last row, to a row of a large text file, may move on past it)
  if (!refInBounds(a, t.ref.r1 > MAX_ROW ? MAX_REF_ROW : MAX_ROW)) return refError(t)
  return fmtRef(normalizeRef(a), t.sheetText)
}

/** Copy/fill: shift relative parts of every reference by (dr, dc). Out-of-sheet results become #REF!. */
export function translateFormula(text: string, dr: number, dc: number): string {
  if (!dr && !dc) return text
  return rewrite(text, (t) => (t.t === 'ref' ? translateRef(t, dr, dc) : undefined))
}

/**
 * `translateFormula` of one text by many offsets (the cells of a shared formula): the text is read once.
 * Unlexable text is returned unchanged, as there.
 */
export function translator(text: string): (dr: number, dc: number) => string {
  const lead = text.startsWith('=') ? '=' : ''
  const src = text.slice(lead.length)
  let refs: (Token & { t: 'ref' })[]
  try {
    refs = tokenize(src).filter((t): t is Token & { t: 'ref' } => t.t === 'ref')
  } catch {
    return () => text
  }
  // the text between the references, kept as it is
  const between: string[] = []
  let pos = 0
  for (const t of refs) {
    between.push(src.slice(pos, t.s))
    pos = t.e
  }
  const tail = src.slice(pos)
  return (dr, dc) => {
    if (!dr && !dc) return text
    let out = lead
    for (let i = 0; i < refs.length; i++) out += between[i] + translateRef(refs[i], dr, dc)
    return out + tail
  }
}

/** New [lo, hi] after inserting (delta>0) / deleting (delta<0) at `at`; null when entirely deleted. */
function shiftSpan(lo: number, hi: number, at: number, delta: number, max: number): [number, number] | null {
  if (delta > 0) {
    if (lo >= at) lo += delta
    if (hi >= at) hi += delta
    if (lo > max) return null
    return [lo, Math.min(hi, max)]
  }
  const n = -delta
  const end = at + n
  if (hi < at) return [lo, hi]
  if (lo >= end) return [lo - n, hi - n]
  if (lo >= at && hi < end) return null
  return [lo < at ? lo : at, hi >= end ? hi - n : at - 1]
}

export interface ShiftOp {
  /** sheet whose rows/cols change (undefined: the formula's own sheet) */
  sheet?: string
  /** sheet the formula lives on (unqualified refs belong to it; undefined: same as `sheet`) */
  formulaSheet?: string
  axis: 'row' | 'col'
  /** 0-based index of the first inserted/deleted row/col */
  at: number
  /** >0 insert count, <0 delete count */
  delta: number
}

/** Rewrite references after inserting/deleting rows or columns. */
export function shiftFormula(text: string, op: ShiftOp): string {
  if (!op.delta) return text
  const affects = (refSheet: string | undefined): boolean => {
    if (op.sheet === undefined) return refSheet === undefined || (op.formulaSheet !== undefined && sameSheet(refSheet, op.formulaSheet))
    const s = refSheet ?? op.formulaSheet
    return s === undefined || sameSheet(s, op.sheet)
  }
  return rewrite(text, (t) => {
    if (t.t !== 'ref' || !affects(t.ref.sheet)) return undefined
    const a: RefAddr = { ...t.ref }
    if (op.axis === 'row') {
      if (a.kind === 'cols') return undefined
      const s = shiftSpan(a.r0, a.r1, op.at, op.delta, a.r1 > MAX_ROW ? MAX_REF_ROW : MAX_ROW)
      if (!s) return refError(t)
      ;[a.r0, a.r1] = s
    } else {
      if (a.kind === 'rows') return undefined
      const s = shiftSpan(a.c0, a.c1, op.at, op.delta, MAX_COL)
      if (!s) return refError(t)
      ;[a.c0, a.c1] = s
    }
    if (a.r0 === t.ref.r0 && a.r1 === t.ref.r1 && a.c0 === t.ref.c0 && a.c1 === t.ref.c1) return undefined
    return fmtRef(a, t.sheetText)
  })
}

/** cells moved from one place to another (cut and pasted): the block, as it was, and how far it went */
export interface MoveOp {
  from: string
  to: string
  r0: number
  c0: number
  r1: number
  c1: number
  dr: number
  dc: number
}

/**
 * Rewrite references after cells were moved: a reference to cells of the block (all of them inside it) follows them,
 * the others stay. `home`: the sheet the formula is on now; `was`: the sheet it was on, for a formula that moved with
 * the block to another sheet (what it read of its old sheet, it still reads there).
 */
export function moveRefsInFormula(text: string, op: MoveOp, home: string, was = home): string {
  return rewrite(text, (t) => {
    if (t.t !== 'ref') return undefined
    const a = t.ref
    const sheet = a.sheet ?? was
    const inside = sameSheet(sheet, op.from) && (a.kind === 'cell' || a.kind === 'area') && a.r0 >= op.r0 && a.r1 <= op.r1 && a.c0 >= op.c0 && a.c1 <= op.c1
    const target = inside ? op.to : sheet
    const local = sameSheet(target, home)
    // as written: nothing of it moved, and it names its sheet as it has to
    if (!inside && (a.sheet === undefined) === local) return undefined
    const b: RefAddr = inside ? { ...a, r0: a.r0 + op.dr, r1: a.r1 + op.dr, c0: a.c0 + op.dc, c1: a.c1 + op.dc } : { ...a }
    return fmtRef({ ...b, sheet: local ? undefined : target }, !local && a.sheet !== undefined && sameSheet(a.sheet, target) ? t.sheetText : undefined)
  })
}

/** References to a sheet that is no more (deleted): each becomes #REF!, as the names and errors qualified by it. */
export function dropSheetInFormula(text: string, name: string): string {
  return rewrite(text, (t) => {
    if ((t.t === 'ref' && t.ref.sheet !== undefined && sameSheet(t.ref.sheet, name)) || ((t.t === 'err' || t.t === 'name') && t.sheet !== undefined && sameSheet(t.sheet, name))) return '#REF!'
    return undefined
  })
}

/** Rename sheet qualifiers (`Old!A1`, `'Old'!A1:B2`, `Old!#REF!`, `Old!Name`). */
export function renameSheetInFormula(text: string, oldName: string, newName: string): string {
  const pre = `${quoteSheet(newName)}!`
  return rewrite(text, (t) => {
    if (t.t === 'ref' && t.ref.sheet !== undefined && sameSheet(t.ref.sheet, oldName))
      return fmtRef({ ...t.ref, sheet: newName })
    if ((t.t === 'err' || t.t === 'name') && t.sheet !== undefined && sameSheet(t.sheet, oldName)) return pre + t.v
    return undefined
  })
}

/**
 * Functions newer than the file format (Excel 2010 and later): a file names them with a prefix, and Excel takes one
 * written without for a name it does not know (#NAME?).
 */
const FUTURE = new Set(
  (
    'ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW ' +
    'CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH ' +
    'COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FILTERXML FLOOR.MATH ' +
    'FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE ' +
    'GAUSS HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV ' +
    'MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC ' +
    'PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE RRI SCAN SEC SECH SEQUENCE ' +
    'SHEET SHEETS SINGLE SKEW.P SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW ' +
    'UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST'
  ).split(' '),
)
/** these are of the worksheet namespace as well */
const FUTURE_WS = new Set(['FILTER', 'SORT'])

/** The formula as a file holds it: the functions newer than the format given their prefix (one typed has none). */
export function withFunctionPrefixes(text: string): string {
  if (!/[A-Za-z]\(/.test(text)) return text
  return rewrite(text, (t) => {
    if (t.t !== 'func') return undefined
    const src = text.slice(text.startsWith('=') ? t.s + 1 : t.s, text.startsWith('=') ? t.e + 1 : t.e)
    if (src.startsWith('_')) return undefined
    if (FUTURE.has(t.v)) return `_xlfn.${src}`
    if (FUTURE_WS.has(t.v)) return `_xlfn._xlws.${src}`
    return undefined
  })
}

/** The formula as it is typed: without the prefixes a file gives the newer functions. */
export function stripFunctionPrefixes(text: string): string {
  if (!text.includes('_xl') && !text.includes('_XL')) return text
  return rewrite(text, (t) => {
    if (t.t !== 'func') return undefined
    const src = text.slice(text.startsWith('=') ? t.s + 1 : t.s, text.startsWith('=') ? t.e + 1 : t.e)
    const bare = src.replace(/^(?:_xlfn\.|_xlws\.|_xludf\.)+/i, '')
    return bare === src ? undefined : bare
  })
}

/** functions computed at every change: what they read is known when they run, or what they give changes by itself */
const VOLATILE = new Set(['INDIRECT', 'OFFSET', 'NOW', 'TODAY', 'RAND', 'RANDBETWEEN'])

/** whether the formula calls a volatile function */
export function isVolatile(ast: Ast): boolean {
  switch (ast.type) {
    case 'call':
      return VOLATILE.has(ast.name) || ast.args.some(isVolatile)
    case 'unary':
    case 'percent':
      return isVolatile(ast.arg)
    case 'binary':
    case 'range':
      return isVolatile(ast.left) || isVolatile(ast.right)
    case 'union':
      return ast.items.some(isVolatile)
    default:
      return false
  }
}

function toRef(a: RefAddr): CellRef | RangeRef {
  const sheet = a.sheet === undefined ? {} : { sheet: a.sheet }
  if (a.kind === 'cell') return { ...sheet, r: a.r0, c: a.c0, absR: a.ar0, absC: a.ac0 }
  return { ...sheet, r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }
}

/** Every cell/range the formula reads directly (static references; INDIRECT/OFFSET targets are dynamic). */
export function references(ast: Ast): (CellRef | RangeRef)[] {
  const out: (CellRef | RangeRef)[] = []
  const walk = (n: Ast): void => {
    switch (n.type) {
      case 'ref':
        out.push(toRef(n.ref))
        break
      case 'range':
        if (n.left.type === 'ref' && n.right.type === 'ref' && sameSheet(n.left.ref.sheet, n.right.ref.sheet)) {
          const l = n.left.ref
          const r = n.right.ref
          out.push(
            toRef({ ...l, kind: 'area', r0: Math.min(l.r0, r.r0), c0: Math.min(l.c0, r.c0), r1: Math.max(l.r1, r.r1), c1: Math.max(l.c1, r.c1) }),
          )
        } else {
          walk(n.left)
          walk(n.right)
        }
        break
      case 'unary':
      case 'percent':
        walk(n.arg)
        break
      case 'binary':
        walk(n.left)
        walk(n.right)
        break
      case 'union':
        n.items.forEach(walk)
        break
      case 'call':
        n.args.forEach(walk)
        break
    }
  }
  walk(ast)
  return out
}
