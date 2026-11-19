/** Tree-walking evaluator. Top-level operands use implicit intersection; inside functions ranges are array-evaluated. */
import { getDate1904, setDate1904 } from './dates'
import { guard, pick, type Ev } from './fn'
import { FUNCTIONS } from './functions'
import { FormulaSyntaxError } from './lexer'
import { parseFormula, type Ast, type BinOp } from './parser'
import { sameSheet, type RefAddr } from './refs'
import { references } from './transform'
import { ERR, EvalSignal, FormulaError, MAX_COL, MAX_ROW, UnsupportedFormula, type CellProvider, type EvalContext, type Value } from './types'
import { ArrayV, compare, RangeV, toNum, toStr, UnionV, type Val } from './values'

/** the last + or - of a formula gives 0 for what is only the noise of the binary digits left (0.1+0.2-0.3), as in Excel */
const NOISE = 1e-15

function arith(op: BinOp, x: number, y: number, last = false): Value {
  let r: number
  switch (op) {
    case '+':
      r = x + y
      if (last && Math.abs(r) <= NOISE * Math.max(Math.abs(x), Math.abs(y))) r = 0
      break
    case '-':
      r = x - y
      if (last && Math.abs(r) <= NOISE * Math.max(Math.abs(x), Math.abs(y))) r = 0
      break
    case '*':
      r = x * y
      break
    case '/':
      if (y === 0) return ERR.DIV0
      r = x / y
      break
    default:
      if (x === 0 && y === 0) return ERR.NUM
      if (x === 0 && y < 0) return ERR.DIV0
      r = Math.pow(x, y)
  }
  return Number.isFinite(r) ? r : ERR.NUM
}

export function binop(op: BinOp, a: Value, b: Value, last = false): Value {
  if (a instanceof FormulaError) return a
  if (b instanceof FormulaError) return b
  switch (op) {
    case '&':
      return (toStr(a) as string) + (toStr(b) as string)
    case '=':
      return compare(a, b) === 0
    case '<>':
      return compare(a, b) !== 0
    case '<':
      return compare(a, b) < 0
    case '>':
      return compare(a, b) > 0
    case '<=':
      return compare(a, b) <= 0
    case '>=':
      return compare(a, b) >= 0
  }
  const x = toNum(a)
  if (x instanceof FormulaError) return x
  const y = toNum(b)
  if (y instanceof FormulaError) return y
  return arith(op, x, y, last)
}

type Operand = Value | Value[][]

function map1(v: Operand, f: (x: Value) => Value): Val {
  return Array.isArray(v) ? new ArrayV(v.map((row) => row.map(f))) : f(v)
}

function map2(a: Operand, b: Operand, f: (x: Value, y: Value) => Value): Val {
  if (!Array.isArray(a) && !Array.isArray(b)) return f(a, b)
  const da = Array.isArray(a) ? a : [[a]]
  const db = Array.isArray(b) ? b : [[b]]
  const h = Math.max(da.length, db.length)
  const w = Math.max(da[0].length, db[0].length)
  const out: Value[][] = []
  for (let i = 0; i < h; i++) {
    const row: Value[] = []
    for (let j = 0; j < w; j++) row.push(f(pick(da, i, j), pick(db, i, j)))
    out.push(row)
  }
  return new ArrayV(out)
}

export class Evaluator implements Ev {
  readonly provider: CellProvider
  readonly sheet: string | undefined
  readonly r: number
  readonly c: number
  private readonly strict: boolean
  /** the formula being evaluated: its last operation is the one that rounds noise to zero */
  private root: Ast | undefined
  /** the larger extent of the sheets the whole columns / rows of the formula are on (see `clip`) */
  private extent: { rows: number; cols: number } | undefined

  constructor(ctx: EvalContext, root?: Ast) {
    this.root = root
    this.provider = ctx.provider
    this.sheet = ctx.sheet
    this.r = ctx.r
    this.c = ctx.c
    this.strict = !!ctx.strict
  }

  /** what the engine cannot compute: #NAME?, or the evaluation stopped (`EvalContext.strict`) */
  private unsupported(): Val {
    if (this.strict) throw new UnsupportedFormula()
    return ERR.NAME
  }

  eval(ast: Ast, top: boolean): Val {
    switch (ast.type) {
      case 'num':
      case 'str':
      case 'bool':
        return ast.value
      case 'err':
        return new FormulaError(ast.code)
      case 'missing':
        return null
      case 'name':
        return this.unsupported()
      case 'array':
        return new ArrayV(ast.rows)
      case 'ref':
        return this.ref(ast.ref)
      case 'range':
        return this.rangeOp(ast.left, ast.right)
      case 'union': {
        const areas: RangeV[] = []
        for (const it of ast.items) {
          const v = this.eval(it, false)
          if (v instanceof RangeV) areas.push(v)
          else if (v instanceof UnionV) areas.push(...v.areas)
          else return v instanceof FormulaError ? v : ERR.VALUE
        }
        return new UnionV(areas)
      }
      case 'unary': {
        const v = this.operand(this.eval(ast.arg, top), top)
        if (ast.op === '+') return Array.isArray(v) ? new ArrayV(v) : v
        return map1(v, (x) => {
          const n = toNum(x)
          return n instanceof FormulaError ? n : n === 0 ? 0 : -n
        })
      }
      case 'percent':
        return map1(this.operand(this.eval(ast.arg, top), top), (x) => {
          const n = toNum(x)
          return n instanceof FormulaError ? n : n / 100
        })
      case 'binary': {
        const l = this.operand(this.eval(ast.left, top), top)
        const r = this.operand(this.eval(ast.right, top), top)
        const op = ast.op
        const last = ast === this.root
        return map2(l, r, (a, b) => binop(op, a, b, last))
      }
      case 'call':
        return this.call(ast.name, ast.args, top)
    }
  }

  private ref(a: RefAddr): Val {
    if (a.sheet !== undefined && !this.hasSheet(a.sheet)) return ERR.REF
    // past Excel's last row only a sheet read from a large text file has such rows: no row there, no cell (a name, as in Excel)
    if (a.r1 > MAX_ROW && a.r1 >= (this.provider.bounds?.(a.sheet ?? this.sheet).rows ?? 0)) return ERR.NAME
    return new RangeV(a.sheet ?? this.sheet, a.r0, a.c0, a.r1, a.c1)
  }

  private rangeOp(left: Ast, right: Ast): Val {
    const l = this.eval(left, false)
    if (l instanceof FormulaError) return l
    const r = this.eval(right, false)
    if (r instanceof FormulaError) return r
    if (!(l instanceof RangeV) || !(r instanceof RangeV) || !sameSheet(l.sheet, r.sheet)) return ERR.VALUE
    return new RangeV(l.sheet, Math.min(l.r0, r.r0), Math.min(l.c0, r.c0), Math.max(l.r1, r.r1), Math.max(l.c1, r.c1))
  }

  private call(name: string, args: Ast[], top: boolean): Val {
    if (!Object.hasOwn(FUNCTIONS, name)) return this.unsupported()
    const def = FUNCTIONS[name]
    if (args.length < def.min || args.length > def.max) return ERR.VALUE
    if (def.kind === 'lazy') return guard(() => def.fn(args, this, top))
    const vals = args.map((a) => this.eval(a, def.scalarArgs && top))
    return guard(() => def.fn(vals, this, top))
  }

  hasSheet(name: string): boolean {
    return this.provider.hasSheet ? this.provider.hasSheet(name) : true
  }

  now(): Date {
    return this.provider.now?.() ?? new Date()
  }

  cell(sheet: string | undefined, r: number, c: number): Value {
    const v = this.provider.value(sheet, r, c)
    return v === undefined ? null : v
  }

  /**
   * The rows and columns of the longest sheet among those the formula has whole columns / rows of: `A:A` and `B:B` of
   * two sheets must have the same size (SUMPRODUCT(Data!C:C,T!A:A)), as they do in Excel, without reading the blank
   * rows of a million.
   */
  private wholeExtent(): { rows: number; cols: number } {
    if (this.extent) return this.extent
    let rows = 0
    let cols = 0
    if (this.root && this.provider.bounds) {
      const seen = new Set<string>()
      for (const ref of references(this.root)) {
        if (!('r0' in ref) || (ref.r1 < MAX_ROW && ref.c1 < MAX_COL)) continue
        const key = (ref.sheet ?? this.sheet ?? '').toLowerCase()
        if (seen.has(key)) continue
        seen.add(key)
        const b = this.provider.bounds(ref.sheet ?? this.sheet)
        rows = Math.max(rows, b.rows)
        cols = Math.max(cols, b.cols)
      }
    }
    return (this.extent = { rows, cols })
  }

  /** Clip whole-row/whole-column style ranges to the provider's used extent. */
  private clip(rg: RangeV): { r1: number; c1: number } {
    let { r1, c1 } = rg
    if ((r1 >= MAX_ROW || c1 === MAX_COL) && this.provider.bounds) {
      const b = this.provider.bounds(rg.sheet)
      const ext = this.wholeExtent()
      // a whole column is all the rows the sheet has: more than Excel's when it was read from a large text file
      if (r1 >= MAX_ROW) r1 = Math.max(rg.r0, Math.max(b.rows, ext.rows) - 1)
      if (c1 === MAX_COL) c1 = Math.max(rg.c0, Math.min(c1, Math.max(b.cols, ext.cols) - 1))
    }
    return { r1, c1 }
  }

  forEach(rg: RangeV, cb: (v: Value, i: number, j: number) => void): void {
    const { r1, c1 } = this.clip(rg)
    const p = this.provider
    const kept = p.block?.(rg.sheet, rg.r0, rg.c0, r1, c1)
    if (kept) {
      for (let i = 0; i < kept.length; i++) for (let j = 0, row = kept[i]; j < row.length; j++) cb(row[j], i, j)
      return
    }
    if (p.read) return p.read(rg.sheet, rg.r0, rg.c0, r1, c1, cb)
    for (let r = rg.r0; r <= r1; r++) for (let c = rg.c0; c <= c1; c++) cb(this.cell(rg.sheet, r, c), r - rg.r0, c - rg.c0)
  }

  /** The values of a range / array / scalar as rows. Those of a range may be shared (see `CellProvider.block`): read, never changed. */
  toArray(v: Val): Value[][] {
    if (v instanceof RangeV) {
      const { r1, c1 } = this.clip(v)
      const p = this.provider
      const kept = p.block?.(v.sheet, v.r0, v.c0, r1, c1)
      if (kept) return kept
      const out: Value[][] = []
      if (p.read) {
        const last = c1 - v.c0
        let row: Value[] = []
        p.read(v.sheet, v.r0, v.c0, r1, c1, (x, _i, j) => {
          row.push(x)
          if (j === last) {
            out.push(row)
            row = []
          }
        })
        return out
      }
      for (let r = v.r0; r <= r1; r++) {
        const row: Value[] = []
        for (let c = v.c0; c <= c1; c++) row.push(this.cell(v.sheet, r, c))
        out.push(row)
      }
      return out
    }
    if (v instanceof ArrayV) return v.data
    if (v instanceof UnionV) throw ERR.VALUE
    return [[v]]
  }

  /** Implicit intersection of a range with the formula cell's row/column. */
  intersect(rg: RangeV): Value {
    if (rg.h === 1 && rg.w === 1) return this.cell(rg.sheet, rg.r0, rg.c0)
    if (rg.w === 1 && this.r >= rg.r0 && this.r <= rg.r1) return this.cell(rg.sheet, this.r, rg.c0)
    if (rg.h === 1 && this.c >= rg.c0 && this.c <= rg.c1) return this.cell(rg.sheet, rg.r0, this.c)
    return ERR.VALUE
  }

  scalar(v: Val): Value {
    if (v instanceof RangeV) return this.intersect(v)
    if (v instanceof ArrayV) return v.data[0][0]
    if (v instanceof UnionV) return ERR.VALUE
    return v
  }

  operand(v: Val, top: boolean): Operand {
    if (v instanceof RangeV) {
      if (v.h === 1 && v.w === 1) return this.cell(v.sheet, v.r0, v.c0)
      return top ? this.intersect(v) : this.toArray(v)
    }
    if (v instanceof ArrayV) return v.h === 1 && v.w === 1 ? v.data[0][0] : v.data
    if (v instanceof UnionV) return ERR.VALUE
    return v
  }
}

/**
 * Evaluate a parsed formula for the cell at ctx.{sheet,r,c}. Errors come back as FormulaError; it throws nothing but
 * an `EvalSignal` (one the provider throws, or `UnsupportedFormula` when `ctx.strict`).
 */
export function evaluate(ast: Ast, ctx: EvalContext): Value {
  const prev = getDate1904()
  setDate1904(!!ctx.provider.date1904)
  try {
    const ev = new Evaluator(ctx, ast)
    const v = ev.eval(ast, true)
    const x = v instanceof RangeV ? ev.intersect(v) : v instanceof ArrayV ? v.data[0][0] : v instanceof UnionV ? ERR.VALUE : v
    if (x === null) return 0
    if (typeof x === 'number') return Number.isFinite(x) ? x + 0 : ERR.NUM
    return x
  } catch (e) {
    if (e instanceof EvalSignal) throw e
    return e instanceof FormulaError ? e : ERR.VALUE
  } finally {
    setDate1904(prev)
  }
}

/**
 * Evaluate an array formula (entered over `h` × `w` cells): no implicit intersection at the top, and a value for
 * each cell. A result of one row / column / value repeats over the cells; cells past a larger one are #N/A.
 */
export function evaluateArray(ast: Ast, ctx: EvalContext, h: number, w: number): Value[][] {
  const prev = getDate1904()
  setDate1904(!!ctx.provider.date1904)
  const fill = (f: (i: number, j: number) => Value) => Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => f(i, j)))
  try {
    const ev = new Evaluator(ctx, ast)
    const v = ev.eval(ast, false)
    const data = v instanceof UnionV ? [[ERR.VALUE as Value]] : ev.toArray(v)
    return fill((i, j) => {
      const x = pick(data, i, j)
      if (x === null) return 0
      if (typeof x === 'number') return Number.isFinite(x) ? x + 0 : ERR.NUM
      return x
    })
  } catch (e) {
    if (e instanceof EvalSignal) throw e
    const err = e instanceof FormulaError ? e : ERR.VALUE
    return fill(() => err)
  } finally {
    setDate1904(prev)
  }
}

/** Parse + evaluate. Syntax errors evaluate to #NAME?. */
export function evaluateFormula(text: string, ctx: EvalContext): Value {
  let ast: Ast
  try {
    ast = parseFormula(text)
  } catch (e) {
    if (e instanceof FormulaSyntaxError) return ERR.NAME
    return ERR.VALUE
  }
  return evaluate(ast, ctx)
}
