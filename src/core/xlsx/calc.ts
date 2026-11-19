/**
 * Workbook recalculation: formula values are kept in the cell entries; after an edit
 * only the dependents (transitively) of the changed cells are re-evaluated.
 */
import { EvalSignal, evaluate, evaluateArray, FormulaError, isVolatile, parseFormula, references, UnsupportedFormula, type Ast, type CellProvider, type Value } from './formula'
import { isoToSerial } from './formula/dates'
import { expandShared } from './structure'
import { colAt, ownRow, MAX_COL, MAX_ROWS, T_BOOL, T_DATE, T_EMPTY, T_ERR, T_NUM, T_SST, T_STR, type SheetData, type XlsxWorkbook } from './workbook'

type Id = string // `${sheet}:${key}`
const idOf = (sheet: number, key: number) => `${sheet}:${key}`

/** a range some formulas read, and those formulas (many read the same one: a table looked up from every row) */
interface RangeDep {
  r0: number
  c0: number
  r1: number
  c1: number
  deps: Set<Id>
}

/**
 * Ranges by the rows they cover, to find those holding a cell without looking at them all (a sheet may have one per
 * formula): `fine` by blocks of 128 rows for the ranges of few blocks, `coarse` by blocks of 16 384 for the others.
 */
interface RangeIndex {
  fine: Map<number, RangeDep[]>
  coarse: Map<number, RangeDep[]>
}

/** the ranges read of one sheet: those a few columns wide under each of their columns, the wide ones apart */
interface SheetRanges {
  byKey: Map<string, RangeDep>
  cols: Map<number, RangeIndex>
  wide: RangeIndex
}

const FINE = 7
const COARSE = 14
/** blocks of rows a range is listed under, at most (a whole column: 64 of the coarse ones) */
const MAX_BLOCKS = 64
/** columns a range is listed under each of, at most */
const NARROW = 16

const newIndex = (): RangeIndex => ({ fine: new Map(), coarse: new Map() })

function listRange(ix: RangeIndex, g: RangeDep) {
  const add = (m: Map<number, RangeDep[]>, b: number) => {
    const list = m.get(b)
    if (list) list.push(g)
    else m.set(b, [g])
  }
  const last = Math.min(g.r1, MAX_ROWS - 1)
  if (last >> FINE <= (g.r0 >> FINE) + MAX_BLOCKS - 1) for (let b = g.r0 >> FINE; b <= last >> FINE; b++) add(ix.fine, b)
  else for (let b = Math.min(g.r0 >> COARSE, MAX_BLOCKS - 1); b <= Math.min(last >> COARSE, MAX_BLOCKS - 1); b++) add(ix.coarse, b)
}

/** what a formula reads, as registered (to take it back when the formula changes) */
interface Reads {
  cells: Id[]
  ranges: RangeDep[]
}

/**
 * Formulas computing one another, nested: past this many the innermost is computed on its own first (see `settle`).
 * A chain of a few hundred would otherwise overflow the stack: a column where every row adds to the one above.
 */
const MAX_DEPTH = 48

/** cells of an array formula followed one by one (dependents, order of computing): a larger array is taken as its first cell alone */
const MAX_FILL = 250_000

/** blocks wider than this are not kept (each of their columns is checked for what was stored since) */
const KEPT_COLS = 16
/** cells kept in all, at most (8 bytes each and more) */
const KEPT_CELLS = 4_000_000

/** the formula `id` is too deep in the stack to compute now */
class Deep extends EvalSignal {
  constructor(readonly id: Id) {
    super()
  }
}

export class Calc {
  private asts = new Map<Id, Ast | null>()
  private cellDeps = new Map<Id, Set<Id>>()
  /** by sheet read */
  private rangeDeps = new Map<number, SheetRanges>()
  private registered = new Map<Id, Reads>()
  /** ranges no formula reads any more, still listed (dropped when they are many) */
  private idle = 0
  /**
   * Formulas computed anew at every change: what they read is only known when they run (INDIRECT, OFFSET), or what
   * they give changes by itself (NOW, RAND). As in Excel.
   */
  private volatile = new Set<Id>()
  /** the cells an array formula fills (other than its own) → the formula's cell: reading one waits for the formula */
  private filledBy = new Map<Id, Id>()
  private names = new Map<string, number>()
  /** blocks of cells kept while a pass computes (`block`), those asked for once, and the cells they hold in all */
  private kept = new Map<string, { rows: Value[][]; at: number }>()
  private asked = new Set<string>()
  private keptCells = 0
  /** a clock of the values stored in this pass, and when one was last stored in a column (sheet * MAX_COL + column) */
  private clock = 0
  private stored = new Map<number, number>()
  private indexed = false
  /** whether the cells computed now are noted as changed (`SheetData.revalued`): not those computed when a file is read */
  private track = true
  private depth = 0

  constructor(private wb: XlsxWorkbook) {}

  /** the sheets changed (one loaded, added, moved, renamed, deleted) or formulas did under the index: it is built anew */
  invalidate() {
    this.indexed = false
    this.asts.clear()
    this.cellDeps.clear()
    this.rangeDeps.clear()
    this.registered.clear()
    this.volatile.clear()
    this.filledBy.clear()
    this.names.clear()
    this.idle = 0
  }

  /** the cells the array formula at (sheet, key) fills, when it is one */
  private arrayAt(sheet: number, key: number) {
    const d = this.wb.sheetsData.get(sheet)
    const rg = d?.arrays.get(key)
    return rg && d!.formulas.has(key) ? rg : undefined
  }

  private sheetIndex(name: string | undefined, self: number): number {
    if (!name) return self
    let i = this.names.get(name)
    if (i === undefined) {
      const lc = name.toLowerCase()
      i = this.wb.sheets.findIndex((s) => s.name.toLowerCase() === lc)
      this.names.set(name, i)
    }
    return i
  }

  private ast(sheet: number, key: number): Ast | null {
    const id = idOf(sheet, key)
    let a = this.asts.get(id)
    if (a === undefined) {
      const f = this.wb.sheetsData.get(sheet)?.formulas.get(key)
      try {
        a = f ? parseFormula(f) : null
      } catch {
        a = null
      }
      this.asts.set(id, a)
    }
    return a
  }

  private index() {
    if (this.indexed) return
    this.indexed = true
    for (const [sheet, d] of this.wb.sheetsData) {
      // the cells of a shared formula each compute: they are given their own text
      expandShared(d)
      for (const key of d.formulas.keys()) this.addDeps(sheet, key)
    }
  }

  private addDeps(sheet: number, key: number) {
    const a = this.ast(sheet, key)
    if (!a) return
    const dep = idOf(sheet, key)
    const reads: Reads = { cells: [], ranges: [] }
    for (const ref of references(a)) {
      const s = this.sheetIndex(ref.sheet, sheet)
      if (s < 0) continue
      if ('r' in ref) {
        const id = idOf(s, ref.r * MAX_COL + ref.c)
        let set = this.cellDeps.get(id)
        if (!set) this.cellDeps.set(id, (set = new Set()))
        set.add(dep)
        reads.cells.push(id)
      } else {
        let sr = this.rangeDeps.get(s)
        if (!sr) this.rangeDeps.set(s, (sr = { byKey: new Map(), cols: new Map(), wide: newIndex() }))
        const k = `${ref.r0}:${ref.c0}:${ref.r1}:${ref.c1}`
        let g = sr.byKey.get(k)
        if (!g) {
          // (a whole column reaches the rows past Excel's last, which a sheet read from a large text file has)
          g = { r0: ref.r0, c0: ref.c0, r1: ref.r1 >= MAX_ROWS - 1 ? Infinity : ref.r1, c1: ref.c1, deps: new Set() }
          sr.byKey.set(k, g)
          if (g.c1 - g.c0 >= NARROW) listRange(sr.wide, g)
          else {
            for (let c = g.c0; c <= g.c1; c++) {
              let ix = sr.cols.get(c)
              if (!ix) sr.cols.set(c, (ix = newIndex()))
              listRange(ix, g)
            }
          }
        } else if (!g.deps.size) this.idle--
        g.deps.add(dep)
        reads.ranges.push(g)
      }
    }
    if (reads.cells.length || reads.ranges.length) this.registered.set(dep, reads)
    if (isVolatile(a)) this.volatile.add(dep)
    const rg = this.arrayAt(sheet, key)
    // (an array over a sheet's worth of cells is not mapped: what reads its cells then reads them as they are)
    if (rg && (rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1) <= MAX_FILL) {
      for (let r = rg.r0; r <= rg.r1; r++) for (let c = rg.c0; c <= rg.c1; c++) if (r * MAX_COL + c !== key) this.filledBy.set(idOf(sheet, r * MAX_COL + c), dep)
    }
  }

  /** a formula was added / changed / removed at (sheet, key) */
  formulaChanged(sheet: number, key: number) {
    if (!this.indexed) return
    const id = idOf(sheet, key)
    this.asts.delete(id)
    this.volatile.delete(id)
    const was = this.registered.get(id)
    if (was) {
      this.registered.delete(id)
      for (const c of was.cells) this.cellDeps.get(c)?.delete(id)
      for (const g of was.ranges) if (g.deps.delete(id) && !g.deps.size) this.idle++
    }
    this.addDeps(sheet, key)
    if (this.idle > 8192) this.sweep()
  }

  /** drops the ranges no formula reads any more, once they are many (formulas changed over and over, each with its own range) */
  private sweep() {
    for (const [s, sr] of this.rangeDeps) {
      const next: SheetRanges = { byKey: new Map(), cols: new Map(), wide: newIndex() }
      for (const [k, g] of sr.byKey) {
        if (!g.deps.size) continue
        next.byKey.set(k, g)
        if (g.c1 - g.c0 >= NARROW) listRange(next.wide, g)
        else {
          for (let c = g.c0; c <= g.c1; c++) {
            let ix = next.cols.get(c)
            if (!ix) next.cols.set(c, (ix = newIndex()))
            listRange(ix, g)
          }
        }
      }
      this.rangeDeps.set(s, next)
    }
    this.idle = 0
  }

  /** the value of a cell entry (its type and what it holds) as the formula engine sees it */
  private entryValue(t: number, v: unknown): Value {
    switch (t) {
      case T_NUM:
        return v as number
      case T_SST:
        return this.wb.sst[v as number] ?? ''
      case T_STR:
        return String(v)
      case T_BOOL:
        return !!v
      case T_ERR:
        return new FormulaError(String(v) as ConstructorParameters<typeof FormulaError>[0])
      case T_EMPTY:
        return null
      case T_DATE:
        // an ISO date is the number of its day (and time), as for a date typed in
        return isoToSerial(String(v), this.wb.date1904) ?? String(v)
      default:
        return String(v)
    }
  }

  /** current value of a cell as the formula engine sees it */
  private value(sheet: number, r: number, c: number): Value {
    const d = this.wb.sheetsData.get(sheet)
    // a row of a page not in memory (a sheet read from its file by pages): the formula gets #VALUE!, not a blank
    if (d?.pages && !d.pages.has(r)) throw new Error('Row not in memory')
    const row = d?.rows[r]
    if (!row) return null
    const k = colAt(row, c)
    return k < row.length && row[k] === c ? this.entryValue(row[k + 1] as number, row[k + 2]) : null
  }

  /**
   * The cells of a block, as `provider.value` gives them one by one, without what that costs per cell (an id made
   * for each, the row searched from its start): a sum down a column of 10 000 rows, in each of 10 000 rows, is fifty
   * million of them. Only the cells holding a formula go the long way (they may be still to compute, or part of a
   * circular reference). Returns whether a cell was one being computed (read as #REF!, for now only).
   */
  private read(s: number, r0: number, c0: number, r1: number, c1: number, pending: Set<Id>, visiting: Set<Id>, cb: (v: Value, i: number, j: number) => void): boolean {
    const d = this.wb.sheetsData.get(s)
    const filled = this.filledBy.size > 0
    let circular = false
    for (let r = r0; r <= r1; r++) {
      if (d?.pages && !d.pages.has(r)) throw new Error('Row not in memory')
      let row = d?.rows[r]
      let k = row ? colAt(row, c0) : 0
      const base = r * MAX_COL
      for (let c = c0; c <= c1; c++) {
        let v: Value = null
        if (d && (d.formulas.has(base + c) || (filled && this.filledBy.has(idOf(s, base + c))))) {
          let id = idOf(s, base + c)
          if (filled) id = this.filledBy.get(id) ?? id
          if (visiting.has(id)) {
            v = new FormulaError('#REF!')
            circular = true
          } else {
            if (pending.has(id)) this.compute(id, pending, visiting)
            v = this.value(s, r, c)
          }
          // computing may have given the row cells
          row = d.rows[r]
          k = row ? colAt(row, c + 1) : 0
        } else if (row && k < row.length && row[k] === c) {
          v = this.entryValue(row[k + 1] as number, row[k + 2])
          k += 4
        }
        cb(v, r - r0, c - c0)
      }
    }
    return circular
  }

  /** see `CellProvider.block`: kept from the second time a block is asked for, while nothing is stored in its columns */
  private block(s: number, r0: number, c0: number, r1: number, c1: number, pending: Set<Id>, visiting: Set<Id>): Value[][] | undefined {
    if (c1 - c0 >= KEPT_COLS) return undefined
    const key = `${s}:${r0}:${c0}:${r1}:${c1}`
    const cells = (r1 - r0 + 1) * (c1 - c0 + 1)
    const hit = this.kept.get(key)
    if (hit) {
      let fresh = true
      for (let c = c0; c <= c1 && fresh; c++) fresh = (this.stored.get(s * MAX_COL + c) ?? 0) <= hit.at
      if (fresh) return hit.rows
      this.kept.delete(key)
      this.keptCells -= cells
    }
    // (a block asked for once is not kept: a running total has one of its own per row)
    if (!this.asked.has(key)) {
      this.asked.add(key)
      return undefined
    }
    if (this.keptCells + cells > KEPT_CELLS) return undefined
    const rows: Value[][] = []
    const last = c1 - c0
    let row: Value[] = []
    const circular = this.read(s, r0, c0, r1, c1, pending, visiting, (v, _i, j) => {
      row.push(v)
      if (j === last) {
        rows.push(row)
        row = []
      }
    })
    // (what a circular reference reads is #REF! while it is computed, not after)
    if (!circular) {
      this.kept.set(key, { rows, at: this.clock })
      this.keptCells += cells
    }
    return rows
  }

  /** a value was stored in column `c` of `sheet`: the blocks kept over it are stale */
  private touched(sheet: number, c: number) {
    this.stored.set(sheet * MAX_COL + c, ++this.clock)
  }

  private forget() {
    this.kept.clear()
    this.asked.clear()
    this.stored.clear()
    this.keptCells = 0
  }

  private provider(pending: Set<Id>, visiting: Set<Id>): CellProvider {
    return {
      value: (sheetName, r, c) => {
        const s = this.sheetIndex(sheetName, this.cur)
        if (s < 0) return new FormulaError('#REF!')
        let id = idOf(s, r * MAX_COL + c)
        // a cell an array formula fills has its value once that formula is computed
        if (this.filledBy.size) id = this.filledBy.get(id) ?? id
        // circular reference: an error, not the cell's previous value (which would make results depend on history)
        if (visiting.has(id)) return new FormulaError('#REF!')
        if (pending.has(id)) this.compute(id, pending, visiting)
        return this.value(s, r, c)
      },
      read: (sheetName, r0, c0, r1, c1, cb) => {
        const s = this.sheetIndex(sheetName, this.cur)
        if (s >= 0) return void this.read(s, r0, c0, r1, c1, pending, visiting, cb)
        const err = new FormulaError('#REF!')
        for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) cb(err, r - r0, c - c0)
      },
      block: (sheetName, r0, c0, r1, c1) => {
        const s = this.sheetIndex(sheetName, this.cur)
        return s < 0 ? undefined : this.block(s, r0, c0, r1, c1, pending, visiting)
      },
      bounds: (sheetName) => {
        const d = this.wb.sheetsData.get(this.sheetIndex(sheetName, this.cur))
        return { rows: (d?.maxRow ?? -1) + 1, cols: (d?.maxCol ?? -1) + 1 }
      },
      hasSheet: (name) => this.sheetIndex(name, -1) >= 0,
      date1904: this.wb.date1904,
    }
  }

  private cur = 0

  private compute(id: Id, pending: Set<Id>, visiting: Set<Id>) {
    const [sheet, key] = id.split(':').map(Number)
    const d = this.wb.sheetsData.get(sheet)
    if (!d || !d.formulas.has(key)) {
      pending.delete(id)
      return
    }
    if (this.depth >= MAX_DEPTH) throw new Deep(id)
    pending.delete(id)
    visiting.add(id)
    const a = this.ast(sheet, key)
    const prev = this.cur
    this.cur = sheet
    this.depth++
    let v: Value | undefined
    const rg = this.arrayAt(sheet, key)
    let filled: Value[][] | undefined
    try {
      // text the parser does not read (a structured reference, a reference to another workbook…) is not computed either
      if (!a) throw new UnsupportedFormula()
      const ctx = { provider: this.provider(pending, visiting), sheet: this.wb.sheets[sheet]?.name, r: Math.floor(key / MAX_COL), c: key % MAX_COL, strict: true }
      if (rg) {
        filled = evaluateArray(a, ctx, rg.r1 - rg.r0 + 1, rg.c1 - rg.c0 + 1)
        v = filled[0][0]
      } else v = evaluate(a, ctx)
    } catch (e) {
      if (e instanceof Deep) {
        // not computed: as it was, for the next try
        pending.add(id)
        throw e
      }
      v = e instanceof UnsupportedFormula ? undefined : new FormulaError('#VALUE!')
    } finally {
      this.depth--
      this.cur = prev
      visiting.delete(id)
    }
    // a function this engine does not have, a defined name: the value the file holds is the best there is. One the
    // engine computed in its place (#NAME?) would be saved over it, and read by what does not compute for itself.
    if (v === undefined) {
      if (cached(d, key)) return
      v = new FormulaError('#NAME?')
    }
    if (store(d, key, v) && this.track) d.revalued.add(key)
    this.touched(sheet, key % MAX_COL)
    // an array formula gives each cell it fills its value
    if (rg && filled) {
      for (let c = rg.c0; c <= rg.c1; c++) this.touched(sheet, c)
      for (let r = rg.r0; r <= rg.r1; r++) {
        for (let c = rg.c0; c <= rg.c1; c++) {
          const k = r * MAX_COL + c
          if (k !== key && !d.formulas.has(k) && store(d, k, filled[r - rg.r0][c - rg.c0]) && this.track) d.revalued.add(k)
        }
      }
    }
  }

  /**
   * Computes `start` and what it reads. Formulas nested too deep are computed from the innermost out: `stack` holds
   * the ones given up for now, each read (through others) by the one under it. So what the top one reads is never
   * one of those below, unless the reference is circular: they are `visiting` meanwhile, and read as such.
   */
  private settle(start: Id, pending: Set<Id>, visiting: Set<Id>) {
    const stack: Id[] = [start]
    while (stack.length) {
      const id = stack[stack.length - 1]
      visiting.delete(id)
      if (!pending.has(id)) {
        stack.pop()
        continue
      }
      try {
        this.depth = 0
        this.compute(id, pending, visiting)
        stack.pop()
      } catch (e) {
        if (!(e instanceof Deep)) throw e
        visiting.add(id)
        stack.push(e.id)
      }
    }
  }

  /**
   * The rows that computing the formulas a change touches would read: the formulas given (added: their text, not yet
   * in the sheet) and those depending on the changed cells. For sheets read from their file by pages, to have those
   * rows in memory first.
   */
  reads(changed: { sheet: number; key: number }[], added: { sheet: number; key: number; text: string }[]): { sheet: number; r0: number; r1: number }[] {
    this.index()
    const out: { sheet: number; r0: number; r1: number }[] = []
    const collect = (a: Ast, self: number) => {
      for (const ref of references(a)) {
        const s = this.sheetIndex(ref.sheet, self)
        if (s < 0) continue
        const max = this.wb.sheetsData.get(s)?.maxRow ?? -1
        if ('r' in ref) out.push({ sheet: s, r0: ref.r, r1: ref.r })
        else out.push({ sheet: s, r0: ref.r0, r1: ref.r1 >= MAX_ROWS - 1 ? max : Math.min(ref.r1, max) })
      }
    }
    for (const f of added) {
      try {
        collect(parseFormula(f.text), f.sheet)
      } catch {
        // not a formula after all: nothing read
      }
    }
    for (const id of this.affected(changed, [])) {
      const [s, key] = id.split(':').map(Number)
      const a = this.ast(s, key)
      if (a) collect(a, s)
    }
    return out
  }

  /** the formula cells to compute after `changed` cells change: the given ones, then every dependent (transitively) */
  private affected(changed: { sheet: number; key: number }[], formulas: { sheet: number; key: number }[], volatile = false): Id[] {
    const pending = new Set<Id>()
    const queue: Id[] = []
    const push = (id: Id) => {
      if (pending.has(id)) return
      pending.add(id)
      queue.push(id)
    }
    for (const f of formulas) push(idOf(f.sheet, f.key))
    const seen = new Set<Id>()
    // not recursive: a column where each row reads the one above is as deep as the sheet is long
    const todo: Id[] = []
    // the ranges whose readers are all in already (one read by every row is met once per row otherwise)
    const fired = new Set<RangeDep>()
    const hit = (list: RangeDep[] | undefined, r: number, c: number) => {
      if (!list) return
      for (const g of list) {
        if (r < g.r0 || r > g.r1 || c < g.c0 || c > g.c1 || fired.has(g)) continue
        fired.add(g)
        for (const dep of g.deps) {
          push(dep)
          todo.push(dep)
        }
      }
    }
    const visit = (first: Id) => {
      todo.push(first)
      while (todo.length) {
        const id = todo.pop()!
        if (seen.has(id)) continue
        seen.add(id)
        const at = id.indexOf(':')
        const s = Number(id.slice(0, at))
        const key = Number(id.slice(at + 1))
        const r = Math.floor(key / MAX_COL)
        const c = key % MAX_COL
        // what reads a cell an array formula fills depends on that formula
        const rg = pending.has(id) ? this.arrayAt(s, key) : undefined
        if (rg && (rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1) <= MAX_FILL) {
          for (let ar = rg.r0; ar <= rg.r1; ar++) for (let ac = rg.c0; ac <= rg.c1; ac++) if (ar !== r || ac !== c) todo.push(idOf(s, ar * MAX_COL + ac))
        }
        for (const dep of this.cellDeps.get(id) ?? []) {
          push(dep)
          todo.push(dep)
        }
        const sr = this.rangeDeps.get(s)
        if (sr) {
          const col = sr.cols.get(c)
          if (col) {
            hit(col.fine.get(r >> FINE), r, c)
            hit(col.coarse.get(Math.min(r >> COARSE, MAX_BLOCKS - 1)), r, c)
          }
          hit(sr.wide.fine.get(r >> FINE), r, c)
          hit(sr.wide.coarse.get(Math.min(r >> COARSE, MAX_BLOCKS - 1)), r, c)
        }
      }
    }
    for (const ch of changed) visit(idOf(ch.sheet, ch.key))
    for (const f of formulas) visit(idOf(f.sheet, f.key))
    // whatever changed, the volatile ones may give something else now, and what reads them with them
    if (changed.length || formulas.length || volatile) {
      for (const id of this.volatile) {
        push(id)
        visit(id)
      }
    }
    return queue
  }

  /**
   * Evaluates the given formula cells and everything that depends on the changed cells. `track`: the cells whose
   * value changes are noted for the next save (not when the values are those a file left out).
   */
  recalc(changed: { sheet: number; key: number }[], formulas: { sheet: number; key: number }[] = [], track = true, volatile = false): Set<number> {
    this.index()
    const queue = this.affected(changed, formulas, volatile)
    const pending = new Set<Id>(queue)
    const sheets = new Set<number>()
    const visiting = new Set<Id>()
    this.track = track
    // (what was kept of the last pass is of the cells as they were then)
    this.forget()
    try {
      for (const id of queue) {
        if (!pending.has(id)) continue
        this.settle(id, pending, visiting)
      }
    } finally {
      this.track = true
      this.forget()
    }
    for (const id of queue) sheets.add(Number(id.slice(0, id.indexOf(':'))))
    return sheets
  }

  /**
   * Formulas whose value is missing or not to be trusted (files written without cached values); `clock`: and those
   * whose value is of the moment (TODAY, NOW, RAND…), with what reads them. Returns the sheets computed in.
   */
  computeMissing(cells: { sheet: number; key: number }[], clock = false): Set<number> {
    if (!cells.length && !clock) return new Set()
    return this.recalc([], cells, false, clock)
  }
}

/** whether a formula cell holds a value (read from the file, or computed): not one just typed, shown as its text meanwhile */
function cached(d: SheetData, key: number): boolean {
  const row = d.rows[Math.floor(key / MAX_COL)]
  const c = key % MAX_COL
  if (!row) return false
  for (let k = 0; k < row.length; k += 4) {
    if (row[k] !== c) continue
    const t = row[k + 1]
    // (a formula just typed holds its own text, "=" before it, until it is computed: a text result may start with "=" too)
    return t !== T_EMPTY && !(t === T_STR && row[k + 2] === '=' + d.formulas.get(key))
  }
  return false
}

/** sets the value of a formula cell; false when it is the one it had */
function store(d: SheetData, key: number, v: Value): boolean {
  const r = Math.floor(key / MAX_COL)
  const c = key % MAX_COL
  // a formula cell was edited: its page is in memory for good
  if (d.pages && !d.pages.has(r)) throw new Error('Row not in memory')
  d.pages?.pin(r)
  let row = ownRow(d, r)
  if (!row) d.rows[r] = row = []
  let t: number
  let val: unknown
  if (v instanceof FormulaError) {
    t = T_ERR
    val = v.code
  } else if (typeof v === 'number') {
    t = T_NUM
    val = Number.isFinite(v) ? v : 0
  } else if (typeof v === 'boolean') {
    t = T_BOOL
    val = v ? 1 : 0
  } else if (v === null) {
    t = T_NUM
    val = 0
  } else {
    t = T_STR
    val = v
  }
  for (let k = 0; k < row.length; k += 4) {
    if (row[k] === c) {
      if (row[k + 1] === t && row[k + 2] === val) return false
      row[k + 1] = t
      row[k + 2] = val
      return true
    }
  }
  let k = 0
  while (k < row.length && (row[k] as number) < c) k += 4
  row.splice(k, 0, c, t, val, 0)
  return true
}
