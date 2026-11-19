/**
 * Second round of the QA pass of 18/11/2026 (findings rated medium and low; ids as in its report): recalculation,
 * what is typed, sorting, formats of whole columns, collaboration.
 */
import { describe, expect, test } from 'bun:test'
import { at, book, openBook, rows, type Opened } from './xlsx-builder'

const recalculated = async (o: Opened) => {
  while (!o.events.some((e) => e.event === 'xlsx:progress' && e.data.recalc)) await new Promise((r) => setTimeout(r, 5))
}
const num = async (o: Opened, sheet: number, ref: string) => Number(((await o.engine.call('xlsx.cell', [sheet, at(ref).r, at(ref).c])) as any)?.value)

describe('P1-013: a sheet where every row reads the whole table', () => {
  const N = 6000
  test('running totals, COUNTIF and SUMIFS down a column follow one edit, with the right values', async () => {
    const grid: any[][] = []
    for (let r = 1; r <= N; r++) {
      grid.push([r % 50, r % 7, r, { f: `SUM($C$1:C${r})`, v: 0 }, { f: `COUNTIF(A:A,A${r})`, v: 0 }, { f: `SUMIFS(C:C,A:A,A${r},B:B,B${r})`, v: 0 }, { f: `VLOOKUP(${(r % 40) + 1},$A$1:$C$60,3,FALSE)`, v: 0 }])
    }
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows(grid) }] }))
    await o.set(0, 'Z1', '1')
    const t0 = performance.now()
    await o.set(0, 'C1', '1000')
    await o.set(0, 'A2', '49')
    const ms = performance.now() - t0
    // by hand
    const A = grid.map((g) => g[0] as number)
    const B = grid.map((g) => g[1] as number)
    const C = grid.map((g) => g[2] as number)
    C[0] = 1000
    A[1] = 49
    const sumifs = (r: number) => C.reduce((s, c, i) => (A[i] === A[r] && B[i] === B[r] ? s + c : s), 0)
    for (const r of [0, 1, 2, 48, 49, 349, N - 1]) {
      expect(await num(o, 0, `D${r + 1}`)).toBe(C.slice(0, r + 1).reduce((x, y) => x + y, 0))
      expect(await num(o, 0, `E${r + 1}`)).toBe(A.filter((a) => a === A[r]).length)
      expect(await num(o, 0, `F${r + 1}`)).toBe(sumifs(r))
      expect(await num(o, 0, `G${r + 1}`)).toBe(C[A.slice(0, 60).indexOf(((r + 1) % 40) + 1)])
    }
    // the cells were read twenty million times per column: seconds, not minutes (the bound is for a slow machine)
    expect(ms).toBeLessThan(60_000)
  }, 120_000)

  test('a block kept while computing is read again once a formula in it has its value', async () => {
    // B reads A whole, twice over (kept from the second time); A2 is a formula still to compute when B1 reads it
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1, { f: 'SUM($A$1:$A$3)', v: 0 }], [{ f: 'A1*10', v: 0 }, { f: 'SUM($A$1:$A$3)', v: 0 }], [3, { f: 'SUM($A$1:$A$3)+C3', v: 0 }, { f: 'A2+1', v: 0 }]]) }] }))
    await o.set(0, 'A1', '2')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 1)]).toEqual(['25', '25', '46'])
    await o.set(0, 'A3', '=B1')
    // a circular reference through a kept block: an error, the same however often it is computed
    const first = [await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 0)]
    await o.set(0, 'A1', '2')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 0)]).toEqual(first)
    expect(first).toContain('#REF!')
    await o.set(0, 'A3', '4')
    expect([await o.shown(0, 0, 1), await o.shown(0, 1, 1), await o.shown(0, 2, 1)]).toEqual(['26', '26', '47'])
  })
})

describe('P1-025: TODAY, NOW and RAND are of the moment the file is read', () => {
  test('what the file holds for them, and for what reads them, is computed again; the file is not changed by it', async () => {
    const src = await book({
      sheets: [
        { name: 'Due', rows: rows([[{ f: 'TODAY()', v: 45000 }, 44990, { f: 'A1-B1', v: 10 }, { f: 'NOW()', v: 45000.5 }, { f: 'RAND()', v: 0.123 }, { f: 'B1*2', v: 1 }]]) },
        { name: 'Other', rows: rows([[{ f: 'Due!C1+1', v: 11 }]]) },
      ],
    })
    const o = await openBook(src)
    await recalculated(o)
    const today = await num(o, 0, 'A1')
    expect(today).toBeGreaterThan(46000)
    expect(Number.isInteger(today)).toBe(true)
    expect(await num(o, 0, 'C1')).toBe(today - 44990)
    expect(await num(o, 0, 'D1')).toBeGreaterThanOrEqual(today)
    expect(await num(o, 0, 'E1')).not.toBe(0.123)
    // what does not read the clock keeps the value of the file, right or not: nothing asked to compute it
    expect(await num(o, 0, 'F1')).toBe(1)
    await o.load(1)
    expect(await num(o, 1, 'A1')).toBe(today - 44990 + 1)
    expect(await o.engine.call('dirty', [])).toBe(false)
  })
})

describe('what is typed (P1-027, N1-007, N1-022)', () => {
  const STYLES = (fmts: string[]) =>
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${fmts.length}">${fmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${f}"/>`).join('')}</numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${fmts.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${fmts.map((_, i) => `<xf numFmtId="${164 + i}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('')}</cellXfs></styleSheet>`
  const cell = async (o: Opened, ref: string) => (await o.engine.call('xlsx.cell', [0, at(ref).r, at(ref).c])) as any

  test('"+A1+B1", "-A1" and "@SUM(…)" are formulas; "-5", "- item" and a phone number are not', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'D', rows: rows([[1, 2, 3]]) }] }))
    await o.set(0, 'A3', '+A1+B1')
    await o.set(0, 'B3', '-C1')
    await o.set(0, 'C3', '@SUM(A1:C1)')
    await o.set(0, 'D3', '+D!A1+D!B1')
    expect([await o.formula(0, 2, 0), await o.shown(0, 2, 0)]).toEqual(['=+A1+B1', '3'])
    expect([await o.formula(0, 2, 1), await o.shown(0, 2, 1)]).toEqual(['=-C1', '-3'])
    expect([await o.formula(0, 2, 2), await o.shown(0, 2, 2)]).toEqual(['=SUM(A1:C1)', '6'])
    expect(await o.shown(0, 2, 3)).toBe('3')
    await o.set(0, 'A5', '-5')
    await o.set(0, 'B5', '- item')
    await o.set(0, 'C5', '+84 912 345 678')
    await o.set(0, 'D5', '@username')
    await o.set(0, 'E5', '-abc')
    expect((await cell(o, 'A5')).type).toBe('number')
    for (const ref of ['B5', 'C5', 'D5', 'E5']) expect([(await cell(o, ref)).type, (await cell(o, ref)).formula]).toEqual(['string', undefined])
  })

  test('an error typed is an error; a number longer than a number holds stays as typed; amounts are numbers shown as typed', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))
    await o.set(0, 'A2', '#N/A')
    await o.set(0, 'B2', '=ISNA(A2)')
    await o.set(0, 'C2', '=IFERROR(A2,"none")')
    expect([(await cell(o, 'A2')).type, await o.shown(0, 1, 1), await o.shown(0, 1, 2)]).toEqual(['error', 'TRUE', 'none'])
    await o.set(0, 'A3', '9704123456789012345')
    expect([(await cell(o, 'A3')).type, await o.shown(0, 2, 0)]).toEqual(['string', '9704123456789012345'])
    await o.set(0, 'B3', '123456789012345')
    expect((await cell(o, 'B3')).type).toBe('number')
    const typed: [string, string, number, string][] = [
      ['A4', '1,234,567', 1234567, '1,234,567'],
      ['B4', '$50', 50, '$50'],
      ['C4', '€1,200.5', 1200.5, '€1,200.50'],
      ['D4', '50 ₫', 50, '50 ₫'],
      ['E4', '-$3.75', -3.75, '-$3.75'],
      ['F4', '1,234.5', 1234.5, '1,234.5'],
    ]
    for (const [ref, input, n, shown] of typed) {
      await o.set(0, ref, input)
      expect([ref, Number((await cell(o, ref)).value), await o.shown(0, at(ref).r, at(ref).c)]).toEqual([ref, n, shown])
    }
    // not amounts: an address, a decimal comma, two signs
    for (const [ref, input] of [['A6', '192.168.100.200'], ['B6', '1,5'], ['C6', '$5€'], ['D6', '12,34,567']] as const) {
      await o.set(0, ref, input)
      expect([ref, (await cell(o, ref)).type]).toEqual([ref, 'string'])
    }
    const re = await openBook((await o.save()).blob)
    expect([await re.shown(0, 3, 1), await re.shown(0, 3, 2), await re.shown(0, 1, 0)]).toEqual(['$50', '€1,200.50', '#N/A'])
  })

  test('a formula typed into a cell without a format shows as what it computes', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[{ s: 1, v: 45292 }, { s: 2, v: 0.25 }, 5, { s: 1, v: 45300 }]]) }], styles: STYLES(['dd/mm/yyyy', '0.0%']) }))
    expect(await o.shown(0, 0, 0)).toBe('01/01/2024')
    await o.set(0, 'A2', '=A1')
    await o.set(0, 'B2', '=A1+30')
    await o.set(0, 'C2', '=DATE(2024,2,29)')
    await o.set(0, 'D2', '=EDATE(A1,1)')
    await o.set(0, 'E2', '=B1')
    await o.set(0, 'F2', '=D1-A1')
    await o.set(0, 'G2', '=C1+1')
    await o.set(0, 'H2', '=MAX(A1:D1)')
    expect([await o.shown(0, 1, 0), await o.shown(0, 1, 1), await o.shown(0, 1, 2), await o.shown(0, 1, 3)]).toEqual(['01/01/2024', '31/01/2024', '29/02/2024', '01/02/2024'])
    expect([await o.shown(0, 1, 4), await o.shown(0, 1, 5), await o.shown(0, 1, 6), await o.shown(0, 1, 7)]).toEqual(['25.0%', '8', '6', '09/01/2024'])
    await o.set(0, 'A3', '=TODAY()')
    expect(await o.shown(0, 2, 0)).toMatch(/^\d{2}\/\d{2}\/20\d{2}$/)
    // a cell that has a format keeps it
    await o.set(0, 'B1', '=A1')
    expect(await o.shown(0, 0, 1)).toBe('4529200.0%')
    // undone with its format
    await o.set(0, 'J2', '=A1')
    await o.engine.call('xlsx.undo', [])
    await o.set(0, 'J2', '7')
    expect(await o.shown(0, 1, 9)).toBe('7')
  })
})

describe('sorting (P2-021, P2-022, N1-020, P2-036)', () => {
  const col = async (o: Opened, c: number, n: number) => Promise.all(Array.from({ length: n }, (_, r) => o.shown(0, r, c)))

  test('Excel\'s order: numbers, text by its letters, FALSE and TRUE, errors, then the cells without anything', async () => {
    const values: any[] = ['apple', 10, 'item2', true, 0, null, 'b', -1, 'Apple', 'item10', false, '10', 9, '9']
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows(values.map((v, i) => [v, i])) }] }))
    await o.set(0, 'A5', '=NA()')
    const range = { r0: 0, c0: 0, r1: values.length - 1, c1: 1 }
    await o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range, col: 0 }])
    expect(await col(o, 0, values.length)).toEqual(['-1', '9', '10', '10', '9', 'apple', 'Apple', 'b', 'item10', 'item2', 'FALSE', 'TRUE', '#N/A', ''])
    // the cells beside go with their rows
    expect((await col(o, 1, values.length)).slice(0, 3)).toEqual(['7', '12', '1'])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range, col: 0, desc: true }])
    expect(await col(o, 0, values.length)).toEqual(['#N/A', 'TRUE', 'FALSE', 'item2', 'item10', 'b', 'apple', 'Apple', '9', '10', '10', '9', '-1', ''])
    await o.engine.call('xlsx.undo', [])
    await o.engine.call('xlsx.undo', [])
    expect(await col(o, 0, 4)).toEqual(['apple', '10', 'item2', 'TRUE'])
    expect(await o.formula(0, 4, 0)).toBe('=NA()')
  })

  test('formulas reading their row go with it, in the file too; a range with merged cells is refused; a fill skips the cells under a merge', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[3, { f: 'A1*2', v: 6 }, 'c'], [1, { f: 'A2*2', v: 2 }, 'a'], [2, { f: 'A3*2', v: 4 }, 'b']]), tail: '<mergeCells count="1"><mergeCell ref="E2:E3"/></mergeCells>' }] }))
    await o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range: { r0: 0, c0: 0, r1: 2, c1: 2 }, col: 0 }])
    expect([await o.formula(0, 0, 1), await o.shown(0, 0, 1), await o.shown(0, 0, 2), await o.formula(0, 2, 1), await o.shown(0, 2, 1)]).toEqual(['=A1*2', '2', 'a', '=A3*2', '6'])
    const saved = await o.save()
    const xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c>')
    const re = await openBook(saved.blob)
    expect([await re.shown(0, 2, 0), await re.shown(0, 2, 1), await re.shown(0, 2, 2)]).toEqual(['3', '6', 'c'])
    await expect(o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range: { r0: 0, c0: 0, r1: 2, c1: 4 }, col: 0 }])).rejects.toThrow(/merged cells/)
    // E2:E3 is one cell: a fill down D1:E3 writes D2, D3 and E2, not the E3 nobody sees
    await o.set(0, 'D1', 'x')
    await o.set(0, 'E1', 'y')
    await o.engine.call('xlsx.sheetOp', [0, { op: 'fill', range: { r0: 0, c0: 3, r1: 2, c1: 4 }, dir: 'down' }])
    expect([await o.shown(0, 2, 3), await o.shown(0, 1, 4), await o.shown(0, 2, 4)]).toEqual(['x', 'y', ''])
  })

  test('a comment goes with its row, and back when the sort is undone', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[3, 'c'], [1, 'a'], [2, 'b']]) }] }))
    const sheet = o.info.xlsx.sheets[0].path
    await o.engine.call('comments.set', ['c1', { id: 'c1', author: { name: 'Ann' }, at: { sheet, r: 0, c: 1 }, text: 'on c', time: '2026-11-19T00:00:00Z' }])
    const where = async () => ((await o.engine.call('comments.list', [])) as any[]).find((c) => c.id === 'c1').at.r
    await o.engine.call('xlsx.sheetOp', [0, { op: 'sort', range: { r0: 0, c0: 0, r1: 2, c1: 1 }, col: 0 }])
    expect([await o.shown(0, 2, 1), await where()]).toEqual(['c', 2])
    await o.engine.call('xlsx.undo', [])
    expect(await where()).toBe(0)
    await o.engine.call('xlsx.redo', [])
    expect(await where()).toBe(2)
  })
})

describe('Undo and Redo go through every state again (snapshots share the rows they were taken of)', () => {
  const dump = async (o: Opened) => {
    const b: any = await o.engine.call('xlsx.block', [0, 0, 60])
    const cells = b.rows.map((r: any) => [r.r, r.cells.filter((c: any) => c[1] !== '' || c[2]).map((c: any) => [c[0], c[1], c[2]])]).filter((r: any) => r[1].length)
    const formulas: string[] = []
    for (let r = 0; r < 40; r++) for (let c = 0; c < 8; c++) {
      const f = await o.formula(0, r, c)
      if (f) formulas.push(`${r},${c}${f}`)
    }
    // (the sheet as the worker holds it: asking through an operation would be a step of the history under test)
    const d = (o.engine as any).xlsx.sheetsData.get(0)
    return JSON.stringify([cells, formulas, [...d.merges].sort((a: any, b: any) => a.r0 - b.r0 || a.c0 - b.c0)])
  }
  for (const start of [3, 11, 29, 47]) test(`seed ${start}`, async () => {
    let seed = start
    const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296)
    const int = (n: number) => Math.floor(rand() * n)
    const grid: any[][] = []
    for (let r = 0; r < 14; r++) grid.push([r % 5, `t${r}`, { f: `A${r + 1}*2`, v: (r % 5) * 2 }, r < 4 ? { f: `SUM(A1:A${r + 2})`, v: [1, 3, 6, 10][r] } : null, 14 - r])
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows(grid) }, { name: 'T', rows: rows([[{ f: 'S!A3+S!C5', v: 10 }]]) }] }))
    await o.load(1)
    const states = [await dump(o)]
    const steps = 26
    let done = 0
    for (let i = 0; i < steps; i++) {
      const r = int(12)
      const c = int(5)
      const x = rand()
      const call = (m: string, ...a: unknown[]) => o.engine.call(m as never, a)
      try {
        let res: any
        if (x < 0.25) res = await call('xlsx.set', 0, [{ r, c, input: rand() < 0.4 ? `=A${r + 1}+B${int(10) + 1}` : rand() < 0.5 ? String(i) : `v${i}` }])
        else if (x < 0.36) res = await call('xlsx.format', 0, { r0: r, c0: c, r1: r + 1, c1: c + 1 }, { b: rand() < 0.5 })
        else if (x < 0.46) res = await call('xlsx.sheetOp', 0, { op: 'insertRows', at: r, n: 1 + int(2) })
        else if (x < 0.54) res = await call('xlsx.sheetOp', 0, { op: 'deleteRows', at: r, n: 1 })
        else if (x < 0.6) res = await call('xlsx.sheetOp', 0, { op: 'insertCols', at: c, n: 1 })
        else if (x < 0.66) res = await call('xlsx.sheetOp', 0, { op: 'deleteCols', at: c, n: 1 })
        else if (x < 0.8) res = await call('xlsx.sheetOp', 0, { op: 'sort', range: { r0: r, c0: 0, r1: r + 5, c1: 4 }, col: c, desc: rand() < 0.5 })
        else if (x < 0.87) res = await call('xlsx.sheetOp', 0, { op: 'clear', range: { r0: r, c0: c, r1: r + 1, c1: c + 1 } })
        else if (x < 0.94) res = await call('xlsx.sheetOp', 0, { op: 'fill', range: { r0: r, c0: c, r1: r + 2, c1: c }, dir: 'down' })
        else res = await call('xlsx.sheetOp', 0, { op: 'merge', range: { r0: 20 + int(4), c0: c, r1: 25, c1: c + 1 } })
        // (a sort that moves nothing, a clear of nothing: no step)
        const now = await dump(o)
        if (res && now !== states[states.length - 1]) {
          states.push(now)
          done++
        }
      } catch {
        // refused (a sort over merged cells…): nothing changed
        expect(await dump(o)).toBe(states[states.length - 1])
      }
    }
    expect(done).toBeGreaterThan(10)
    const final = states[states.length - 1]
    const seen = new Set(states)
    // back, state by state (a step that changed nothing shown is one Undo all the same: every state met is one that was)
    for (let k = 0; k < 60; k++) {
      const res = await o.engine.call('xlsx.undo', [])
      if (!res) break
      expect(seen.has(await dump(o))).toBe(true)
    }
    expect(await dump(o)).toBe(states[0])
    for (let k = 0; k < 60; k++) {
      const res = await o.engine.call('xlsx.redo', [])
      if (!res) break
      expect(seen.has(await dump(o))).toBe(true)
    }
    expect(await dump(o)).toBe(final)
    // and the file written from there opens as that state
    const re = await openBook((await o.save()).blob)
    await re.load(1)
    expect(await dump(re)).toBe(final)
  }, 60_000)
})

describe('P2-042: a format given to whole columns, rows or the sheet', () => {
  const STYLES = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'
  const big = () => book({ sheets: [{ name: 'S', rows: rows(Array.from({ length: 3000 }, (_, r) => [r, `t${r}`, r * 2])) }], styles: STYLES })
  const bold = (o: Opened, st: number) => o.info.xlsx.styles[st]?.fontWeight === 700 || ((o.engine as any).xlsx.styles.xfs[st]?.css.fontWeight === 700)
  const styleOf = async (o: Opened, r: number, c: number) => ((await o.engine.call('xlsx.block', [0, r, r + 1])) as any).rows[0]?.cells.find((x: any) => x[0] === c)?.[2] ?? -1

  test('a whole column: every cell of it, the column itself for the rest, a few bytes in the file', async () => {
    const o = await openBook(await big())
    const res: any = await o.engine.call('xlsx.format', [0, { r0: 0, c0: 1, r1: 1_048_575, c1: 1 }, { b: true }])
    expect(res.layout.colInfo.find((ci: any) => ci.min === 1 && ci.max === 1)?.x?.style).toBeDefined()
    for (const r of [0, 1500, 2999]) expect(bold(o, await styleOf(o, r, 1))).toBe(true)
    expect(bold(o, await styleOf(o, 0, 0))).toBe(false)
    // no cells made down to the last row
    expect((o.engine as any).xlsx.sheetsData.get(0).maxRow).toBe(2999)
    // what is typed below takes the column's format
    await o.set(0, 'B5000', 'new')
    expect(bold(o, await styleOf(o, 4999, 1))).toBe(true)
    await o.set(0, 'C5000', 'plain')
    expect(bold(o, await styleOf(o, 4999, 2))).toBe(false)
    const saved = await o.save()
    const xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).toMatch(/<cols><col min="2" max="2" [^>]*style="\d+"\/><\/cols>/)
    expect(xml).toMatch(/<c r="B3000" s="\d+" t="inlineStr">/)
    // (the rows the sheet had, and the one typed into: none made to carry the format)
    expect((xml.match(/<row /g) ?? []).length).toBe(3001)
    const re = await openBook(saved.blob)
    expect(bold(re, await styleOf(re, 2999, 1))).toBe(true)
    await re.set(0, 'B6000', 'x')
    expect(bold(re, await styleOf(re, 5999, 1))).toBe(true)
    // undone in one step: the cells and the column
    await o.engine.call('xlsx.undo', [])
    await o.engine.call('xlsx.undo', [])
    await o.engine.call('xlsx.undo', [])
    expect(bold(o, await styleOf(o, 1500, 1))).toBe(false)
    expect((await (await o.save()).text('xl/worksheets/sheet1.xml'))).not.toContain('<cols>')
  })

  test('the whole sheet, then a row over it, then some columns: each place shows the format given last', async () => {
    const o = await openBook(await big())
    const all = { r0: 0, c0: 0, r1: 1_048_575, c1: 16_383 }
    const t0 = performance.now()
    await o.engine.call('xlsx.format', [0, all, { b: true }])
    expect(performance.now() - t0).toBeLessThan(20_000)
    const d = (o.engine as any).xlsx.sheetsData.get(0)
    expect(d.maxCol).toBe(2)
    expect(bold(o, await styleOf(o, 2999, 2))).toBe(true)
    await o.set(0, 'XFD9', 'far')
    expect(bold(o, await styleOf(o, 8, 16_383))).toBe(true)
    // a row: its own format, over the sheet's
    await o.engine.call('xlsx.format', [0, { r0: 4, c0: 0, r1: 4, c1: 16_383 }, { b: false, i: true }])
    const row5: any = ((await o.engine.call('xlsx.block', [0, 4, 5])) as any).rows[0]
    expect(row5.s).toBeGreaterThan(0)
    const css = (st: number) => (o.engine as any).xlsx.styles.xfs[st].css
    expect([css(row5.s).fontWeight, css(row5.s).fontStyle]).toEqual([undefined, 'italic'])
    await o.set(0, 'K5', 'in the row')
    expect(css(await styleOf(o, 4, 10)).fontStyle).toBe('italic')
    const saved = await o.save()
    const xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).toMatch(/<col min="1" max="16384" [^>]*style="\d+"\/>/)
    expect(xml).toMatch(/<row r="5"[^>]* s="\d+" customFormat="1"/)
    const re = await openBook(saved.blob)
    expect(css(await styleOf(re, 4, 0)).fontStyle).toBe('italic')
    expect(((await re.engine.call('xlsx.block', [0, 4, 5])) as any).rows[0].s).toBeGreaterThan(0)
  })

  test('a range too large to format cell by cell says so; the cells of a large one are kept by number, not one by one', async () => {
    const o = await openBook(await big())
    await expect(o.engine.call('xlsx.format', [0, { r0: 0, c0: 0, r1: 900_000, c1: 9 }, { b: true }])).rejects.toThrow(/whole columns or rows/)
    await o.engine.call('xlsx.format', [0, { r0: 0, c0: 0, r1: 2999, c1: 2 }, { i: true }])
    const ch = (o.engine as any).xlsxEd.undoStack.at(-1)
    expect(ch.cells.length).toBe(0)
    expect(ch.restyle.keys.length).toBe(9000)
  })
})
