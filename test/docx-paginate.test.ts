import { describe, expect, test } from 'bun:test'
import type { Block, PageSetup, ParaStyle } from '../src/core/docx/types'
import { DEFAULT_PAGE } from '../src/core/docx/types'
import { columnsOf, pageAt, pagesOf, paginate, pieceOf, type Lines, type PageSlice } from '../src/react/docx/paginate'

const page: PageSetup = { ...DEFAULT_PAGE }

interface Spec {
  /** paragraph: line heights; table: row heights */
  p?: number[]
  t?: number[]
  /** space before / after a paragraph */
  before?: number
  after?: number
  head?: number
  keep?: number[]
  pb?: boolean
  ps?: ParaStyle
  sect?: string
  /** a block measured as a whole only (no lines read) */
  h?: number
}

/** blocks with their measures: lines / rows of the given heights */
function doc(specs: Spec[]) {
  const blocks: Block[] = []
  const heights = new WeakMap<Block, number>()
  const lines = new WeakMap<Block, Lines>()
  for (const s of specs) {
    if (s.sect) {
      const b: Block = { t: 'sect', page, kind: s.sect }
      blocks.push(b)
      heights.set(b, 0)
      continue
    }
    const b: Block = s.t ? { t: 'tbl', cols: [100], rows: s.t.map(() => ({ cells: [] })) } : { t: 'p', ps: s.ps ?? {}, runs: [], pb: s.pb }
    blocks.push(b)
    if (s.h !== undefined) {
      heights.set(b, s.h)
      continue
    }
    const units = s.p ?? s.t!
    const ys = [s.before ?? 0]
    for (const u of units) ys.push(ys[ys.length - 1] + u)
    heights.set(b, ys[ys.length - 1] + (s.after ?? 0))
    lines.set(b, { ys, head: s.head, keep: s.keep })
  }
  return { blocks, heights, lines }
}

const run = (specs: Spec[], avail = 100, cols = 1, lastKind?: string) => {
  const d = doc(specs)
  return paginate(d.blocks, d.heights, d.blocks.length, avail, cols, lastKind, d.lines)
}

/** each page as its pieces: `b` (whole block b) or `b:from-to` (lines / rows of block b) */
function pieces(pages: PageSlice[]): string[][] {
  return pages.map((p) =>
    columnsOf(p).flatMap((c) => {
      const out: string[] = []
      for (let bi = c.start; bi < c.end; bi++) {
        const { from, to } = pieceOf(c, bi)
        out.push(from === 0 && to === undefined ? String(bi) : `${bi}:${from}-${to ?? 'end'}`)
      }
      return out
    }),
  )
}

const lines = (n: number, h = 10) => Array<number>(n).fill(h)

describe('paragraphs split between lines', () => {
  test('the lines that fit stay, the rest goes on on the next page', () => {
    // 6 lines of 10 fit under 40: 4 stay
    const pages = run([{ p: [60] }, { p: lines(6) }])
    expect(pieces(pages)).toEqual([['0', '1:0-4'], ['1:4-end']])
    expect(pages.map((p) => p.used)).toEqual([100, 20])
  })

  test('a paragraph longer than a page is cut on every page', () => {
    expect(pieces(run([{ p: lines(25) }]))).toEqual([['0:0-10'], ['0:10-20'], ['0:20-end']])
  })

  test('the space before counts on its first page only; the space after needs no room', () => {
    // before 15: 15 + 8 lines = 95, the 9th line does not fit
    expect(pieces(run([{ p: lines(12), before: 15 }]))).toEqual([['0:0-8'], ['0:8-end']])
    // the lines fit, the space after does not: the paragraph stays whole, the next one starts the next page
    const pages = run([{ p: lines(9), after: 20 }, { p: lines(2) }])
    expect(pieces(pages)).toEqual([['0'], ['1']])
    expect(pages[0].used).toBe(100)
  })

  test('a paragraph measured as a whole (lines not read yet) moves whole', () => {
    expect(pieces(run([{ p: [70] }, { h: 50 }]))).toEqual([['0'], ['1']])
  })

  test('a line taller than the page: its page grows, nothing is lost', () => {
    const pages = run([{ p: [30] }, { p: [10, 150, 10] }])
    expect(pieces(pages)).toEqual([['0', '1:0-1'], ['1:1-2'], ['1:2-end']])
    expect(pages.map((p) => p.used)).toEqual([40, 150, 10])
  })
})

describe('tables split between rows', () => {
  test('rows that fit stay; the header rows repeat on each following page', () => {
    // 25 rows of 10, one header row: 10 rows on page 1, then the header + 9 rows per page
    const pages = run([{ t: lines(25), head: 1 }])
    expect(pieces(pages)).toEqual([['0:0-10'], ['0:10-19'], ['0:19-end']])
    expect(pages.map((p) => p.used)).toEqual([100, 100, 70])
  })

  test('a table after a paragraph starts on its page when its header and a row fit there', () => {
    expect(pieces(run([{ p: lines(8) }, { t: lines(5), head: 1 }]))).toEqual([['0', '1:0-2'], ['1:2-end']])
    // only the header would fit: the table starts on the next page
    expect(pieces(run([{ p: lines(9) }, { t: lines(5), head: 1 }]))).toEqual([['0'], ['1']])
  })

  test('a header taller than the page is not repeated', () => {
    expect(pieces(run([{ t: [120, 10, 10], head: 1 }]))).toEqual([['0:0-1'], ['0:1-end']])
  })

  test('a row is never split: one taller than the page makes its page grow', () => {
    const pages = run([{ t: [10, 130, 10] }])
    expect(pieces(pages)).toEqual([['0:0-1'], ['0:1-2'], ['0:2-end']])
    expect(pages.map((p) => p.used)).toEqual([10, 130, 10])
    // with a header row: it comes first on the grown page
    const h = run([{ t: [10, 20, 130, 10], head: 1 }])
    expect(pieces(h)).toEqual([['0:0-2'], ['0:2-3'], ['0:3-end']])
    expect(h.map((p) => p.used)).toEqual([30, 140, 20])
  })

  test('boundaries a merged cell spans are kept together', () => {
    // rows 3-6 hold a merged cell too tall to be cut at 4, 5 or 6: 5 rows would fit, 3 stay
    expect(pieces(run([{ t: lines(12), keep: [4, 5, 6] }], 55))).toEqual([['0:0-3'], ['0:3-8'], ['0:8-end']])
  })
})

describe('breaks, sections, columns', () => {
  test('a forced page break and a section break still start a page, after split blocks too', () => {
    expect(pieces(run([{ p: lines(15) }, { p: lines(2), pb: true }, { p: lines(2) }]))).toEqual([['0:0-10'], ['0:10-end'], ['1', '2']])
    expect(pieces(run([{ p: lines(15) }, { sect: 'nextPage' }, { p: lines(2) }], 100, 1, 'nextPage'))).toEqual([['0:0-10'], ['0:10-end', '1'], ['2']])
    expect(pieces(run([{ p: lines(15) }, { sect: 'nextPage' }, { p: lines(2) }], 100, 1, 'continuous'))).toEqual([['0:0-10'], ['0:10-end', '1', '2']])
  })

  test('text columns are filled one after the other, blocks split between them', () => {
    const pages = run([{ p: lines(8) }, { p: lines(6) }, { t: lines(30), head: 1 }], 100, 2)
    // column 1: block 0 and 2 lines of block 1; column 2: the rest of block 1 and the table up to row 6; then the
    // header and 9 rows in each column
    expect(pieces(pages)).toEqual([
      ['0', '1:0-2', '1:2-end', '2:0-6'],
      ['2:6-15', '2:15-24'],
      ['2:24-end'],
    ])
    expect(columnsOf(pages[0]).map((c) => [c.start, c.from, c.end, c.to])).toEqual([
      [0, 0, 2, 2],
      [1, 2, 3, 6],
    ])
  })

  test('the pages of a block, the page of an offset', () => {
    const d = doc([{ p: [50] }, { p: lines(24) }, { p: [10] }])
    const pages = paginate(d.blocks, d.heights, d.blocks.length, 100, 1, undefined, d.lines)
    expect(pieces(pages)).toEqual([['0', '1:0-5'], ['1:5-15'], ['1:15-end', '2']])
    expect(pagesOf(pages, 0)).toEqual([0, 0])
    expect(pagesOf(pages, 1)).toEqual([0, 2])
    expect(pagesOf(pages, 2)).toEqual([2, 2])
    expect(pagesOf(pages, 3)).toEqual([-1, -1])
    const l = d.lines.get(d.blocks[1])!
    // line starts not read yet: the first page
    expect(pageAt(pages, 1, 400, l)).toBe(0)
    l.at = []
    l.at[5] = 100
    l.at[15] = 300
    expect([0, 99, 100, 299, 300, 500].map((o) => pageAt(pages, 1, o, l))).toEqual([0, 0, 1, 1, 2, 2])
  })
})
