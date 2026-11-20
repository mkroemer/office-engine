/**
 * What the measure pass reads of a rendered block beyond its height: where its rows (a table) or lines (a paragraph)
 * are, so that pagination can cut it between two of them.
 */
import type { Table } from '../../core/docx/types'
import type { Lines } from './paginate'

/**
 * The rows of a top-level table (`el`: its rendered block): their tops and the bottom of the last, the leading header
 * rows, and the row boundaries a merged cell spans when what it holds would not fit above them.
 */
export function tableLines(el: HTMLElement, t: Table): Lines | undefined {
  const trs = Array.from(el.querySelector('table')?.tBodies[0]?.rows ?? [])
  if (!trs.length || trs.length !== t.rows.length) return undefined
  const top = el.getBoundingClientRect().top
  const ys = trs.map((tr) => tr.getBoundingClientRect().top - top)
  ys.push(trs[trs.length - 1].getBoundingClientRect().bottom - top)
  let head = 0
  while (head < t.rows.length && t.rows[head].header) head++
  const keep = new Set<number>()
  trs.forEach((tr, r) => {
    for (const td of Array.from(tr.cells)) {
      if (td.rowSpan <= 1) continue
      // a merged cell going on on the next page keeps what it holds above the break
      const kids = td.children
      const cs = getComputedStyle(td)
      let need = parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom)
      if (kids.length) need += kids[kids.length - 1].getBoundingClientRect().bottom - kids[0].getBoundingClientRect().top
      for (let b = r + 1; b < Math.min(r + td.rowSpan, trs.length); b++) if (ys[b] - ys[r] < need - 0.5) keep.add(b)
    }
  })
  return { ys, head: head < t.rows.length ? head : 0, keep: keep.size ? [...keep] : undefined }
}
