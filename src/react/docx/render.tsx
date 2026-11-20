import { memo, type CSSProperties, type PointerEvent as RPointerEvent, type ReactNode } from 'react'
import { cssColor } from '../../core/color'
import type { Block, Inline, Paragraph, ParaStyle, RunStyle, Table, TableCell, TrackMark } from '../../core/docx/types'
import { safeHref } from '../../core/url'
import type { DocumentClient } from '../../core/worker/client'
import { useMediaUrl } from '../common/hooks'
import { dict } from '../../core/dict'
import { gridOf, type GridCell } from './tables'

export interface RenderCtx {
  client: DocumentClient
  /** 0-based page index, -1 while measuring */
  page: number
  pages: number
  measuring: boolean
  editable: boolean
  /** editing: a column edge of the top-level table `bi` was dragged */
  colResize?: (bi: number, cols: number[]) => void
}

const pt = (v: number | undefined) => (v ? `${+v.toFixed(2)}pt` : undefined)

const UNDERLINE_STYLE: Record<string, CSSProperties['textDecorationStyle']> = dict({
  double: 'double', dotted: 'dotted', dottedHeavy: 'dotted', dash: 'dashed', dashedHeavy: 'dashed', dashLong: 'dashed',
  wave: 'wavy', wavyHeavy: 'wavy', wavyDouble: 'wavy',
})

const TRACK_COLORS = ['#c0392b', '#1a73e8', '#188038', '#a142f4', '#e37400', '#00838f', '#d01884', '#5f6368']

const authorSlot = new Map<string, number>()

/** one colour per author of tracked changes, given in the order authors are met (as Word does): the first 8 differ */
export function authorColor(author: string | undefined): string {
  const k = author ?? ''
  let i = authorSlot.get(k)
  if (i === undefined) authorSlot.set(k, (i = authorSlot.size))
  return TRACK_COLORS[i % TRACK_COLORS.length]
}

/** inserted text underlined, deleted text struck through, in the author's colour */
function trackCss(tc: TrackMark): CSSProperties {
  const color = authorColor(tc.author)
  return tc.k === 'ins'
    ? { color, textDecoration: 'underline', textDecorationColor: color }
    : { color, textDecoration: 'line-through', textDecorationColor: color }
}

export function runCss(s: RunStyle | undefined): CSSProperties {
  const css: CSSProperties = {}
  if (!s) return css
  if (s.b) css.fontWeight = 700
  else if (s.b === false) css.fontWeight = 400
  if (s.i) css.fontStyle = 'italic'
  else if (s.i === false) css.fontStyle = 'normal'
  const u = s.u && s.u !== 'none'
  const deco = [u ? 'underline' : '', s.strike || s.dstrike ? 'line-through' : ''].filter(Boolean).join(' ')
  if (deco) css.textDecoration = deco
  else if (s.u === 'none' || s.strike === false) css.textDecoration = 'none'
  if (u && UNDERLINE_STYLE[s.u!]) css.textDecorationStyle = UNDERLINE_STYLE[s.u!]
  if (s.dstrike) css.textDecorationStyle = 'double'
  // colors of the model come from the file, a peer or the clipboard: `background` would load a url() in one
  if (s.color) css.color = cssColor(s.color)
  if (s.sz) css.fontSize = `${s.va ? s.sz * 0.65 : s.sz}pt`
  if (s.font) css.fontFamily = s.font
  if (s.hl || s.shd) css.background = cssColor(s.hl ?? s.shd)
  if (s.va) css.verticalAlign = s.va === 'sup' ? 'super' : 'sub'
  if (s.caps) css.textTransform = 'uppercase'
  if (s.smallCaps) css.fontVariant = 'small-caps'
  if (s.spacing) css.letterSpacing = `${s.spacing}pt`
  return css
}

function lineHeight(ps: ParaStyle): string | number | undefined {
  if (!ps.line) return 1.17
  if (ps.lineRule === 'exact') return `${ps.line}pt`
  if (ps.lineRule === 'atLeast') return `max(${ps.line}pt, 1.17em)`
  return +(ps.line * 1.17).toFixed(3)
}

/** contextual spacing flags: 'b' drop space before, 'a' drop space after */
export type Spacing = '' | 'b' | 'a' | 'ba'

export function contextual(p: Block | undefined, prev: Block | undefined, next: Block | undefined): Spacing {
  if (!p || p.t !== 'p' || !p.ps.contextual) return ''
  const b = prev?.t === 'p' && prev.sid === p.sid ? 'b' : ''
  const a = next?.t === 'p' && next.sid === p.sid ? 'a' : ''
  return (b + a) as Spacing
}

function paraCss(ps: ParaStyle, sp: Spacing): CSSProperties {
  const css: CSSProperties = {
    paddingTop: sp.includes('b') ? undefined : pt(ps.before),
    paddingBottom: sp.includes('a') ? undefined : pt(ps.after),
    paddingLeft: pt(ps.indL),
    paddingRight: pt(ps.indR),
    textIndent: pt(ps.indFirst),
    textAlign: ps.align,
    lineHeight: lineHeight(ps),
  }
  if (ps.shd) css.background = cssColor(ps.shd)
  const b = (x: ParaStyle['bTop']) => (x ? `${x.w}pt ${x.style} ${cssColor(x.color) ?? 'currentColor'}` : undefined)
  if (ps.bTop) css.borderTop = b(ps.bTop)
  if (ps.bBottom) css.borderBottom = b(ps.bBottom)
  if (ps.bLeft) css.borderLeft = b(ps.bLeft)
  if (ps.bRight) css.borderRight = b(ps.bRight)
  return css
}

function Img({ ctx, target, w, h, alt }: { ctx: RenderCtx; target: string; w: number; h: number; alt?: string }) {
  const url = useMediaUrl(ctx.client, target, !ctx.measuring)
  const size = { width: `${w}pt`, height: `${h}pt` }
  if (ctx.measuring || !url) return <span className="de-img-missing" style={{ ...size, borderStyle: 'none' }} />
  if (url === 'unsupported') {
    return (
      <span className="de-img-missing" style={size} title={target}>
        {target.split('.').pop()?.toUpperCase()} image
      </span>
    )
  }
  return <img src={url} alt={alt ?? ''} style={{ ...size, verticalAlign: 'bottom', display: 'inline-block' }} draggable={false} />
}

function InlineView({ x, ctx, o }: { x: Inline; ctx: RenderCtx; o: number }): ReactNode {
  // an object inside a tracked change: outlined in the author's colour (dimmed when deleted)
  const objTc = x.t !== 'text' ? x.tc : undefined
  const obj = {
    'data-o': o,
    'data-obj': '',
    contentEditable: ctx.editable ? false : undefined,
    ...(objTc ? { 'data-tc': objTc.k, title: `${objTc.k === 'ins' ? 'Inserted' : 'Deleted'}${objTc.author ? ` by ${objTc.author}` : ''}` } : {}),
  } as const
  switch (x.t) {
    case 'text': {
      const tc = x.tc
      const span = (
        <span
          data-o={o}
          style={tc ? { ...runCss(x.s), ...trackCss(tc) } : runCss(x.s)}
          data-cm={x.cm?.join(' ')}
          data-tc={tc?.k}
          title={tc ? `${tc.k === 'ins' ? 'Inserted' : 'Deleted'}${tc.author ? ` by ${tc.author}` : ''}${tc.date ? ` · ${new Date(tc.date).toLocaleString()}` : ''}` : undefined}
        >
          {x.text}
        </span>
      )
      const href = ctx.measuring || ctx.editable ? undefined : safeHref(x.href)
      return href ? (
        <a href={href} target={href.startsWith('#') ? undefined : '_blank'} rel="noreferrer noopener">
          {span}
        </a>
      ) : (
        span
      )
    }
    case 'field': {
      const v = x.kind === 'PAGE' ? (ctx.page >= 0 ? ctx.page + 1 : 1) : ctx.pages || 1
      return (
        <span {...obj} style={runCss(x.s)}>
          {v}
        </span>
      )
    }
    case 'ref':
      return (
        <span {...obj} style={runCss(x.s)}>
          {x.text}
        </span>
      )
    case 'img': {
      const img = <Img ctx={ctx} target={x.target} w={x.w} h={x.h} alt={x.alt} />
      const href = ctx.measuring || ctx.editable ? undefined : safeHref(x.href)
      return (
        <span {...obj} data-cm={x.cm?.join(' ')}>
          {href ? (
            <a href={href} target="_blank" rel="noreferrer noopener">
              {img}
            </a>
          ) : (
            img
          )}
        </span>
      )
    }
    case 'box':
      return (
        <span
          {...obj}
          className="de-tbox"
          style={{ width: x.w ? `${x.w}pt` : undefined, minHeight: x.h ? `${x.h}pt` : undefined, background: cssColor(x.fill), border: x.border }}
        >
          {x.blocks.map((b, i) => (
            <BlockView key={i} block={b} ctx={ctx} />
          ))}
        </span>
      )
    case 'shape':
      return x.label ? (
        <span {...obj} className="de-ph" style={{ width: `${x.w}pt`, height: `${x.h}pt` }}>
          {x.label}
        </span>
      ) : (
        <span {...obj} />
      )
  }
}

function ParaView({ p, ctx, pa, sp }: { p: Paragraph; ctx: RenderCtx; pa?: string; sp: Spacing }) {
  const hang = p.ps.indFirst && p.ps.indFirst < 0 ? -p.ps.indFirst : 0
  const visible = p.runs.length > 0
  const style = paraCss(p.ps, sp)
  if (!visible) {
    style.fontSize = pt(p.markSz)
    if (p.markFont) style.fontFamily = p.markFont
  }
  const editable = ctx.editable && pa !== undefined
  let o = 0
  const runs = p.runs.map((x, i) => {
    const node = <InlineView key={i} x={x} ctx={ctx} o={o} />
    o += x.t === 'text' ? x.text.length : 1
    return node
  })
  return (
    <div className="de-p" style={style} data-pa={pa} contentEditable={ctx.editable && !editable ? false : undefined}>
      {p.marker && (
        <span className="de-marker" contentEditable={ctx.editable ? false : undefined} style={{ ...runCss(p.marker.s), minWidth: p.marker.tab && hang ? `${hang}pt` : undefined }}>
          {p.marker.text}
          {p.marker.tab && !hang ? '\t' : ''}
        </span>
      )}
      {runs}
      {!visible &&
        (editable ? (
          <span data-o={0} data-empty="">
            {'\u200B'}
          </span>
        ) : (
          '\u200B'
        ))}
    </div>
  )
}

const grids = new WeakMap<Table, ReturnType<typeof gridOf>>()
const gridCached = (t: Table) => {
  let g = grids.get(t)
  if (!g) grids.set(t, (g = gridOf(t)))
  return g
}
const gridRows = new WeakMap<Table, GridCell[][]>()
/** the cells of each row with their grid column */
const gridRowsOf = (t: Table) => {
  let r = gridRows.get(t)
  if (!r) {
    const rows: GridCell[][] = t.rows.map(() => [])
    for (const c of gridCached(t).cells) rows[c.ri].push(c)
    gridRows.set(t, (r = rows))
  }
  return r
}

/**
 * A table, or the rows [from, to) of one split across pages. A part after the first starts with copies of the header
 * rows (`head`): display only, not editable, not addressed. A merged cell cut by the page end is cut too: its content
 * stays on the page it starts on, an empty cell of the same look goes on below.
 */
function TableView({ t, ctx, pa, from = 0, to, head = 0 }: { t: Table; ctx: RenderCtx; pa?: string; from?: number; to?: number; head?: number }) {
  const total = t.cols.reduce((a, b) => a + b, 0)
  const style: CSSProperties = { width: total ? `${total}pt` : '100%' }
  if (t.align === 'center') style.margin = '0 auto'
  else if (t.align === 'right') style.marginLeft = 'auto'
  else if (t.indL) style.marginLeft = `${t.indL}pt`
  const end = Math.min(to ?? t.rows.length, t.rows.length)
  const cut = from > 0 || end < t.rows.length
  const cellView = (c: TableCell, ri: number, ci: number, rowSpan: number, copy: boolean, rest = false) => (
    <td
      key={rest ? `m${ci}` : ci}
      colSpan={c.colSpan > 1 ? c.colSpan : undefined}
      rowSpan={rowSpan > 1 ? rowSpan : undefined}
      contentEditable={rest && ctx.editable ? false : undefined}
      style={{
        borderTop: c.borders.t,
        borderRight: c.borders.r,
        borderBottom: c.borders.b,
        borderLeft: c.borders.l,
        background: cssColor(c.shd),
        verticalAlign: c.vAlign,
        padding: `${c.pad[0]}pt ${c.pad[1]}pt ${c.pad[2]}pt ${c.pad[3]}pt`,
      }}
    >
      {!rest &&
        c.blocks.map((b, i) => (
          <BlockView
            key={b.t === 'p' ? `${i}:${b.ver ?? 0}` : i}
            block={b}
            ctx={ctx}
            pa={!copy && pa !== undefined && b.t === 'p' && b.pid !== undefined && !b.ro ? `${pa}.${ri}.${ci}.${i}` : undefined}
            sp={contextual(b, c.blocks[i - 1], c.blocks[i + 1])}
            nested
          />
        ))}
    </td>
  )
  /** row `ri`, its merged cells cut at row `limit` */
  const rowView = (ri: number, limit: number, copy: boolean) => {
    const row = t.rows[ri]
    let cells: ReactNode[]
    if (!cut) cells = row.cells.map((c, ci) => cellView(c, ri, ci, c.rowSpan, false))
    else {
      const at: { g: number; node: ReactNode }[] = []
      for (const gc of gridRowsOf(t)[ri]) at.push({ g: gc.g, node: cellView(gc.cell, ri, gc.ci, Math.min(gc.cell.rowSpan, limit - ri), copy) })
      // merged cells from the page before go on here, empty
      if (!copy && ri === from)
        for (const gc of gridCached(t).cells) if (gc.ri < from && gc.ri + gc.cell.rowSpan > from) at.push({ g: gc.g, node: cellView(gc.cell, ri, gc.ci, Math.min(gc.ri + gc.cell.rowSpan, limit) - ri, false, true) })
      cells = at.sort((a, b) => a.g - b.g).map((x) => x.node)
    }
    return (
      <tr key={copy ? `h${ri}` : ri} className={copy ? 'de-head-copy' : undefined} contentEditable={copy && ctx.editable ? false : undefined} style={{ height: row.h ? `${row.h}pt` : undefined }}>
        {cells}
      </tr>
    )
  }
  const resize = ctx.colResize && pa !== undefined && !pa.includes('.') && total > 0 ? ctx.colResize : undefined
  const drag = (i: number) => (e: RPointerEvent<HTMLDivElement>) => {
    const tbl = (e.currentTarget.parentElement as HTMLElement).querySelector('table')
    if (!tbl || !resize) return
    e.preventDefault()
    e.stopPropagation()
    const ptPerPx = total / tbl.getBoundingClientRect().width
    const x0 = e.clientX
    const handle = e.currentTarget
    handle.setPointerCapture(e.pointerId)
    const widths = (dx: number) => {
      const cols = [...t.cols]
      const d = dx * ptPerPx
      if (i < cols.length - 1) {
        const room = Math.max(-(cols[i] - 12), Math.min(d, cols[i + 1] - 12))
        cols[i] += room
        cols[i + 1] -= room
      } else cols[i] = Math.max(12, cols[i] + d)
      return cols
    }
    const move = (ev: PointerEvent) => {
      const cols = widths(ev.clientX - x0)
      tbl.querySelectorAll('col').forEach((c, k) => ((c as HTMLElement).style.width = `${cols[k]}pt`))
      tbl.style.width = `${cols.reduce((a, b) => a + b, 0)}pt`
    }
    const up = (ev: PointerEvent) => {
      handle.removeEventListener('pointermove', move)
      handle.removeEventListener('pointerup', up)
      if (Math.abs(ev.clientX - x0) > 1) resize(Number(pa), widths(ev.clientX - x0))
    }
    handle.addEventListener('pointermove', move)
    handle.addEventListener('pointerup', up)
  }
  let edge = 0
  const handles = resize
    ? t.cols.map((w, i) => {
        edge += w
        return <div key={i} className="de-col-resize" contentEditable={false} style={{ left: `calc(${edge}pt - 3px)` }} onPointerDown={drag(i)} />
      })
    : null
  const table = (
    <table className="de-tbl" style={style}>
      {total > 0 && (
        <colgroup>
          {t.cols.map((w, i) => (
            <col key={i} style={{ width: `${w}pt` }} />
          ))}
        </colgroup>
      )}
      <tbody>
        {from > 0 && Array.from({ length: Math.min(head, from) }, (_, ri) => rowView(ri, head, true))}
        {Array.from({ length: end - from }, (_, k) => rowView(from + k, end, false))}
      </tbody>
    </table>
  )
  if (!handles) return table
  const off = style.marginLeft && style.marginLeft !== 'auto' ? style.marginLeft : undefined
  return (
    <div className="de-tbl-wrap" style={{ marginLeft: off, width: style.width, marginRight: t.align === 'center' ? 'auto' : undefined, ...(t.align === 'center' ? { marginLeft: 'auto' } : t.align === 'right' ? { marginLeft: 'auto' } : {}) }}>
      {table}
      {handles}
    </div>
  )
}

export interface BlockViewProps {
  block: Block
  ctx: RenderCtx
  /** editable address (top-level: block index; table: block index used as prefix) */
  pa?: string
  sp?: Spacing
  nested?: boolean
  /**
   * A part of a block split across pages: the characters [from, to) of a paragraph, the rows [from, to) of a table;
   * to undefined: to its end
   */
  from?: number
  to?: number
  /** a table part after the first: its header rows shown again above it */
  head?: number
}

export const BlockView = memo(function BlockView({ block, ctx, pa, sp = '', nested, from, to, head }: BlockViewProps) {
  switch (block.t) {
    case 'p':
      return <ParaView p={block} ctx={ctx} pa={pa} sp={sp} />
    case 'tbl':
      return (
        <div style={{ overflow: 'visible' }} contentEditable={ctx.editable && (nested || pa === undefined) ? false : undefined}>
          <TableView t={block} ctx={ctx} pa={nested ? undefined : pa} from={from} to={to} head={head} />
        </div>
      )
    default:
      return <div />
  }
})
