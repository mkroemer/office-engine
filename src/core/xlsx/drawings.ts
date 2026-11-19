/**
 * The pictures and charts on a sheet, read from its drawing part (`xl/drawings/drawingN.xml`, linked from the sheet's
 * rels), each with the place it is anchored to: two cell corners (`twoCellAnchor`), one corner and a size
 * (`oneCellAnchor`), or a position on the sheet and a size (`absoluteAnchor`). A picture carries its media part, which
 * the view shows where the browser can draw it (not EMF / WMF); a chart carries its kind and title, for a placeholder
 * the size of the chart. Shapes, text boxes and the like are left out, as are drawings in a group the anchor's own
 * frame does not hold.
 */
import { isWebImage, type OpcPackage } from '../opc'
import { desc, descAll, kid, textOf, type XEl } from '../xml'
import { MAX_COL, MAX_ROWS, MAX_SHEET_WIDTH_PX, MAX_SHEET_HEIGHT_PX, ooxmlInt } from './limits'

/** A cell corner: 0-based column and row, and the offset into that cell in CSS pixels */
export interface DrawingCorner {
  c: number
  r: number
  dx: number
  dy: number
}

export interface SheetDrawing {
  kind: 'picture' | 'chart'
  from?: DrawingCorner
  to?: DrawingCorner
  /** position from the sheet's top left, px (absoluteAnchor) */
  at?: { x: number; y: number }
  /** px (oneCellAnchor, absoluteAnchor) */
  size?: { w: number; h: number }
  /** a picture's media part, where the browser can draw it */
  target?: string
  /** a chart's title, or a picture's description */
  title: string
  /** a chart's kind: 'column', 'bar', 'line', 'pie', … ('' where it is none the reader knows) */
  chart?: string
}

/** DrawingML measures in EMU: 914400 to the inch, 96 CSS pixels to the inch */
const EMU_PER_PX = 9525

const CHART_KINDS: Record<string, string> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie',
  doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter', bubbleChart: 'bubble', radarChart: 'radar',
  stockChart: 'stock', surfaceChart: 'surface', surface3DChart: 'surface',
}

/** an attribute whatever its namespace prefix: `r:id` is the usual, not the only one */
const attr = (el: XEl | undefined, local: string) => {
  if (!el) return undefined
  for (const [k, v] of Object.entries(el.a)) if (k === local || k.endsWith(':' + local)) return v
  return undefined
}

function emu(value: string | undefined, maxPx: number, signed = false): number | undefined {
  const n = ooxmlInt(value, signed ? -maxPx * EMU_PER_PX : 0, maxPx * EMU_PER_PX)
  return n === undefined ? undefined : n / EMU_PER_PX
}

function corner(el: XEl | undefined): DrawingCorner | undefined {
  if (!el) return undefined
  const c = ooxmlInt(textOf(kid(el, 'col')), 0, MAX_COL - 1)
  const r = ooxmlInt(textOf(kid(el, 'row')), 0, MAX_ROWS - 1)
  // an offset left out is none: the corner is at the cell's own (a wrong one still discards the anchor)
  const off = (name: string, max: number) => (kid(el, name) ? emu(textOf(kid(el, name)), max, true) : 0)
  const dx = off('colOff', MAX_SHEET_WIDTH_PX)
  const dy = off('rowOff', MAX_SHEET_HEIGHT_PX)
  if (c === undefined || r === undefined || dx === undefined || dy === undefined) return undefined
  return { c, r, dx, dy }
}

/** Reject an invalid anchor as a whole; do not move corrupt drawings to the sheet's last cell. */
function placement(anchor: XEl): Pick<SheetDrawing, 'from' | 'to' | 'at' | 'size'> | undefined {
  const from = corner(kid(anchor, 'from'))
  if (anchor.n === 'twoCellAnchor') {
    const to = corner(kid(anchor, 'to'))
    return from && to ? { from, to } : undefined
  }
  const ext = kid(anchor, 'ext')
  const w = emu(ext?.a.cx, MAX_SHEET_WIDTH_PX)
  const h = emu(ext?.a.cy, MAX_SHEET_HEIGHT_PX)
  if (w === undefined || h === undefined) return undefined
  const size = { w, h }
  if (anchor.n === 'oneCellAnchor') return from ? { from, size } : undefined
  const pos = kid(anchor, 'pos')
  const x = emu(pos?.a.x, MAX_SHEET_WIDTH_PX, true)
  const y = emu(pos?.a.y, MAX_SHEET_HEIGHT_PX, true)
  return x !== undefined && y !== undefined ? { at: { x, y }, size } : undefined
}

/** a chart's kind and title, from its chart part */
async function chartOf(pkg: OpcPackage, part: string): Promise<{ chart: string; title: string }> {
  const root = await pkg.xml(part)
  const plot = desc(root, 'plotArea')
  const type = plot?.c.find((e) => e.n.endsWith('Chart'))
  let chart = type ? (CHART_KINDS[type.n] ?? '') : ''
  // a bar chart is drawn with upright columns unless its direction says bars
  if (type?.n.startsWith('bar') && attr(kid(type, 'barDir'), 'val') !== 'bar') chart = 'column'
  const title = desc(root, 'title')
  return { chart, title: title ? descAll(title, 't').map((t) => t.t).join('').trim() : '' }
}

/** The pictures and charts on the sheet at `sheetPath`, in the order they are drawn */
export async function readSheetDrawings(pkg: OpcPackage, sheetPath: string): Promise<SheetDrawing[]> {
  const sheetRels = await pkg.rels(sheetPath)
  const out: SheetDrawing[] = []
  for (const rel of sheetRels.values()) {
    if (rel.external || !rel.type.endsWith('/drawing') || !pkg.has(rel.target)) continue
    const root = await pkg.xml(rel.target)
    const rels = await pkg.rels(rel.target)
    for (const anchor of root?.c ?? []) {
      if (anchor.n !== 'twoCellAnchor' && anchor.n !== 'oneCellAnchor' && anchor.n !== 'absoluteAnchor') continue
      const place = placement(anchor)
      if (!place) continue
      const chartRef = desc(anchor, 'chart')
      const blip = desc(anchor, 'blip')
      if (chartRef) {
        const target = rels.get(attr(chartRef, 'id') ?? '')
        const chart = target && !target.external && pkg.has(target.target) ? await chartOf(pkg, target.target) : { chart: '', title: '' }
        out.push({ kind: 'chart', ...place, ...chart })
      } else if (blip) {
        const target = rels.get(attr(blip, 'embed') ?? '')
        const pr = desc(anchor, 'cNvPr')
        out.push({
          kind: 'picture',
          ...place,
          target: target && !target.external && isWebImage(target.target) ? target.target : undefined,
          title: pr?.a.descr || pr?.a.title || '',
        })
      }
    }
  }
  return out
}
