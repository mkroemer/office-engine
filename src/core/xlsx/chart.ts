/**
 * An Excel chart part (`xl/charts/chartN.xml`) read into what the SVG renderer (chart-svg.ts) draws: its plot groups
 * (bar / column, line, area, pie, doughnut, scatter; 3D ones as their flat form; several groups make a combo chart, each
 * on its own axes), series with their values, axes, title, legend and the plot area's manual layout.
 *
 * Series values are read from the workbook's cells (`RangeReader`), so a chart follows edited cells; where a reference
 * cannot be read there (a defined name, another workbook, an unknown sheet) the values the file saved with the chart
 * (`numCache` / `strCache`) stand in. Colors are resolved here against the workbook's theme: the view gets CSS colors.
 * A chart of another kind (radar, bubble, stock, surface, the newer chartex kinds) has no model: it stays a placeholder.
 */
import { colorOf, cssColor, drawingColor, type ColorCtx } from '../color'
import type { Theme } from '../theme'
import { kid, kids, num, textOf, type XEl } from '../xml'
import { formatValue } from './numfmt'

export type ChartType = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter'
export type Grouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard'

export interface ChartText {
  text: string
  /** pt */
  size: number
  bold: boolean
  color?: string
}

export interface DataLabels {
  value: boolean
  percent: boolean
  category: boolean
  series: boolean
  /** ctr, inEnd, inBase, outEnd, t, b, l, r, bestFit */
  pos?: string
  format?: string
  sep?: string
}

export interface ChartPoint {
  color?: string
  /** pie slices: % of the radius */
  explosion?: number
}

export interface ChartSeries {
  name: string
  /** category labels (bar, line, area, pie) */
  cats: string[]
  /** x values (scatter) */
  xs?: (number | null)[]
  values: (number | null)[]
  /** the values' number format */
  format?: string
  /** fill (bars, areas, slices); null: none */
  fill: string | null
  /** line (lines, scatter; bar outlines); null: none */
  line: { color: string; width: number; dash?: string } | null
  marker: { symbol: string; size: number; fill?: string; line?: string } | null
  smooth: boolean
  invertIfNegative: boolean
  points: Record<number, ChartPoint>
  labels?: DataLabels
}

export interface ChartGroup {
  type: ChartType
  /** bars: 'col' upright, 'bar' horizontal */
  dir: 'col' | 'bar'
  grouping: Grouping
  /** % of a bar's width */
  gapWidth: number
  /** % of a bar's width, -100…100 */
  overlap: number
  varyColors: boolean
  /** doughnut: % of the radius */
  holeSize: number
  /** pie: degrees clockwise from 12 o'clock */
  firstSliceAng: number
  scatterStyle?: string
  /** the ids of its category (x) and value (y) axes */
  axIds: string[]
  series: ChartSeries[]
}

export interface ChartAxis {
  id: string
  kind: 'cat' | 'val' | 'date'
  /** where it is drawn: l, r, t, b */
  pos: string
  deleted: boolean
  reverse: boolean
  logBase?: number
  min?: number
  max?: number
  major?: number
  format?: string
  gridlines: boolean
  /** tick labels shown */
  labels: boolean
  title?: ChartText
  crossAx: string
  crosses: 'autoZero' | 'min' | 'max'
  crossesAt?: number
}

/** fractions of the chart's size */
export interface ChartLayout {
  x: number
  y: number
  w: number
  h: number
  /** 'inner': the plot's own box; 'outer': with its axis labels */
  target: 'inner' | 'outer'
}

export interface ChartModel {
  title?: ChartText
  legend?: { pos: string; overlay: boolean; deleted: number[] }
  groups: ChartGroup[]
  axes: ChartAxis[]
  plot?: ChartLayout
  /** pt: the chart's base text size */
  fontSize: number
  textColor: string
  bg?: string | null
  border?: string | null
  date1904: boolean
}

/** a reference's cells, row by row: values (numbers or text), number formats, and which lie in hidden rows / columns */
export interface RangeValues {
  values: (number | string | null)[]
  formats: (string | undefined)[]
  hidden: boolean[]
}

/** reads a series reference (`'Sheet 1'!$B$2:$B$9`) from the workbook; null where it cannot */
export type RangeReader = (ref: string) => Promise<RangeValues | null>

const GROUPS: Record<string, ChartType> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter',
}

const val = (el: XEl | undefined, def?: string) => el?.a.val ?? def
const on = (el: XEl | undefined, def = false) => (el ? el.a.val === undefined || el.a.val === '1' || el.a.val === 'true' : def)
const EMU_PER_PT = 12700

/** the text of a title / rich text body, its size and look from the first run or the body's defaults */
function richText(el: XEl | undefined, ctx: ColorCtx, defSize: number, defBold: boolean): ChartText | undefined {
  const rich = kid(kid(el, 'tx'), 'rich')
  if (!rich) return undefined
  const paras = kids(rich, 'p')
  const text = paras.map((p) => kids(p, 'r').map((r) => textOf(kid(r, 't'))).join('') + kids(p, 'fld').map((f) => textOf(kid(f, 't'))).join('')).join('\n').trim()
  if (!text) return undefined
  const rPr = kid(kids(paras[0], 'r')[0], 'rPr') ?? kid(kid(paras[0], 'pPr'), 'defRPr')
  return { text, ...runLook(rPr, ctx, defSize, defBold) }
}

function runLook(rPr: XEl | undefined, ctx: ColorCtx, defSize: number, defBold: boolean) {
  const sz = num(rPr?.a.sz)
  return {
    size: sz ? sz / 100 : defSize,
    bold: rPr?.a.b !== undefined ? rPr.a.b === '1' || rPr.a.b === 'true' : defBold,
    color: safe(colorOf(kid(rPr, 'solidFill'), ctx)),
  }
}

const safe = (c: string | undefined) => (c && cssColor(c) ? c : undefined)

/** a text body's default run size and look (`c:txPr`) */
function txLook(el: XEl | undefined, ctx: ColorCtx) {
  const rPr = kid(kid(kids(kid(el, 'txPr'), 'p')[0], 'pPr'), 'defRPr')
  return rPr ? runLook(rPr, ctx, 0, false) : undefined
}

/** an accent of the theme, darker or lighter in later rounds (series 7 on), as Excel colors series without their own */
function autoColor(i: number, ctx: ColorCtx): string {
  const accent = `accent${(i % 6) + 1}`
  const round = Math.floor(i / 6)
  const mods: XEl[] = []
  const mod = (n: string, v: number) => mods.push({ n, a: { val: String(v) }, c: [], t: '' })
  if (round % 4 === 1) mod('lumMod', 60000)
  else if (round % 4 === 2) (mod('lumMod', 80000), mod('lumOff', 20000))
  else if (round % 4 === 3) mod('lumMod', 80000)
  return drawingColor({ n: 'schemeClr', a: { val: accent }, c: mods, t: '' }, ctx) ?? '#4472c4'
}

/** a shape's fill: a color, null for none, undefined where it says nothing */
function fillOf(spPr: XEl | undefined, ctx: ColorCtx): string | null | undefined {
  if (!spPr) return undefined
  if (kid(spPr, 'noFill')) return null
  const solid = kid(spPr, 'solidFill')
  if (solid) return safe(colorOf(solid, ctx)) ?? undefined
  const grad = kid(spPr, 'gradFill')
  // a gradient: its first stop's color
  if (grad) return safe(colorOf(kid(kid(grad, 'gsLst'), 'gs'), ctx)) ?? undefined
  return undefined
}

function lineOf(spPr: XEl | undefined, ctx: ColorCtx): { color?: string; width?: number; dash?: string } | null | undefined {
  const ln = kid(spPr, 'ln')
  if (!ln) return undefined
  if (kid(ln, 'noFill')) return null
  const w = num(ln.a.w)
  const dash = val(kid(ln, 'prstDash'))
  return { color: safe(colorOf(kid(ln, 'solidFill'), ctx)), width: w !== undefined ? w / EMU_PER_PT : undefined, dash: dash && dash !== 'solid' ? dash : undefined }
}

function labelsOf(el: XEl | undefined): DataLabels | undefined {
  if (!el || on(kid(el, 'delete'))) return undefined
  const l: DataLabels = {
    value: on(kid(el, 'showVal')),
    percent: on(kid(el, 'showPercent')),
    category: on(kid(el, 'showCatName')),
    series: on(kid(el, 'showSerName')),
    pos: val(kid(el, 'dLblPos')),
    format: kid(el, 'numFmt')?.a.formatCode,
    sep: textOf(kid(el, 'separator')) || undefined,
  }
  return l.value || l.percent || l.category || l.series ? l : undefined
}

/** a cache's points by index (`c:pt idx`), `ptCount` long */
function cacheOf(cache: XEl | undefined): { values: (string | null)[]; format?: string } | null {
  if (!cache) return null
  const n = num(val(kid(cache, 'ptCount')), 0)!
  const values: (string | null)[] = new Array(n).fill(null)
  for (const pt of kids(cache, 'pt')) {
    const i = num(pt.a.idx, -1)!
    if (i >= 0 && i < 100_000) values[i] = textOf(kid(pt, 'v'))
  }
  return { values, format: textOf(kid(cache, 'formatCode')) || undefined }
}

interface Read {
  values: (number | string | null)[]
  format?: string
  hidden?: boolean[]
}

/** a title whose text is a cell (`c:tx/c:strRef`): the cell's text, else the cached one */
async function referencedText(el: XEl, ctx: ColorCtx, read: RangeReader, size: number): Promise<ChartText | undefined> {
  const text = (await source(kid(el, 'tx'), read))?.values.filter((x) => x !== null && x !== '').join(' ')
  return text ? { text: String(text), size, bold: false, ...txLook(el, ctx) } : undefined
}

/** a data source (`c:cat`, `c:val`, `c:xVal`, `c:yVal`, `c:tx`): its reference read from the cells, else its cache */
async function source(el: XEl | undefined, read: RangeReader): Promise<Read | null> {
  if (!el) return null
  const ref = kid(el, 'numRef') ?? kid(el, 'strRef') ?? kid(el, 'multiLvlStrRef')
  const lit = kid(el, 'numLit') ?? kid(el, 'strLit')
  const f = textOf(kid(ref, 'f'))
  if (f) {
    const got = await read(f).catch(() => null)
    if (got) return { values: got.values, format: got.formats.find((x) => x && x !== 'General'), hidden: got.hidden }
  }
  const cache = cacheOf(kid(ref, 'numCache') ?? kid(ref, 'strCache') ?? lit) ?? multiLevel(kid(ref, 'multiLvlStrCache'))
  return cache ? { values: cache.values, format: cache.format } : null
}

/** multi-level categories: the innermost level (the labels next to the axis) */
function multiLevel(cache: XEl | undefined) {
  const lvl = kids(cache, 'lvl')[0]
  if (!cache || !lvl) return null
  return cacheOf({ ...lvl, c: [kid(cache, 'ptCount')!, ...lvl.c].filter(Boolean) })
}

const toNum = (v: number | string | null | undefined): number | null => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null
  if (v === null || v === undefined || v === '') return null
  const n = Number(v)
  return Number.isFinite(n) ? n : null
}

async function seriesOf(ser: XEl, i: number, type: ChartType, varyColors: boolean, ctx: ColorCtx, read: RangeReader, plotVisOnly: boolean, date1904: boolean): Promise<ChartSeries> {
  const tx = await source(kid(ser, 'tx'), read)
  const name = tx?.values.filter((x) => x !== null && x !== '').join(' ') || textOf(kid(kid(ser, 'tx'), 'v')) || `Series${num(val(kid(ser, 'idx')), i)! + 1}`
  const scatter = type === 'scatter'
  const catSrc = await source(kid(ser, scatter ? 'xVal' : 'cat'), read)
  const valSrc = await source(kid(ser, scatter ? 'yVal' : 'val'), read)
  let values = (valSrc?.values ?? []).map(toNum)
  // categories that are numbers or dates are shown in their format, text as it is
  const catText = (v: number | string | null) => {
    if (v === null) return ''
    const n = typeof v === 'number' ? v : catSrc?.format ? toNum(v) : null
    return n === null ? String(v) : formatValue(n, catSrc?.format || 'General', date1904).text
  }
  let cats = (catSrc?.values ?? []).map(catText)
  let xs = scatter ? (catSrc?.values ?? []).map(toNum) : undefined
  // hidden rows / columns are left out where the chart shows visible cells only
  const hidden = valSrc?.hidden
  if (plotVisOnly && hidden?.some(Boolean)) {
    const keep = (_: unknown, k: number) => !hidden[k]
    values = values.filter(keep)
    cats = cats.filter(keep)
    xs = xs?.filter(keep)
  }
  if (!cats.length) cats = values.map((_, k) => String(k + 1))
  const spPr = kid(ser, 'spPr')
  const color = autoColor(num(val(kid(ser, 'idx')), i)!, ctx)
  const ownFill = fillOf(spPr, ctx)
  const ownLine = lineOf(spPr, ctx)
  const lined = type === 'line' || type === 'scatter'
  const fill = lined ? null : ownFill === undefined ? color : ownFill
  const line =
    ownLine === null
      ? null
      : lined
        ? { color: ownLine?.color ?? color, width: ownLine?.width ?? 2.25, dash: ownLine?.dash }
        : ownLine?.color
          ? { color: ownLine.color, width: ownLine.width ?? 0.75, dash: ownLine.dash }
          : null
  const mk = kid(ser, 'marker')
  const symbol = val(kid(mk, 'symbol'), lined ? 'auto' : 'none')!
  const markerSpPr = kid(mk, 'spPr')
  const marker =
    symbol === 'none' || !lined
      ? null
      : {
          symbol: symbol === 'auto' ? ['diamond', 'square', 'triangle', 'x', 'star', 'circle'][i % 6] : symbol,
          size: num(val(kid(mk, 'size')), 5)!,
          fill: fillOf(markerSpPr, ctx) ?? line?.color ?? color,
          line: lineOf(markerSpPr, ctx)?.color ?? line?.color ?? color,
        }
  const points: Record<number, ChartPoint> = {}
  // pie slices (and the points of a single series told to vary its colors) get the theme's accents in turn
  if (varyColors) values.forEach((_, k) => (points[k] = { color: autoColor(k, ctx) }))
  for (const dPt of kids(ser, 'dPt')) {
    const k = num(val(kid(dPt, 'idx')), -1)!
    if (k < 0) continue
    const c = fillOf(kid(dPt, 'spPr'), ctx)
    points[k] = { ...points[k], ...(c !== undefined && c !== null ? { color: c } : {}), ...(kid(dPt, 'explosion') ? { explosion: num(val(kid(dPt, 'explosion')), 0) } : {}) }
  }
  const explosion = num(val(kid(ser, 'explosion')))
  if (explosion) values.forEach((_, k) => (points[k] = { explosion, ...points[k] }))
  return {
    name,
    cats,
    xs,
    values,
    format: valSrc?.format,
    fill,
    line,
    marker,
    smooth: on(kid(ser, 'smooth')),
    invertIfNegative: on(kid(ser, 'invertIfNegative')),
    points,
    labels: labelsOf(kid(ser, 'dLbls')),
  }
}

function axisOf(el: XEl, ctx: ColorCtx, defSize: number): ChartAxis {
  const scaling = kid(el, 'scaling')
  const fmt = kid(el, 'numFmt')
  const crosses = val(kid(el, 'crosses'))
  return {
    id: val(kid(el, 'axId'), '')!,
    kind: el.n === 'catAx' ? 'cat' : el.n === 'dateAx' ? 'date' : 'val',
    pos: val(kid(el, 'axPos'), 'b')!,
    deleted: on(kid(el, 'delete')),
    reverse: val(kid(scaling, 'orientation')) === 'maxMin',
    logBase: num(val(kid(scaling, 'logBase'))),
    min: num(val(kid(scaling, 'min'))),
    max: num(val(kid(scaling, 'max'))),
    major: num(val(kid(el, 'majorUnit'))),
    format: fmt && fmt.a.sourceLinked !== '1' ? fmt.a.formatCode : undefined,
    gridlines: !!kid(el, 'majorGridlines'),
    labels: val(kid(el, 'tickLblPos'), 'nextTo') !== 'none',
    title: richText(kid(el, 'title'), ctx, defSize, true),
    crossAx: val(kid(el, 'crossAx'), '')!,
    crosses: crosses === 'min' || crosses === 'max' ? crosses : 'autoZero',
    crossesAt: num(val(kid(el, 'crossesAt'))),
  }
}

function layoutOf(el: XEl | undefined): ChartLayout | undefined {
  const m = kid(kid(el, 'layout'), 'manualLayout')
  if (!m) return undefined
  const x = num(val(kid(m, 'x')))
  const y = num(val(kid(m, 'y')))
  const w = num(val(kid(m, 'w')))
  const h = num(val(kid(m, 'h')))
  // only 'edge' positions (fractions from the top left) are read; 'factor' ones move Excel's own layout
  if (x === undefined || y === undefined || w === undefined || h === undefined || val(kid(m, 'xMode')) === 'factor' || val(kid(m, 'yMode')) === 'factor') return undefined
  return { x, y, w, h, target: val(kid(m, 'layoutTarget')) === 'inner' ? 'inner' : 'outer' }
}

/**
 * The chart in `root` (a chart part's `c:chartSpace`), or null where it has no kind this reads (it stays a placeholder).
 * `read` gives a series reference's cells; `theme` resolves its colors.
 */
export async function readChart(root: XEl | undefined, opts: { theme: Theme; date1904: boolean; read: RangeReader }): Promise<ChartModel | null> {
  const chart = kid(root, 'chart')
  const plot = kid(chart, 'plotArea')
  if (!chart || !plot) return null
  const ctx: ColorCtx = { theme: opts.theme }
  const base = txLook(root, ctx)
  const fontSize = base?.size || 10
  const textColor = base?.color ?? '#595959'
  const plotVisOnly = on(kid(chart, 'plotVisOnly'), true)
  const groups: ChartGroup[] = []
  let seriesIndex = 0
  for (const g of plot.c) {
    const type = GROUPS[g.n]
    if (!type) {
      // a kind this cannot draw (radar, bubble, stock, surface): the whole chart stays a placeholder
      if (g.n.endsWith('Chart')) return null
      continue
    }
    const grouping = val(kid(g, 'grouping'), type === 'bar' ? 'clustered' : 'standard') as Grouping
    const varyColors = on(kid(g, 'varyColors'), type === 'pie' || type === 'doughnut')
    const series: ChartSeries[] = []
    // a pie shows its first series, a slice per point in its own color; other kinds vary colors for a lone series only
    const vary = varyColors && (type === 'pie' || type === 'doughnut' || kids(g, 'ser').length === 1)
    for (const ser of kids(g, 'ser')) series.push(await seriesOf(ser, seriesIndex++, type, vary, ctx, opts.read, plotVisOnly, opts.date1904))
    const groupLabels = labelsOf(kid(g, 'dLbls'))
    if (groupLabels) for (const s of series) s.labels ??= groupLabels
    groups.push({
      type,
      dir: val(kid(g, 'barDir')) === 'bar' ? 'bar' : 'col',
      grouping,
      gapWidth: num(val(kid(g, 'gapWidth')), 150)!,
      overlap: num(val(kid(g, 'overlap')), grouping === 'stacked' || grouping === 'percentStacked' ? 100 : 0)!,
      varyColors,
      holeSize: num(val(kid(g, 'holeSize')), 50)!,
      firstSliceAng: num(val(kid(g, 'firstSliceAng')), 0)!,
      scatterStyle: val(kid(g, 'scatterStyle')),
      axIds: kids(g, 'axId').map((a) => a.a.val ?? ''),
      series,
    })
  }
  if (!groups.length || !groups.some((g) => g.series.length)) return null
  const axes = plot.c.filter((a) => a.n === 'catAx' || a.n === 'valAx' || a.n === 'dateAx').map((a) => axisOf(a, ctx, fontSize))
  // a title without text of its own is Excel's automatic one: the series' name for one series, else "Chart Title"; a chart
  // of one series without a title element has that automatic title too unless it was deleted
  const single = groups.length === 1 && groups[0].series.length === 1
  const titleEl = kid(chart, 'title')
  const auto = (): ChartText => ({ text: single ? groups[0].series[0].name : 'Chart Title', size: fontSize * 1.4, bold: false, ...txLook(titleEl, ctx) })
  let title = titleEl ? (richText(titleEl, ctx, fontSize * 1.4, false) ?? (await referencedText(titleEl, ctx, opts.read, fontSize * 1.4)) ?? auto()) : undefined
  if (!titleEl && single && !on(kid(chart, 'autoTitleDeleted')) && groups[0].type !== 'scatter') title = auto()
  if (title) {
    if (!title.size) title.size = fontSize * 1.4
    title.color ??= textColor
  }
  const legend = kid(chart, 'legend')
  const spPr = kid(root, 'spPr')
  const border = lineOf(spPr, ctx)
  return {
    title,
    legend: legend ? { pos: val(kid(legend, 'legendPos'), 'r')!, overlay: on(kid(legend, 'overlay')), deleted: kids(legend, 'legendEntry').filter((e) => on(kid(e, 'delete'))).map((e) => num(val(kid(e, 'idx')), -1)!) } : undefined,
    groups,
    axes,
    plot: layoutOf(plot),
    fontSize,
    textColor,
    bg: fillOf(spPr, ctx),
    border: border === null ? null : border?.color,
    date1904: opts.date1904,
  }
}
