/**
 * An Excel chart (chart.ts `ChartModel`) drawn as a self-contained SVG picture of `w` × `h` CSS pixels, DOM-free (it
 * runs anywhere): bars and columns (clustered, stacked, 100%), lines (markers, smoothing), areas, scatter, pie and
 * doughnut, combined on primary and secondary axes; axis bounds chosen the way Excel does when the file fixes none
 * (zero included, steps of 1, 2 or 5), major gridlines, tick and data labels in the series' number formats, title and
 * legend. Text is measured with the diagram renderers' widths (diagram/svg.ts), so labels are placed without a browser.
 */
import { lineHeight, pathData, Svg, textWidth, wrap, type Paint, type Pt } from '../diagram/svg'
import type { ChartAxis, ChartGroup, ChartModel, ChartSeries, ChartText } from './chart'
import { formatValue } from './numfmt'

const PX = 96 / 72
const GRID = '#d9d9d9'
const AXIS = '#bfbfbf'
const PAD = 8

interface Box {
  x: number
  y: number
  w: number
  h: number
}

/** a value axis' scale: its bounds, step and ticks, and how a value maps to pixels along `len` */
interface Scale {
  min: number
  max: number
  ticks: number[]
  log?: number
  format?: string
}

const DASH: Record<string, number[]> = {
  dash: [4, 3], sysDash: [3, 1], dot: [1, 1], sysDot: [1, 1], dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1],
  lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1],
}
const dashOf = (d: string | undefined, w: number) => (d && DASH[d] ? DASH[d].map((x) => x * Math.max(1, w)).join(' ') : undefined)

const fmt = (v: number, f: string | undefined, date1904: boolean) => formatValue(v, f || 'General', date1904).text

/** Excel's automatic bounds: zero kept in view unless the values sit far from it, a step of 1, 2 or 5 × 10ⁿ */
export function niceScale(lo: number, hi: number, opts: { min?: number; max?: number; major?: number; zero?: boolean; ticks?: number } = {}): Scale {
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) (lo = 0), (hi = 1)
  if (lo > hi) [lo, hi] = [hi, lo]
  if (opts.zero !== false) {
    if (lo > 0 && lo < (hi * 5) / 6) lo = 0
    if (hi < 0 && hi > (lo * 5) / 6) hi = 0
  }
  if (opts.min !== undefined) lo = opts.min
  if (opts.max !== undefined) hi = opts.max
  if (lo === hi) hi = lo === 0 ? 1 : lo + Math.abs(lo) * 0.1
  const want = Math.max(2, opts.ticks ?? 6)
  let step = opts.major
  if (!step || step <= 0) {
    const rough = (hi - lo) / want
    const mag = 10 ** Math.floor(Math.log10(rough))
    step = ([1, 2, 5, 10].find((m) => m * mag >= rough) ?? 10) * mag
  }
  let min = opts.min ?? Math.floor(lo / step + 1e-9) * step
  let max = opts.max ?? Math.ceil(hi / step - 1e-9) * step
  // the largest value is not left on the top line
  if (opts.max === undefined && max === hi && hi !== 0 && hi > 0) max += step
  if (opts.min === undefined && min === lo && lo < 0) min -= step
  if (max <= min) max = min + step
  const ticks: number[] = []
  for (let k = 0, v = min; v <= max + step * 1e-6 && k < 200; k++, v = min + k * step) ticks.push(Math.abs(v) < step * 1e-9 ? 0 : v)
  return { min, max, ticks }
}

function logScale(lo: number, hi: number, base: number, opts: { min?: number; max?: number }): Scale {
  const b = base > 1 ? base : 10
  const p0 = opts.min ?? b ** Math.floor(Math.log(Math.max(lo, 1e-12)) / Math.log(b))
  const p1 = opts.max ?? b ** Math.ceil(Math.log(Math.max(hi, p0 * b)) / Math.log(b))
  const ticks: number[] = []
  for (let v = p0; v <= p1 * 1.0001 && ticks.length < 60; v *= b) ticks.push(v)
  return { min: p0, max: p1, ticks, log: b }
}

/** a value's position along an axis: 0 at its minimum, 1 at its maximum */
function frac(s: Scale, v: number): number {
  if (s.log) {
    const l = (x: number) => Math.log(Math.max(x, 1e-12))
    return (l(v) - l(s.min)) / (l(s.max) - l(s.min))
  }
  return (v - s.min) / (s.max - s.min)
}

function textBlock(svg: Svg, t: ChartText, x: number, y: number, maxW: number, anchor: 'start' | 'middle' | 'end' = 'middle'): number {
  const size = t.size * PX
  const lines = wrap(t.text.split('\n'), Math.max(40, maxW), size, t.bold)
  const h = lines.length * lineHeight(size)
  svg.text(x, y + h / 2, lines, { size, bold: t.bold, color: t.color, anchor })
  return h
}

// ---------------------------------------------------------------- markers and lines

function marker(svg: Svg, x: number, y: number, m: NonNullable<ChartSeries['marker']>) {
  const r = Math.max(1.5, (m.size * PX) / 2)
  const p: Paint = { fill: m.fill ?? 'none', stroke: m.line, sw: 1 }
  switch (m.symbol) {
    case 'circle':
    case 'dot':
      svg.circle(x, y, m.symbol === 'dot' ? r / 2 : r, p)
      break
    case 'square':
      svg.rect(x - r, y - r, r * 2, r * 2, p)
      break
    case 'diamond':
      svg.poly([[x, y - r], [x + r, y], [x, y + r], [x - r, y]], p)
      break
    case 'triangle':
      svg.poly([[x, y - r], [x + r, y + r], [x - r, y + r]], p)
      break
    case 'x':
      svg.line(x - r, y - r, x + r, y + r, { stroke: m.line ?? m.fill, sw: 1.5 }).line(x - r, y + r, x + r, y - r, { stroke: m.line ?? m.fill, sw: 1.5 })
      break
    case 'plus':
      svg.line(x - r, y, x + r, y, { stroke: m.line ?? m.fill, sw: 1.5 }).line(x, y - r, x, y + r, { stroke: m.line ?? m.fill, sw: 1.5 })
      break
    case 'star': {
      const pts: [number, number][] = []
      for (let k = 0; k < 10; k++) {
        const a = -Math.PI / 2 + (k * Math.PI) / 5
        const rr = k % 2 ? r * 0.45 : r
        pts.push([x + rr * Math.cos(a), y + rr * Math.sin(a)])
      }
      svg.poly(pts, p)
      break
    }
    case 'dash':
      svg.line(x - r, y, x + r, y, { stroke: m.line ?? m.fill, sw: 2 })
      break
    default:
      svg.circle(x, y, r, p)
  }
}

/** a line through `pts`, smoothed as Excel does (a Catmull-Rom curve) when asked */
function linePath(pts: Pt[], smooth: boolean): string {
  if (!smooth || pts.length < 3) return pathData(pts.map((p, k) => ({ c: k ? 'L' : 'M', p: [p] })))
  const cmds: { c: 'M' | 'C'; p: Pt[] }[] = [{ c: 'M', p: [pts[0]] }]
  for (let k = 0; k < pts.length - 1; k++) {
    const p0 = pts[Math.max(0, k - 1)]
    const p1 = pts[k]
    const p2 = pts[k + 1]
    const p3 = pts[Math.min(pts.length - 1, k + 2)]
    cmds.push({ c: 'C', p: [{ x: p1.x + (p2.x - p0.x) / 6, y: p1.y + (p2.y - p0.y) / 6 }, { x: p2.x - (p3.x - p1.x) / 6, y: p2.y - (p3.y - p1.y) / 6 }, p2] })
  }
  return pathData(cmds)
}

// ---------------------------------------------------------------- legend

interface LegendEntry {
  label: string
  fill?: string | null
  line?: ChartSeries['line']
  marker?: ChartSeries['marker']
}

function legendEntries(m: ChartModel): LegendEntry[] {
  const g0 = m.groups[0]
  // a pie (or one series of varied colors): the points of its first series are the entries
  const pie = g0.type === 'pie' || g0.type === 'doughnut'
  if (m.groups.length === 1 && g0.varyColors && (pie || g0.series.length === 1)) {
    const s = g0.series[0]
    return s.cats.map((c, k) => ({ label: c, fill: s.points[k]?.color ?? s.fill })).filter((_, k) => !m.legend?.deleted.includes(k))
  }
  const out: LegendEntry[] = []
  let k = 0
  for (const g of m.groups)
    for (const s of g.series) {
      if (!m.legend?.deleted.includes(k++)) out.push({ label: s.name, fill: g.type === 'line' || g.type === 'scatter' ? undefined : s.fill, line: g.type === 'line' || g.type === 'scatter' ? s.line : undefined, marker: s.marker })
    }
  return out
}

function drawLegend(svg: Svg, m: ChartModel, entries: LegendEntry[], box: Box, vertical: boolean, size: number) {
  const sw = size * 0.8
  const lh = lineHeight(size) + 2
  let x = box.x
  let y = box.y
  const rowW = box.w
  for (const e of entries) {
    const w = sw + 4 + textWidth(e.label, size) + 12
    if (!vertical && x + w > box.x + rowW + 1 && x > box.x) {
      x = box.x
      y += lh
    }
    const cy = y + lh / 2
    if (e.line || e.marker) {
      if (e.line) svg.line(x, cy, x + sw * 1.6, cy, { stroke: e.line.color, sw: Math.min(3, e.line.width * PX), dash: dashOf(e.line.dash, e.line.width) })
      if (e.marker) marker(svg, x + sw * 0.8, cy, { ...e.marker, size: Math.min(e.marker.size, 7) })
      svg.text(x + sw * 1.6 + 4, cy, e.label, { size, color: m.textColor })
      x += sw * 1.6 + 4 + textWidth(e.label, size) + 12
    } else {
      svg.rect(x, cy - sw / 2, sw, sw, { fill: e.fill ?? 'none' })
      svg.text(x + sw + 4, cy, e.label, { size, color: m.textColor })
      x += w
    }
    if (vertical) {
      x = box.x
      y += lh
    }
  }
}

function legendSize(entries: LegendEntry[], vertical: boolean, size: number, maxW: number): { w: number; h: number } {
  const lh = lineHeight(size) + 2
  const widths = entries.map((e) => size * 0.8 * (e.line || e.marker ? 1.6 : 1) + 4 + textWidth(e.label, size) + 12)
  if (vertical) return { w: Math.min(maxW, Math.max(0, ...widths)), h: entries.length * lh }
  let rows = 1
  let x = 0
  for (const w of widths) {
    if (x + w > maxW && x > 0) (rows++, (x = 0))
    x += w
  }
  return { w: Math.min(maxW, widths.reduce((a, b) => a + b, 0)), h: rows * lh }
}

// ---------------------------------------------------------------- pie and doughnut

function drawPie(svg: Svg, m: ChartModel, g: ChartGroup, box: Box) {
  const s = g.series[0]
  const vals = s.values.map((v) => (v && v > 0 ? v : 0))
  const total = vals.reduce((a, b) => a + b, 0)
  if (!total) return
  const exploded = Math.max(0, ...Object.values(s.points).map((p) => p.explosion ?? 0))
  const outside = s.labels && (s.labels.pos === 'outEnd' || s.labels.pos === 'bestFit')
  const r = Math.max(4, (Math.min(box.w, box.h) / 2) * (outside ? 0.75 : 0.92) / (1 + exploded / 100))
  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  const hole = g.type === 'doughnut' ? Math.min(0.9, Math.max(0.1, g.holeSize / 100)) * r : 0
  let a = ((g.firstSliceAng - 90) * Math.PI) / 180
  const size = m.fontSize * PX
  vals.forEach((v, k) => {
    if (!v) return
    const sweep = (v / total) * Math.PI * 2
    const mid = a + sweep / 2
    const off = ((s.points[k]?.explosion ?? 0) / 100) * r
    const ox = cx + off * Math.cos(mid)
    const oy = cy + off * Math.sin(mid)
    const fill = s.points[k]?.color ?? s.fill ?? '#4472c4'
    const stroke = s.line?.color ?? '#ffffff'
    const at = (rr: number, ang: number): Pt => ({ x: ox + rr * Math.cos(ang), y: oy + rr * Math.sin(ang) })
    const large = sweep > Math.PI ? 1 : 0
    if (sweep >= Math.PI * 2 - 1e-6) {
      svg.circle(ox, oy, r, { fill, stroke, sw: 1 })
      if (hole) svg.circle(ox, oy, hole, { fill: m.bg ?? '#ffffff' })
    } else {
      const p1 = at(r, a)
      const p2 = at(r, a + sweep)
      const n = (x: number) => Math.round(x * 100) / 100
      const d = hole
        ? `M${n(p1.x)} ${n(p1.y)} A${n(r)} ${n(r)} 0 ${large} 1 ${n(p2.x)} ${n(p2.y)} L${n(at(hole, a + sweep).x)} ${n(at(hole, a + sweep).y)} A${n(hole)} ${n(hole)} 0 ${large} 0 ${n(at(hole, a).x)} ${n(at(hole, a).y)} Z`
        : `M${n(ox)} ${n(oy)} L${n(p1.x)} ${n(p1.y)} A${n(r)} ${n(r)} 0 ${large} 1 ${n(p2.x)} ${n(p2.y)} Z`
      svg.path(d, { fill, stroke, sw: 1 })
    }
    if (s.labels) {
      const parts: string[] = []
      if (s.labels.series) parts.push(s.name)
      if (s.labels.category) parts.push(s.cats[k] ?? '')
      if (s.labels.value) parts.push(fmt(v, s.labels.format ?? s.format, m.date1904))
      if (s.labels.percent) parts.push(fmt(v / total, s.labels.format && s.labels.format.includes('%') ? s.labels.format : '0%', m.date1904))
      const lr = outside ? r * 1.15 : hole ? (r + hole) / 2 : r * 0.65
      const p = at(lr, mid)
      const anchor = outside ? (Math.cos(mid) >= 0 ? 'start' : 'end') : 'middle'
      svg.text(p.x, p.y, parts.join(s.labels.sep ?? (outside ? '; ' : '\n')).split('\n'), { size, color: outside ? m.textColor : '#ffffff', anchor })
    }
    a += sweep
  })
}

// ---------------------------------------------------------------- axes and plot

interface AxisPlan {
  axis: ChartAxis
  scale: Scale
  /** labels' formatted text */
  labels: string[]
}

const isStacked = (g: ChartGroup) => g.grouping === 'stacked' || g.grouping === 'percentStacked'

/** the extent of a group's values on its value axis (sums when stacked, 0…1 when 100%) */
function extent(g: ChartGroup): [number, number] {
  if (g.grouping === 'percentStacked') {
    const neg = g.series.some((s) => s.values.some((v) => v !== null && v < 0))
    return [neg ? -1 : 0, 1]
  }
  if (isStacked(g) && g.type !== 'line') {
    let lo = 0
    let hi = 0
    const n = Math.max(0, ...g.series.map((s) => s.values.length))
    for (let k = 0; k < n; k++) {
      let p = 0
      let q = 0
      for (const s of g.series) {
        const v = s.values[k] ?? 0
        if (v >= 0) p += v
        else q += v
      }
      hi = Math.max(hi, p)
      lo = Math.min(lo, q)
    }
    return [lo, hi]
  }
  const all = g.series.flatMap((s) => (isStacked(g) ? cumulative(g, s) : s.values)).filter((v): v is number => v !== null)
  return all.length ? [Math.min(...all), Math.max(...all)] : [0, 1]
}

/** a stacked line's points: each series on top of those before it */
function cumulative(g: ChartGroup, s: ChartSeries): (number | null)[] {
  const i = g.series.indexOf(s)
  return s.values.map((v, k) => {
    if (v === null) return null
    let sum = 0
    for (let j = 0; j <= i; j++) sum += g.series[j].values[k] ?? 0
    if (g.grouping === 'percentStacked') {
      const tot = g.series.reduce((a, x) => a + Math.abs(x.values[k] ?? 0), 0)
      return tot ? sum / tot : 0
    }
    return sum
  })
}

export function renderChart(m: ChartModel, W: number, H: number): string {
  const svg = new Svg()
  if (m.border !== null) svg.rect(0.5, 0.5, W - 1, H - 1, { stroke: m.border ?? GRID, sw: 1 })
  const size = m.fontSize * PX
  let top = PAD
  if (m.title) top += textBlock(svg, m.title, W / 2, PAD, W - PAD * 2) + 6
  let area: Box = { x: PAD, y: top, w: W - PAD * 2, h: H - top - PAD }

  // legend: its own strip beside the plot, unless it lies over it
  const entries = m.legend ? legendEntries(m) : []
  if (m.legend && entries.length) {
    const pos = m.legend.pos
    const vertical = pos === 'r' || pos === 'l' || pos === 'tr'
    const ls = legendSize(entries, vertical, size, vertical ? area.w * 0.4 : area.w)
    let lb: Box
    if (pos === 't') lb = { x: area.x + (area.w - ls.w) / 2, y: area.y, w: ls.w, h: ls.h }
    else if (pos === 'b') lb = { x: area.x + (area.w - ls.w) / 2, y: area.y + area.h - ls.h, w: ls.w, h: ls.h }
    else if (pos === 'l') lb = { x: area.x, y: area.y + (area.h - ls.h) / 2, w: ls.w, h: ls.h }
    else if (pos === 'tr') lb = { x: area.x + area.w - ls.w, y: area.y, w: ls.w, h: ls.h }
    else lb = { x: area.x + area.w - ls.w, y: area.y + (area.h - ls.h) / 2, w: ls.w, h: ls.h }
    drawLegend(svg, m, entries, lb, vertical, size)
    if (!m.legend.overlay) {
      if (pos === 't') area = { ...area, y: area.y + ls.h + 6, h: area.h - ls.h - 6 }
      else if (pos === 'b') area = { ...area, h: area.h - ls.h - 6 }
      else if (pos === 'l') area = { ...area, x: area.x + ls.w + 8, w: area.w - ls.w - 8 }
      else area = { ...area, w: area.w - ls.w - 8 }
    }
  }
  // the plot's place where the file fixes it
  let inner: Box | null = null
  if (m.plot) {
    const b = { x: m.plot.x * W, y: m.plot.y * H, w: m.plot.w * W, h: m.plot.h * H }
    if (m.plot.target === 'inner') inner = b
    else area = b
  }

  const pie = m.groups.find((g) => g.type === 'pie' || g.type === 'doughnut')
  if (pie) {
    drawPie(svg, m, pie, inner ?? area)
    return svg.toString(W, H, m.bg === null ? '' : (m.bg ?? '#ffffff'))
  }
  drawCartesian(svg, m, area, inner, size)
  return svg.toString(W, H, m.bg === null ? '' : (m.bg ?? '#ffffff'))
}

function drawCartesian(svg: Svg, m: ChartModel, area: Box, fixed: Box | null, size: number) {
  const axis = (id: string) => m.axes.find((a) => a.id === id)
  const groups = m.groups.filter((g) => g.series.length)
  const horizontal = groups.some((g) => g.type === 'bar' && g.dir === 'bar')
  const lh = lineHeight(size)
  // value axes and the groups on them
  const valIds = [...new Set(groups.map((g) => g.axIds[1]).filter(Boolean))]
  const plans = new Map<string, AxisPlan>()
  const cats = groups.find((g) => g.type !== 'scatter')?.series[0]?.cats ?? []
  const n = Math.max(1, ...groups.filter((g) => g.type !== 'scatter').flatMap((g) => g.series.map((s) => s.values.length)))
  const valLen = horizontal ? area.w : area.h
  for (const id of valIds) {
    const a = axis(id)
    const on = groups.filter((g) => g.axIds[1] === id)
    const ext = on.map(extent)
    const lo = Math.min(...ext.map((e) => e[0]))
    const hi = Math.max(...ext.map((e) => e[1]))
    const pct = on.some((g) => g.grouping === 'percentStacked')
    const scale = a?.logBase ? logScale(lo, hi, a.logBase, a) : niceScale(lo, hi, { min: a?.min, max: a?.max, major: a?.major, ticks: Math.round(valLen / 40) })
    const f = a?.format ?? (pct ? '0%' : on[0]?.series.find((s) => s.format)?.format)
    plans.set(id, { axis: a ?? fallbackAxis(id, horizontal ? 'b' : 'l'), scale: { ...scale, format: f }, labels: scale.ticks.map((t) => fmt(t, f, m.date1904)) })
  }
  // scatter: its x axis is a value axis too
  const scatter = groups.find((g) => g.type === 'scatter')
  let xPlan: AxisPlan | null = null
  if (scatter) {
    const xs = scatter.series.flatMap((s) => (s.xs ?? []).filter((v): v is number => v !== null))
    const a = axis(scatter.axIds[0])
    const known = xs.length ? xs : scatter.series[0].values.map((_, k) => k + 1)
    const sc = niceScale(Math.min(...known), Math.max(...known), { min: a?.min, max: a?.max, major: a?.major, ticks: Math.round(area.w / 70) })
    const f = a?.format ?? scatter.series[0]?.format
    xPlan = { axis: a ?? fallbackAxis(scatter.axIds[0], 'b'), scale: { ...sc, format: f }, labels: sc.ticks.map((t) => fmt(t, f, m.date1904)) }
  }
  const catAxis = axis(groups.find((g) => g.type !== 'scatter')?.axIds[0] ?? '')
  const showCats = !!cats.length && !scatter && !(catAxis?.deleted ?? false) && (catAxis?.labels ?? true)
  const primary = valIds.find((id) => !['r', 't'].includes(plans.get(id)!.axis.pos)) ?? valIds[0]
  const secondary = valIds.find((id) => id !== primary)

  // room for labels and axis titles
  const shows = (p: AxisPlan | undefined) => !!p && !p.axis.deleted && p.axis.labels
  const valW = (p: AxisPlan | undefined) => (shows(p) ? Math.max(...p!.labels.map((l) => textWidth(l, size))) + 6 : 0)
  const catW = showCats ? Math.min(area.w * 0.35, Math.max(...cats.map((c) => textWidth(c, size)))) + 6 : 0
  const titleH = (a: ChartAxis | undefined) => (a?.title && !a.deleted ? lineHeight(a.title.size * PX) + 4 : 0)
  // category labels too wide to stand side by side are tilted by 45°, as Excel does
  const catTextW = showCats ? Math.max(...cats.map((c) => textWidth(c, size))) : 0
  const tilt = showCats && !horizontal && catTextW + 2 > (area.w * 0.85) / n
  const catBottom = tilt ? Math.min(area.h * 0.4, catTextW * Math.SQRT1_2 + lh) : lh + 4
  let inner: Box
  if (fixed) inner = fixed
  else {
    const pp = plans.get(primary)
    const sp = secondary ? plans.get(secondary) : undefined
    let left = 0
    let right = 0
    let bottom = 0
    const topPad = lh / 2
    if (horizontal) {
      left = catW + titleH(catAxis)
      bottom = (shows(pp) ? lh + 4 : 0) + titleH(pp?.axis)
      if (sp && shows(sp)) right = 0
    } else {
      left = valW(pp) + titleH(pp?.axis)
      right = sp ? valW(sp) + titleH(sp.axis) : textWidth(cats.at(-1) ?? '', size) / 3
      bottom = (showCats ? catBottom : xPlan ? lh + 4 : 0) + titleH(xPlan?.axis ?? catAxis)
    }
    inner = { x: area.x + left, y: area.y + topPad, w: Math.max(10, area.w - left - right - 4), h: Math.max(10, area.h - topPad - bottom) }
  }

  // positions: value v of plan p, category k (its band's middle)
  const valPos = (p: AxisPlan, v: number) => {
    let f = frac(p.scale, v)
    if (p.axis.reverse) f = 1 - f
    return horizontal ? inner.x + f * inner.w : inner.y + inner.h - f * inner.h
  }
  const catReverse = catAxis?.reverse ?? false
  const band = (horizontal ? inner.h : inner.w) / n
  const catMid = (k: number) => {
    // horizontal bars run from the bottom up, as Excel draws them
    const i = horizontal !== catReverse ? n - 1 - k : k
    return (horizontal ? inner.y : inner.x) + (i + 0.5) * band
  }
  const xPos = (v: number) => inner.x + (xPlan!.axis.reverse ? 1 - frac(xPlan!.scale, v) : frac(xPlan!.scale, v)) * inner.w

  // gridlines and value labels
  for (const id of valIds) {
    const p = plans.get(id)!
    const right = id !== primary
    for (let k = 0; k < p.scale.ticks.length; k++) {
      const pos = valPos(p, p.scale.ticks[k])
      if (p.axis.gridlines) horizontal ? svg.line(pos, inner.y, pos, inner.y + inner.h, { stroke: GRID }) : svg.line(inner.x, pos, inner.x + inner.w, pos, { stroke: GRID })
      if (!shows(p)) continue
      if (horizontal) svg.text(pos, inner.y + inner.h + 4 + lh / 2, p.labels[k], { size, color: m.textColor, anchor: 'middle' })
      else svg.text(right ? inner.x + inner.w + 4 : inner.x - 4, pos, p.labels[k], { size, color: m.textColor, anchor: right ? 'start' : 'end' })
    }
  }
  if (xPlan) {
    for (let k = 0; k < xPlan.scale.ticks.length; k++) {
      const x = xPos(xPlan.scale.ticks[k])
      if (xPlan.axis.gridlines) svg.line(x, inner.y, x, inner.y + inner.h, { stroke: GRID })
      if (shows(xPlan)) svg.text(x, inner.y + inner.h + 4 + lh / 2, xPlan.labels[k], { size, color: m.textColor, anchor: 'middle' })
    }
  }
  // category labels: flat, tilted where they do not fit flat, and every k-th where they overlap even so
  if (showCats) {
    const space = horizontal || tilt ? lh * (tilt ? Math.SQRT2 : 1) : catTextW + 2
    const every = Math.max(1, Math.ceil(space / Math.max(1, band)))
    const y0 = inner.y + inner.h + 4
    cats.forEach((c, k) => {
      if (k % every) return
      if (horizontal) svg.text(inner.x - 4, catMid(k), c, { size, color: m.textColor, anchor: 'end' })
      else if (tilt) {
        const x = catMid(k)
        svg.raw(`<g transform="rotate(-45 ${Math.round(x)} ${Math.round(y0)})">`).text(x, y0, c, { size, color: m.textColor, anchor: 'end' }).raw('</g>')
      } else svg.text(catMid(k), y0 + lh / 2, c, { size, color: m.textColor, anchor: 'middle' })
    })
  }
  // the category axis line, where the value axis is 0 (or its minimum)
  const pp = plans.get(primary)
  if (pp && !(catAxis?.deleted ?? false)) {
    const at = pp.scale.min <= 0 && pp.scale.max >= 0 && !pp.scale.log ? 0 : pp.scale.min
    const pos = valPos(pp, at)
    horizontal ? svg.line(pos, inner.y, pos, inner.y + inner.h, { stroke: AXIS }) : svg.line(inner.x, pos, inner.x + inner.w, pos, { stroke: AXIS })
  }
  // axis titles
  const axisTitle = (a: ChartAxis | undefined, where: 'left' | 'right' | 'bottom') => {
    if (!a?.title || a.deleted) return
    const s = a.title.size * PX
    const o = { size: s, bold: a.title.bold, color: a.title.color ?? m.textColor, anchor: 'middle' as const }
    if (where === 'bottom') svg.text(inner.x + inner.w / 2, area.y + area.h - lineHeight(s) / 2, a.title.text, o)
    else {
      const x = where === 'left' ? area.x + lineHeight(s) / 2 : area.x + area.w - lineHeight(s) / 2
      const y = inner.y + inner.h / 2
      svg.raw(`<g transform="rotate(-90 ${Math.round(x)} ${Math.round(y)})">`).text(x, y, a.title.text, o).raw('</g>')
    }
  }
  if (horizontal) {
    axisTitle(catAxis, 'left')
    axisTitle(pp?.axis, 'bottom')
  } else {
    axisTitle(pp?.axis, 'left')
    if (secondary) axisTitle(plans.get(secondary)?.axis, 'right')
    axisTitle(xPlan?.axis ?? catAxis, 'bottom')
  }

  // the series: areas first, then bars, then lines and points on top
  const order = (g: ChartGroup) => (g.type === 'area' ? 0 : g.type === 'bar' ? 1 : 2)
  for (const g of [...groups].sort((a, b) => order(a) - order(b))) {
    const p = plans.get(g.axIds[1]) ?? plans.get(primary)
    if (!p) continue
    if (g.type === 'bar') drawBars(svg, m, g, p, { valPos, catMid, band, horizontal, size })
    else if (g.type === 'area') drawArea(svg, g, p, { valPos, catMid })
    else if (g.type === 'line') drawLines(svg, m, g, p, { valPos, catMid, size, horizontal })
    else if (g.type === 'scatter' && xPlan) drawScatter(svg, m, g, p, { valPos, xPos, size })
  }
}

function fallbackAxis(id: string, pos: string): ChartAxis {
  return { id, kind: 'val', pos, deleted: false, reverse: false, gridlines: true, labels: true, crossAx: '', crosses: 'autoZero' }
}

interface Geo {
  valPos: (p: AxisPlan, v: number) => number
  catMid: (k: number) => number
}

function label(svg: Svg, m: ChartModel, s: ChartSeries, k: number, v: number, x: number, y: number, size: number, anchor: 'start' | 'middle' | 'end' = 'middle', color?: string) {
  const l = s.labels!
  const parts: string[] = []
  if (l.series) parts.push(s.name)
  if (l.category) parts.push(s.cats[k] ?? '')
  if (l.value) parts.push(fmt(v, l.format ?? s.format, m.date1904))
  if (parts.length) svg.text(x, y, parts.join(l.sep ?? ', '), { size, color: color ?? m.textColor, anchor })
}

function drawBars(svg: Svg, m: ChartModel, g: ChartGroup, p: AxisPlan, geo: Geo & { band: number; horizontal: boolean; size: number }) {
  const { valPos, catMid, band, horizontal, size } = geo
  const stacked = isStacked(g)
  const n = stacked ? 1 : g.series.length
  const ov = Math.max(-1, Math.min(1, g.overlap / 100))
  const w = band / (n - (n - 1) * ov + g.gapWidth / 100)
  const total = n * w - (n - 1) * w * ov
  const base = p.scale.min <= 0 && p.scale.max >= 0 && !p.scale.log ? 0 : p.scale.min
  const len = Math.max(...g.series.map((s) => s.values.length))
  for (let k = 0; k < len; k++) {
    let posSum = 0
    let negSum = 0
    const tot = g.grouping === 'percentStacked' ? g.series.reduce((a, s) => a + Math.abs(s.values[k] ?? 0), 0) || 1 : 1
    g.series.forEach((s, i) => {
      const raw = s.values[k]
      if (raw === null || raw === undefined) return
      const v = raw / tot
      let from = base
      let to = v
      if (stacked) {
        from = v >= 0 ? posSum : negSum
        to = from + v
        if (v >= 0) posSum = to
        else negSum = to
      }
      const off = stacked ? 0 : i * w * (1 - ov)
      const c0 = catMid(k) - total / 2 + off
      const a = valPos(p, from)
      const b = valPos(p, to)
      const fill = raw < 0 && s.invertIfNegative ? '#ffffff' : (s.points[k]?.color ?? s.fill ?? 'none')
      const paint: Paint = { fill, stroke: s.line?.color ?? (raw < 0 && s.invertIfNegative ? (s.fill ?? undefined) : undefined), sw: s.line ? Math.max(0.75, s.line.width * PX) : 1 }
      if (horizontal) svg.rect(Math.min(a, b), c0, Math.abs(b - a), w, paint)
      else svg.rect(c0, Math.min(a, b), w, Math.abs(b - a), paint)
      if (s.labels) {
        const mid = (a + b) / 2
        const inside = stacked || s.labels.pos === 'ctr' || s.labels.pos === 'inEnd' || s.labels.pos === 'inBase'
        if (horizontal) label(svg, m, s, k, raw, inside ? mid : Math.max(a, b) + 4, c0 + w / 2, size, inside ? 'middle' : 'start', inside ? '#ffffff' : undefined)
        else label(svg, m, s, k, raw, c0 + w / 2, inside ? mid : Math.min(a, b) - lineHeight(size) / 2, size, 'middle', inside ? '#ffffff' : undefined)
      }
    })
  }
}

function drawArea(svg: Svg, g: ChartGroup, p: AxisPlan, geo: Geo) {
  const { valPos, catMid } = geo
  const base = p.scale.min <= 0 && p.scale.max >= 0 ? 0 : p.scale.min
  let below: number[] | null = null
  for (const s of g.series) {
    const vals = isStacked(g) ? cumulative(g, s).map((v) => v ?? 0) : s.values.map((v) => v ?? 0)
    const top = vals.map((v, k) => [catMid(k), valPos(p, v)] as [number, number])
    const floor = (below && isStacked(g) ? below.map((v, k) => [catMid(k), valPos(p, v)] as [number, number]) : vals.map((_, k) => [catMid(k), valPos(p, base)] as [number, number])).reverse()
    svg.poly([...top, ...floor], { fill: s.fill ?? 'none', stroke: s.line?.color, sw: s.line ? s.line.width * PX : 0 })
    below = vals
  }
}

function drawLines(svg: Svg, m: ChartModel, g: ChartGroup, p: AxisPlan, geo: Geo & { size: number; horizontal: boolean }) {
  const { valPos, catMid, size } = geo
  for (const s of g.series) {
    const vals = isStacked(g) ? cumulative(g, s) : s.values
    // empty cells leave a gap in the line
    const runs: Pt[][] = [[]]
    vals.forEach((v, k) => {
      if (v === null) runs.push([])
      else runs[runs.length - 1].push({ x: catMid(k), y: valPos(p, v) })
    })
    if (s.line) for (const run of runs) if (run.length > 1) svg.path(linePath(run, s.smooth), { stroke: s.line.color, sw: Math.max(0.75, s.line.width * PX), dash: dashOf(s.line.dash, s.line.width) })
    vals.forEach((v, k) => {
      if (v === null) return
      const x = catMid(k)
      const y = valPos(p, v)
      if (s.marker) marker(svg, x, y, s.marker)
      if (s.labels) label(svg, m, s, k, s.values[k] ?? v, x, y - lineHeight(size) * 0.8, size)
    })
  }
}

function drawScatter(svg: Svg, m: ChartModel, g: ChartGroup, p: AxisPlan, geo: { valPos: Geo['valPos']; xPos: (v: number) => number; size: number }) {
  const { valPos, xPos, size } = geo
  const lines = g.scatterStyle !== 'marker' && g.scatterStyle !== 'none'
  for (const s of g.series) {
    const pts: Pt[] = []
    s.values.forEach((v, k) => {
      const xv = s.xs?.[k] ?? k + 1
      if (v === null || xv === null) return
      pts.push({ x: xPos(xv), y: valPos(p, v) })
    })
    if (lines && s.line && pts.length > 1) svg.path(linePath(pts, s.smooth || g.scatterStyle === 'smooth' || g.scatterStyle === 'smoothMarker'), { stroke: s.line.color, sw: Math.max(0.75, s.line.width * PX), dash: dashOf(s.line.dash, s.line.width) })
    pts.forEach((pt, k) => {
      if (s.marker) marker(svg, pt.x, pt.y, s.marker)
      if (s.labels && s.values[k] !== null) label(svg, m, s, k, s.values[k]!, pt.x, pt.y - lineHeight(size) * 0.8, size)
    })
  }
}
