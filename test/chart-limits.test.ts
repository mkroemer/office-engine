import { describe, expect, test } from 'bun:test'
import { PictureTooLarge } from '../src/core/diagram/svg'
import { DEFAULT_THEME } from '../src/core/theme'
import { parseXml } from '../src/core/xml'
import { MAX_CHART_POINTS, MAX_READ_POINTS, pointBudget, readChart, type ChartModel, type PointBudget, type RangeReader } from '../src/core/xlsx/chart'
import { MAX_3D_MARKS } from '../src/core/xlsx/chart-3d'
import { MAX_CHART_SVG, MAX_LABELS, MAX_LEGEND, MAX_MARKERS, renderChart } from '../src/core/xlsx/chart-svg'
import { chartPicturesHeld, chartUrl, MAX_CHART_PICTURES, MIN_CHART_SVG, releaseChart } from '../src/react/common/chart'

// What a chart may cost, whatever its file says. A series that only claims 100 000 points (`c:ptCount`) is a few bytes
// of the file: a few kilobytes of chart parts must not become gigabytes read or hundreds of megabytes drawn.

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
const AXES = (valAx = '') =>
  `<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/>${valAx}</c:valAx><c:serAx><c:axId val="3"/><c:crossAx val="2"/></c:serAx>`
const AX_IDS = '<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>'
const VIEW = '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/></c:view3D>'
/** a series that says it has `count` points and gives none of them */
const claimed = (count: number) => (i: number) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:val><c:numLit><c:ptCount val="${count}"/></c:numLit></c:val></c:ser>`
/** a series whose values are read from the cells (as many as the test's reader gives) */
const filled = (i: number) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:val><c:numRef><c:f>Data!$B$1:$B$9</c:f></c:numRef></c:val></c:ser>`
const series = (n: number, one: (i: number) => string) => Array.from({ length: n }, (_, i) => one(i)).join('')
const chartXml = (plot: string, head = '') => `<c:chartSpace ${NS}><c:chart><c:autoTitleDeleted val="1"/>${head}<c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`

/** the chart read, its references answered with `cells` values each (none: its saved values stand in) */
async function model(xml: string, opts: { cells?: number; budget?: PointBudget } = {}): Promise<ChartModel> {
  const read: RangeReader = async () => {
    const values = Array.from({ length: opts.cells ?? 0 }, (_, k) => 10 + (k % 7))
    return opts.cells ? { values, formats: values.map(() => 'General'), hidden: values.map(() => false) } : null
  }
  const m = await readChart(parseXml(xml), { theme: DEFAULT_THEME, date1904: false, read, budget: opts.budget })
  if (!m) throw new Error('no chart read')
  return m
}

const COLUMNS = chartXml(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series(2, filled)}${AX_IDS}</c:barChart>${AXES()}`)
const count = (svg: string, tag: string) => svg.split(`<${tag}`).length - 1
/** how many text elements of the picture read `text` */
const texts = (svg: string, text: string) => svg.split(`>${text}</text>`).length - 1

describe('chart limits: the picture', () => {
  test('a radar of series that claim 100 000 points: rings as circles, no spoke and no label per point', async () => {
    // 200 rings (a step of 0.005 from 0 to 1) with a corner per point each were 296 MB, drawn for six seconds
    const radar = chartXml(`<c:radarChart><c:radarStyle val="standard"/>${series(2, claimed(100_000))}${AX_IDS}</c:radarChart>${AXES('<c:majorUnit val="0.005"/>')}`)
    const svg = renderChart(await model(radar), 640, 400)
    expect(svg.length).toBeLessThan(200_000)
    expect(count(svg, 'circle')).toBeGreaterThan(100)
    expect(count(svg, 'polygon')).toBe(0)
    expect(count(svg, 'line')).toBe(0)
    expect(count(svg, 'text')).toBeLessThan(1000)
  })

  test('a radar of a few points keeps its cornered rings and its spokes', async () => {
    const radar = chartXml(`<c:radarChart><c:radarStyle val="marker"/>${series(2, filled)}${AX_IDS}</c:radarChart>${AXES()}`)
    const svg = renderChart(await model(radar, { cells: 6 }), 640, 400)
    expect(count(svg, 'polygon')).toBeGreaterThan(2)
    expect(count(svg, 'line')).toBe(6)
  })

  test('a picture past its room is given up, not finished', async () => {
    // the wireframe of 250 series that claim 1 000 points each: 51 MB of lines
    const wire = await model(chartXml(`<c:surfaceChart><c:wireframe val="1"/>${series(250, claimed(1000))}${AX_IDS}</c:surfaceChart>${AXES()}`))
    expect(() => renderChart(wire, 640, 400)).toThrow(PictureTooLarge)
    // an ordinary chart is drawn in the room it needs, and not in less
    const columns = await model(COLUMNS, { cells: 9 })
    const svg = renderChart(columns, 640, 400)
    expect(svg.length).toBeLessThan(MAX_CHART_SVG / 100)
    expect(() => renderChart(columns, 640, 400, svg.length >> 1)).toThrow(PictureTooLarge)
    expect(renderChart(columns, 640, 400, svg.length)).toBe(svg)
  })

  test('a 3D chart of more marks than it is drawn with is refused before its solids are built', async () => {
    const surface = (rows: number, points: number) => chartXml(`<c:surface3DChart>${series(rows, claimed(points))}${AX_IDS}</c:surface3DChart>${AXES()}`, VIEW)
    const bars = chartXml(`<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>${series(3, filled)}<c:shape val="cylinder"/>${AX_IDS}</c:bar3DChart>${AXES()}`, VIEW)
    const pie = chartXml(`<c:pie3DChart><c:varyColors val="1"/>${series(1, filled)}</c:pie3DChart>`, VIEW)
    // (250 × 1 000 points were cut into their bands for 3.8 s; 100 000 cylinders were built for 3.5 s)
    const refused = [await model(surface(250, 1000)), await model(bars, { cells: 10_000 }), await model(pie, { cells: MAX_3D_MARKS + 1 })]
    const t0 = performance.now()
    for (const m of refused) expect(() => renderChart(m, 640, 400)).toThrow(`${MAX_3D_MARKS} marks`)
    expect(performance.now() - t0).toBeLessThan(1000)
    // within it, they are drawn
    expect(count(renderChart(await model(surface(20, 20)), 640, 400), 'polygon')).toBeGreaterThan(300)
    expect(count(renderChart(await model(bars, { cells: 50 }), 640, 400), 'polygon')).toBeGreaterThan(150)
    expect(count(renderChart(await model(pie, { cells: 50 }), 640, 400), 'polygon')).toBeGreaterThan(50)
  })

  test('series labels and markers past what a picture is drawn with: every k-th point shows its own', async () => {
    // a line of 100 000 points that plots its empty ones as zero, each with a marker and a label, was 28 MB
    const labels = (own = '') => `<c:dLbls>${own}<c:showVal val="1"/></c:dLbls>`
    const line = (n: number, extra: string) =>
      chartXml(`<c:lineChart><c:grouping val="standard"/>${claimed(n)(0).replace('<c:val>', `${extra}<c:val>`)}${AX_IDS}</c:lineChart>${AXES()}`).replace('</c:plotArea>', '</c:plotArea><c:dispBlanksAs val="zero"/>')
    const dense = renderChart(await model(line(100_000, labels())), 640, 400)
    // (the line's markers are diamonds: polygons)
    expect(count(dense, 'polygon')).toBe(MAX_MARKERS)
    expect(texts(dense, '0')).toBeGreaterThan(MAX_LABELS * 0.9)
    expect(count(dense, 'text')).toBeLessThan(MAX_LABELS * 1.2)
    expect(dense.length).toBeLessThan(MAX_CHART_SVG)
    // within them, all are drawn; a point's own label is shown among thinned ones
    const few = renderChart(await model(line(1500, labels())), 640, 400)
    expect(count(few, 'polygon')).toBe(1500)
    expect(texts(few, '0')).toBeGreaterThanOrEqual(1500)
    // (point 7777 is not one of every 50th: it shows the series' name because it says so itself)
    const own = `<c:dLbl><c:idx val="7777"/><c:showVal val="0"/><c:showSerName val="1"/></c:dLbl>`
    const named = line(100_000, `<c:tx><c:v>Only</c:v></c:tx>${labels(own)}`)
    expect(texts(renderChart(await model(named), 640, 400), 'Only')).toBe(1)
  })

  test('a legend lists no more entries than it is drawn with', async () => {
    // a pie's legend has an entry per point: of 20 000 points, and with as many entries deleted one by one
    const deleted = Array.from({ length: 20_000 }, (_, k) => `<c:legendEntry><c:idx val="${k * 2 + 1}"/><c:delete val="1"/></c:legendEntry>`).join('')
    const pie = chartXml(`<c:pieChart><c:varyColors val="1"/>${series(1, filled)}</c:pieChart>`).replace('</c:plotArea>', `</c:plotArea><c:legend><c:legendPos val="r"/>${deleted}</c:legend>`)
    const m = await model(pie, { cells: 40_000 })
    const t0 = performance.now()
    const svg = renderChart(m, 640, 400)
    expect(performance.now() - t0).toBeLessThan(1500)
    // a swatch and a label an entry, a path a slice
    expect(count(svg, 'rect')).toBeLessThan(MAX_LEGEND + 10)
    expect(count(svg, 'path')).toBe(40_000)
  })

  test('an axis whose ticks run past the largest number still draws', async () => {
    // a log axis from 1e308: its next tick is Infinity, which no number format takes
    const log = COLUMNS.replace('<c:majorGridlines/>', '<c:scaling><c:logBase val="2"/><c:min val="1e308"/></c:scaling><c:numFmt formatCode="0.0E+00" sourceLinked="0"/><c:majorGridlines/>')
    expect(renderChart(await model(log, { cells: 9 }), 640, 400).startsWith('<svg')).toBe(true)
  })

  test('a 3D surface whose value axis has no step between its bounds draws no cell, and does not throw', async () => {
    const flat = chartXml(`<c:surface3DChart><c:wireframe val="1"/>${series(3, filled)}${AX_IDS}</c:surface3DChart>${AXES('<c:scaling><c:max val="1"/></c:scaling><c:majorUnit val="1000"/>')}`, VIEW)
    expect(renderChart(await model(flat, { cells: 4 }), 640, 400).startsWith('<svg')).toBe(true)
  })
})

describe('chart limits: the points read', () => {
  // 200 000 points a chart: each within its own budget
  const big = chartXml(`<c:lineChart><c:grouping val="standard"/>${series(2, claimed(100_000))}${AX_IDS}</c:lineChart>${AXES()}`)

  test('charts read together share a budget: one past it is not read, and takes nothing of it', async () => {
    const budget = pointBudget()
    for (let k = 0; k < MAX_READ_POINTS / 200_000; k++) expect((await model(big, { budget })).groups[0].series.length).toBe(2)
    expect(budget.left).toBe(0)
    await expect(model(big, { budget })).rejects.toThrow(`${MAX_READ_POINTS} points`)
    // the chart refused gave back what it had taken, so a small one after it is still read
    expect(budget.left).toBe(0)
    budget.left = 18
    expect((await model(COLUMNS, { cells: 9, budget })).groups[0].series[0].values.length).toBe(9)
    expect(budget.left).toBe(0)
  })

  test('a chart that is not read for its own size, or has no kind read, takes nothing of the shared budget', async () => {
    const budget = pointBudget()
    // three series of 100 000 points: past the chart's own 250 000
    await expect(model(chartXml(`<c:lineChart>${series(3, claimed(100_000))}${AX_IDS}</c:lineChart>${AXES()}`), { budget })).rejects.toThrow(`${MAX_CHART_POINTS} points`)
    expect(budget.left).toBe(MAX_READ_POINTS)
    expect(await readChart(parseXml(chartXml('<c:unknownChart/>')), { theme: DEFAULT_THEME, date1904: false, read: async () => null, budget })).toBeNull()
    expect(budget.left).toBe(MAX_READ_POINTS)
  })

  test('without a shared budget a chart has its own alone', async () => {
    for (let k = 0; k < 7; k++) expect((await model(big)).groups[0].series[0].values.length).toBe(100_000)
  })
})

describe('chart limits: the pictures of a page', () => {
  test('a chart given up for its size holds its room until it is let go; an ordinary chart is drawn all the same', async () => {
    const before = chartPicturesHeld()
    const wire = chartXml(`<c:surfaceChart><c:wireframe val="1"/>${series(250, claimed(1000))}${AX_IDS}</c:surfaceChart>${AXES()}`)
    const warn = console.warn
    const warned: string[] = []
    console.warn = (...a: unknown[]) => void warned.push(a.join(' '))
    const huge = [await model(wire), await model(wire), await model(wire)]
    const columns = await model(COLUMNS, { cells: 9 })
    try {
      // two take the page's room between them
      expect(chartUrl(huge[0], 640, 400)).toBe('')
      expect(chartPicturesHeld() - before).toBe(MAX_CHART_SVG)
      expect(chartUrl(huge[1], 640, 400)).toBe('')
      expect(chartPicturesHeld() - before).toBe(MAX_CHART_PICTURES)
      // the third has the least room only: given up after a megabyte, not after 32
      const t0 = performance.now()
      expect(chartUrl(huge[2], 640, 400)).toBe('')
      const third = performance.now() - t0
      expect(chartPicturesHeld() - before).toBe(MAX_CHART_PICTURES + MIN_CHART_SVG)
      expect(third).toBeLessThan(500)
      expect(warned.length).toBe(3)
      // not drawn again at the same size
      expect(chartUrl(huge[2], 640, 400)).toBe('')
      expect(warned.length).toBe(3)
      // an ordinary chart still has its room
      const url = chartUrl(columns, 640, 400)
      expect(url.startsWith('blob:')).toBe(true)
      expect(chartUrl(columns, 640, 400)).toBe(url)
    } finally {
      console.warn = warn
      for (const m of [...huge, columns]) releaseChart(m)
    }
    expect(chartPicturesHeld()).toBe(before)
  })

  test('copies of a chart that has a key share one picture, kept while one of them is shown', async () => {
    const before = chartPicturesHeld()
    const columns = { ...(await model(COLUMNS, { cells: 9 })), key: 'doc:ppt/charts/chart1.xml' }
    const copies = [structuredClone(columns), structuredClone(columns), structuredClone(columns)]
    const url = chartUrl(copies[0], 640, 400)
    const one = chartPicturesHeld() - before
    expect(url.startsWith('blob:')).toBe(true)
    // read again (a slide after an edit): the picture there is, not one more
    expect(chartUrl(copies[1], 640, 400)).toBe(url)
    expect(chartUrl(copies[2], 640, 400)).toBe(url)
    expect(chartPicturesHeld() - before).toBe(one)
    releaseChart(copies[0])
    releaseChart(copies[1])
    expect(chartUrl(copies[2], 640, 400)).toBe(url)
    expect(chartPicturesHeld() - before).toBe(one)
    releaseChart(copies[2])
    expect(chartPicturesHeld()).toBe(before)
    // another document's chart of the same part name has its own picture
    const other = { ...columns, key: 'doc2:ppt/charts/chart1.xml' }
    expect(chartUrl(other, 640, 400)).not.toBe(url)
    releaseChart(other)
    expect(chartPicturesHeld()).toBe(before)
  })

  test('a chart drawn again at another size holds its new picture alone', async () => {
    const before = chartPicturesHeld()
    const columns = await model(COLUMNS, { cells: 9 })
    const small = chartUrl(columns, 320, 200)
    const one = chartPicturesHeld() - before
    expect(one).toBeGreaterThan(500)
    const large = chartUrl(columns, 640, 400)
    expect(large).not.toBe(small)
    expect(chartPicturesHeld() - before).toBeLessThan(one * 1.5)
    releaseChart(columns)
    expect(chartPicturesHeld()).toBe(before)
    // let go twice: nothing more to give back
    releaseChart(columns)
    expect(chartPicturesHeld()).toBe(before)
  })
})
