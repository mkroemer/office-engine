import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { rowsToXlsx } from '../src/core/convert'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'
import type { SheetDrawing } from '../src/core/xlsx/drawings'
import type { ChartModel } from '../src/core/xlsx/chart'
import { niceScale, renderChart } from '../src/core/xlsx/chart-svg'

const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const ROWS = [
  ['Quarter', 'Sales', 'Cost'],
  ['Q1', '10', '4'],
  ['Q2', '20', '8'],
  ['Q3', '15', '9'],
]

/** a numeric data source: a reference and a cache whose values differ from the cells (so the test tells them apart) */
const numRef = (f: string, cache: number[], fmt = 'General') =>
  `<c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${cache.length}"/>${cache.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`
const strRef = (f: string, cache: string[]) =>
  `<c:strRef><c:f>${f}</c:f><c:strCache><c:ptCount val="${cache.length}"/>${cache.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`
const ser = (i: number, name: string, val: string, extra = '') =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strRef(`Data!$${name}$1`, ['cached name'])}</c:tx>${extra}<c:cat>${strRef('Data!$A$2:$A$4', ['x', 'y', 'z'])}</c:cat><c:val>${val}</c:val></c:ser>`
const AXES = '<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>'

/** a workbook whose sheet "Data" holds ROWS and one chart (its `plotArea` body given); `hideRow` hides a row (1-based) */
async function withChart(plot: string, opts: { hideRow?: number; title?: string; titleXml?: string; style?: number } = {}): Promise<Blob> {
  const base = await rowsToXlsx(ROWS, 'Data')
  const zip = await ZipArchive.open(new BlobSource(base))
  const files = new Map<string, string>()
  let sheet = (await zip.text('xl/worksheets/sheet1.xml')).replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>')
  if (!sheet.includes('xmlns:r=')) sheet = sheet.replace('<worksheet ', `<worksheet xmlns:r="${R}" `)
  if (opts.hideRow) sheet = sheet.replace(`<row r="${opts.hideRow}">`, `<row r="${opts.hideRow}" hidden="1">`)
  files.set('xl/worksheets/sheet1.xml', sheet)
  files.set('xl/worksheets/_rels/sheet1.xml.rels', `${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
  files.set('xl/drawings/_rels/drawing1.xml.rels', `${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${R}/chart" Target="../charts/chart1.xml"/></Relationships>`)
  files.set(
    'xl/drawings/drawing1.xml',
    `${HEAD}<xdr:wsDr ${NS}><xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
  )
  const title = opts.titleXml ?? (opts.title ? `<c:title><c:tx><c:rich><a:p><a:r><a:t>${opts.title}</a:t></a:r></a:p></c:rich></c:tx></c:title>` : '')
  files.set('xl/charts/chart1.xml', `${HEAD}<c:chartSpace ${NS}>${opts.style ? `<c:style val="${opts.style}"/>` : ''}<c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`)
  const ct = (await zip.text('[Content_Types].xml')).replace(
    '</Types>',
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
  )
  files.set('[Content_Types].xml', ct)
  return rebuildZip(zip, files, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}

async function modelOf(book: Blob, engine = new Engine(() => {})): Promise<{ model: ChartModel | undefined; engine: Engine; drawing: SheetDrawing }> {
  if (!(engine as unknown as { xlsx?: unknown }).xlsx) await engine.open(book, {})
  const [drawing] = (await engine.call('xlsx.drawings', [0])) as SheetDrawing[]
  return { model: drawing?.model, engine, drawing }
}

const CLUSTERED = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}${ser(1, 'C', numRef('Data!$C$2:$C$4', [2, 2, 2]))}<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${AXES}`

describe('xlsx charts: model', () => {
  test('series read from the cells, not the saved cache; names, categories, theme colors', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED, { title: 'Revenue' }))
    expect(model).toBeDefined()
    const g = model!.groups[0]
    expect(g).toMatchObject({ type: 'bar', dir: 'col', grouping: 'clustered', gapWidth: 150, axIds: ['1', '2'] })
    expect(g.series.map((s) => s.name)).toEqual(['Sales', 'Cost'])
    expect(g.series[0].cats).toEqual(['Q1', 'Q2', 'Q3'])
    expect(g.series[0].values).toEqual([10, 20, 15])
    expect(g.series[1].values).toEqual([4, 8, 9])
    // no spPr: the theme's first two accents
    expect(g.series[0].fill).toMatch(/^#[0-9a-f]{6}$/i)
    expect(g.series[0].fill).not.toBe(g.series[1].fill)
    expect(model!.title?.text).toBe('Revenue')
    expect(model!.legend?.pos).toBe('b')
    expect(model!.axes.map((a) => [a.id, a.kind, a.gridlines])).toEqual([['1', 'cat', false], ['2', 'val', true]])
  })

  test('a defined name or another workbook: the saved cache stands in', async () => {
    const plot = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef('SalesRange', [7, 8, 9]))}${ser(1, 'C', numRef('[1]Data!$C$2:$C$4', [3, 2, 1]))}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const { model } = await modelOf(await withChart(plot))
    expect(model!.groups[0].series.map((s) => s.values)).toEqual([[7, 8, 9], [3, 2, 1]])
    expect(model!.groups[0].series[0].marker).not.toBeNull()
  })

  test('hidden rows are left out of a chart of visible cells', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED, { hideRow: 3 }))
    expect(model!.groups[0].series[0].values).toEqual([10, 15])
    expect(model!.groups[0].series[0].cats).toEqual(['Q1', 'Q3'])
  })

  test('an edited cell is in the chart read after the edit', async () => {
    const book = await withChart(CLUSTERED)
    const events: { event: string; data: { sheet?: number; done?: boolean } }[] = []
    const engine = new Engine((event, data) => events.push({ event, data: data as { sheet?: number; done?: boolean } }))
    await engine.open(book, {})
    await engine.call('xlsx.load', [0])
    while (!events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === 0 && e.data.done)) await new Promise((r) => setTimeout(r, 5))
    await engine.call('xlsx.set', [0, [{ r: 2, c: 1, input: '42' }]])
    const { model } = await modelOf(book, engine)
    expect(model!.groups[0].series[0].values).toEqual([10, 42, 15])
  })

  test('pie: one series, a color per slice; a single series titles the chart', async () => {
    const plot = `<c:pieChart><c:varyColors val="1"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), '<c:dPt><c:idx val="1"/><c:explosion val="20"/></c:dPt>')}<c:firstSliceAng val="90"/></c:pieChart>`
    const { model } = await modelOf(await withChart(plot))
    const s = model!.groups[0].series[0]
    expect(model!.groups[0].type).toBe('pie')
    expect(model!.groups[0].firstSliceAng).toBe(90)
    expect(Object.keys(s.points)).toEqual(['0', '1', '2'])
    expect(new Set(Object.values(s.points).map((p) => p.color)).size).toBe(3)
    expect(s.points[1].explosion).toBe(20)
    expect(model!.title?.text).toBe('Sales')
  })

  test('a kind not drawn (radar) keeps the placeholder', async () => {
    const plot = `<c:radarChart><c:radarStyle val="marker"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}<c:axId val="1"/><c:axId val="2"/></c:radarChart>${AXES}`
    const { model, drawing } = await modelOf(await withChart(plot))
    expect(model).toBeUndefined()
    expect(drawing.chart).toBe('radar')
  })
})

describe('xlsx charts: looks', () => {
  const colored = (spPr: string) => CLUSTERED.replace('<c:tx>', `<c:spPr>${spPr}</c:spPr><c:tx>`).replace(`${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}`, ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), `<c:spPr>${spPr}</c:spPr>`))

  test('a gradient fill is drawn as a gradient, a pattern fill as a pattern', async () => {
    const grad = '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="1F4E79"/></a:gs><a:gs pos="100000"><a:srgbClr val="9DC3E6"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>'
    const { model } = await modelOf(await withChart(colored(grad)))
    const s = model!.groups[0].series[0]
    expect(s.fill?.toLowerCase()).toBe('#1f4e79')
    expect(s.look).toMatchObject({ kind: 'gradient', angle: 90, radial: false })
    const svg = renderChart(model!, 400, 300)
    expect(svg).toMatch(/<linearGradient id="de-fill\d+"[^>]*>.*stop-color="#9DC3E6"/i)
    expect(svg).toMatch(/<rect [^>]*fill="url\(#de-fill\d+\)"/)

    const patt = '<a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="C55A11"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>'
    const { model: pm } = await modelOf(await withChart(colored(patt)))
    expect(pm!.groups[0].series[0].look).toMatchObject({ kind: 'pattern', preset: 'dkDnDiag', fg: '#c55a11', bg: '#ffffff' })
    expect(renderChart(pm!, 400, 300)).toMatch(/<pattern id="de-fill\d+"[^>]*>.*stroke="#C55A11"/i)
  })

  test('axis and legend text, gridlines and the plot area in the colors the file gives', async () => {
    const tx = (c: string) => `<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="${c}"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>`
    const axes = `<c:catAx><c:axId val="1"/><c:axPos val="b"/>${tx('AA0001')}<c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="4B5563"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:majorGridlines>${tx('AA0002')}<c:crossAx val="1"/></c:valAx><c:spPr><a:solidFill><a:srgbClr val="374151"/></a:solidFill></c:spPr>`
    const plot = CLUSTERED.replace(AXES, axes)
    const book = await withChart(plot)
    const { model } = await modelOf(book)
    model!.legend!.text = { color: '#aa0003' }
    const svg = renderChart(model!, 400, 300)
    expect(svg).toMatch(/fill="#AA0001"[^>]*>Q1</i)
    expect(svg).toMatch(/fill="#AA0002"[^>]*>25</i)
    expect(svg).toMatch(/fill="#aa0003"[^>]*>Sales</)
    expect(svg).toMatch(/<line [^>]*stroke="#4B5563"[^>]*stroke-dasharray/i)
    expect(model!.plotFill?.toLowerCase()).toBe('#374151')
    expect(svg).toContain('fill="#374151"')
  })

  test('the theme gives the default grays; a dark built-in style a dark chart with light text', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED))
    expect(model!.gridColor.toLowerCase()).toBe('#d9d9d9')
    expect(model!.textColor.toLowerCase()).toBe('#595959')
    const dark = await withChart(CLUSTERED, { style: 42 })
    const { model: dm } = await modelOf(dark)
    expect(dm!.bg).toBeTruthy()
    expect(dm!.textColor.toLowerCase()).toBe('#ffffff')
    // style 42 is the second of the dark row: the theme's accents in turn
    expect(dm!.groups[0].series.map((s) => s.fill)).toEqual(model!.groups[0].series.map((s) => s.fill))
    // style 3, one accent's shades: two series, two shades of accent 1
    const { model: mono } = await modelOf(await withChart(CLUSTERED, { style: 3 }))
    const [a, b] = mono!.groups[0].series.map((s) => s.fill)
    expect(a).not.toBe(b)
  })
})

describe('xlsx charts: drawing', () => {
  test('axis bounds as Excel picks them', () => {
    expect(niceScale(4, 20, { ticks: 6 })).toMatchObject({ min: 0, max: 25, ticks: [0, 5, 10, 15, 20, 25] })
    expect(niceScale(95, 100, { ticks: 5 }).min).toBeGreaterThan(80)
    expect(niceScale(-3, 8, { ticks: 6 }).ticks[0]).toBeLessThan(-2)
    expect(niceScale(0, 1, { min: 0, max: 1, major: 0.25 }).ticks).toEqual([0, 0.25, 0.5, 0.75, 1])
  })

  test('a clustered column chart: a bar per value, the categories, gridlines and the legend', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED, { title: 'Revenue' }))
    const svg = renderChart(model!, 480, 300)
    expect(svg.startsWith('<svg')).toBe(true)
    // nothing that runs or loads: url() only names the picture's own clip path, gradients and patterns
    expect(svg).not.toMatch(/<script|href=|url\((?!#de-)/i)
    const fills = new Set(model!.groups[0].series.map((s) => s.fill))
    const bars = [...svg.matchAll(/<rect [^>]*fill="([^"]+)"/g)].filter((m) => fills.has(m[1]))
    // 3 categories × 2 series, plus a legend swatch each
    expect(bars.length).toBe(8)
    for (const t of ['Revenue', 'Q1', 'Q2', 'Q3', 'Sales', 'Cost', '25']) expect(svg).toContain(`>${t}<`)
  })

  test('stacked bars, lines and pies draw without errors and with their values', async () => {
    const stacked = CLUSTERED.replace('<c:grouping val="clustered"/>', '<c:grouping val="stacked"/><c:overlap val="100"/>').replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>')
    const line = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), '<c:dLbls><c:showVal val="1"/></c:dLbls>')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES}`
    const doughnut = `<c:doughnutChart><c:varyColors val="1"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]), '<c:dLbls><c:showPercent val="1"/></c:dLbls>')}<c:holeSize val="60"/></c:doughnutChart>`
    for (const plot of [stacked, line, doughnut]) {
      const { model } = await modelOf(await withChart(plot))
      const svg = renderChart(model!, 400, 260)
      expect(svg).toMatch(/^<svg[\s\S]*<\/svg>$/)
      expect(svg).not.toContain('NaN')
    }
    const { model: lm } = await modelOf(await withChart(line))
    expect(renderChart(lm!, 400, 260)).toContain('>20<')
    const { model: dm } = await modelOf(await withChart(doughnut))
    expect(renderChart(dm!, 400, 260)).toContain('>44%<')
  })

  test('100% stacked: the axis ends at 100%, with no step above', async () => {
    const pct = CLUSTERED.replace('<c:grouping val="clustered"/>', '<c:grouping val="percentStacked"/><c:overlap val="100"/>')
    const { model } = await modelOf(await withChart(pct))
    const svg = renderChart(model!, 400, 300)
    expect(svg).toContain('>100%<')
    expect(svg).not.toContain('>120%<')
    expect(svg).not.toContain('>110%<')
  })

  test('a title from a cell: the cell text, not the automatic one', async () => {
    const titleXml = `<c:title><c:tx>${strRef('Data!$C$1', ['cached title'])}</c:tx></c:title>`
    const { model } = await modelOf(await withChart(CLUSTERED, { titleXml }))
    expect(model!.title?.text).toBe('Cost')
  })

  test('pie: a negative value is a slice of its magnitude, its label keeps the sign', async () => {
    const plot = `<c:pieChart><c:varyColors val="1"/>${ser(0, 'B', numRef('Data!$B$2:$B$4', [1, 1, 1]))}</c:pieChart>`
    const { model } = await modelOf(await withChart(plot))
    const s = model!.groups[0].series[0]
    s.values = [-10, 20, 30]
    s.labels = { value: true, percent: false, category: false, series: false }
    const svg = renderChart(model!, 400, 260)
    expect(svg.match(/<path d="M[^"]* A/g)?.length).toBe(3)
    expect(svg).toContain('>-10<')
  })

  const SCATTER = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:xVal>${numRef('Data!$B$2:$B$4', [1, 2, 3])}</c:xVal><c:yVal>${numRef('Data!$C$2:$C$4', [1, 2, 3])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`
  /** the x position of each value label in a rendered chart, by its text */
  const labelX = (svg: string) => Object.fromEntries([...svg.matchAll(/<text x="([\d.]+)"[^>]*>(\d+)</g)].map((m) => [m[2], Number(m[1])]))

  test('scatter: a point missing its X is not drawn; labels stay with their points', async () => {
    const { model } = await modelOf(await withChart(SCATTER))
    const s = model!.groups[0].series[0]
    s.labels = { value: true, percent: false, category: false, series: false }
    // values no axis shows as a tick label
    s.xs = [1, null, 3]
    s.values = [517, 733, 951]
    let at = labelX(renderChart(model!, 400, 260))
    expect('733' in at).toBe(false)
    expect(at['517'] < at['951']).toBe(true)
    // a missing Y: the next points keep their own labels
    s.xs = [1, 2, 3]
    s.values = [null, 733, 951]
    at = labelX(renderChart(model!, 400, 260))
    expect(at['733'] < at['951']).toBe(true)
  })

  test('series stop at the plot area when the axis range leaves values outside it', async () => {
    const { model } = await modelOf(await withChart(CLUSTERED))
    const val = model!.axes.find((a) => a.kind === 'val')!
    val.max = 5
    const svg = renderChart(model!, 400, 260)
    expect(svg).toContain('<clipPath id="de-plot">')
    expect(svg).toMatch(/<g clip-path="url\(#de-plot\)">[\s\S]*<\/g>/)
  })
})
