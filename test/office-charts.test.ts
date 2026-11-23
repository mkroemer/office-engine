import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import type { Block, Inline } from '../src/core/docx/types'
import type { SlideModel } from '../src/core/pptx/types'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'
import { renderChart } from '../src/core/xlsx/chart-svg'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

/** a column chart part whose values are saved with it (a workbook embedded in the file holds their source) */
const chartPart = (values: number[], view = '') =>
  `${HEAD}<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:autoTitleDeleted val="0"/>${view}<c:plotArea><c:layout/>` +
  `<c:${view ? 'bar3DChart' : 'barChart'}><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${values.length + 1}</c:f><c:strCache><c:ptCount val="${values.length}"/>${values.map((_, i) => `<c:pt idx="${i}"><c:v>Q${i + 1}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$${values.length + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val></c:ser>` +
  `<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:${view ? 'bar3DChart' : 'barChart'}><c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>` +
  `</c:plotArea><c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`

/** a line chart of two series that say they have 100 000 points each and give none: a few hundred bytes of file */
const CLAIMED =
  `${HEAD}<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/>` +
  [0, 1].map((i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:val><c:numLit><c:ptCount val="100000"/></c:numLit></c:val></c:ser>`).join('') +
  `<c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`

const graphic = (rid: string) => `<a:graphic xmlns:a="${A}"><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="${rid}"/></a:graphicData></a:graphic>`

/** the sample with files replaced or added, each edited from its old text (absent: '') */
async function variant(name: string, edits: Record<string, (xml: string) => string>): Promise<Blob> {
  const zip = await ZipArchive.open(new BlobSource(new Blob([await Bun.file(S(name)).arrayBuffer()])))
  const files = new Map<string, string>()
  for (const [path, fn] of Object.entries(edits)) files.set(path, fn(zip.has(path) ? await zip.text(path) : ''))
  return rebuildZip(zip, files)
}

const override = (part: string) => (xml: string) => xml.replace('</Types>', `<Override PartName="/${part}" ContentType="${CHART_TYPE}"/></Types>`)

describe('charts in Word documents', () => {
  async function blocksOf(src: Blob) {
    const events: { event: string; data: any }[] = []
    const engine = new Engine((event, data) => events.push({ event, data }))
    await engine.open(src, {})
    await engine.call('docx.start', [])
    return events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data) as Block[]
  }
  const inlines = (blocks: Block[]): Inline[] => blocks.flatMap((b) => (b.t === 'p' ? b.runs : []))
  const docx = (body: string, chart?: string) =>
    variant('sample.docx', {
      'word/document.xml': (xml) => xml.replace(/<w:body>/, `<w:body>${body}`),
      'word/_rels/document.xml.rels': (xml) => xml.replace('</Relationships>', `<Relationship Id="rIdChart1" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`),
      ...(chart ? { 'word/charts/chart1.xml': () => chart, '[Content_Types].xml': override('word/charts/chart1.xml') } : {}),
    })
  const inline = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="77" name="Chart 1"/>${graphic('rIdChart1')}</wp:inline></w:drawing></w:r></w:p>`

  test('a chart is read from the values saved with it and drawn in its frame', async () => {
    const blocks = await blocksOf(await docx(inline, chartPart([3, 7, 5])))
    const shape = inlines(blocks).find((x) => x.t === 'shape' && x.label === 'Chart') as Extract<Inline, { t: 'shape' }>
    expect(shape).toBeDefined()
    expect(shape.w).toBe(360)
    expect(shape.chart?.groups[0].series[0]).toMatchObject({ name: 'Sales', cats: ['Q1', 'Q2', 'Q3'], values: [3, 7, 5] })
    // the model crosses from the worker as it is
    expect(structuredClone(shape.chart)).toEqual(shape.chart!)
    const svg = renderChart(shape.chart!, 480, 288)
    expect(svg).toContain('>Q2<')
    expect(svg).not.toContain('NaN')
  })

  test('a chart part shown in many places: the places share a budget of points, those past it keep their label', async () => {
    // 400 000 points a place (two series that claim 100 000 each, and their categories): two places fit a million
    const blocks = await blocksOf(await docx(inline.repeat(5), CLAIMED))
    const shapes = inlines(blocks).filter((x) => x.t === 'shape' && x.label === 'Chart') as Extract<Inline, { t: 'shape' }>[]
    expect(shapes.map((x) => !!x.chart)).toEqual([true, true, false, false, false])
    // an ordinary chart is shown in every place
    const small = inlines(await blocksOf(await docx(inline.repeat(5), chartPart([3, 7, 5])))).filter((x) => x.t === 'shape' && x.label === 'Chart') as Extract<Inline, { t: 'shape' }>[]
    expect(small.map((x) => !!x.chart)).toEqual([true, true, true, true, true])
  })

  test('a 3D chart keeps its view; a chart part that is missing leaves its labeled box', async () => {
    const three = await blocksOf(await docx(inline, chartPart([3, 7, 5], '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/></c:view3D>')))
    const shape = inlines(three).find((x) => x.t === 'shape' && x.label === 'Chart') as Extract<Inline, { t: 'shape' }>
    expect(shape.chart?.view3d).toMatchObject({ rotX: 15, rotY: 20, rightAngle: true })
    const missing = await blocksOf(await docx(inline))
    const box = inlines(missing).find((x) => x.t === 'shape' && x.label === 'Chart') as Extract<Inline, { t: 'shape' }>
    expect(box).toBeDefined()
    expect(box.chart).toBeUndefined()
  })
})

describe('charts on slides', () => {
  const frame = (rid: string, id: number) =>
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="2743200"/></p:xfrm>${graphic(rid)}</p:graphicFrame>`
  const pptx = (frames: string, chart?: string) =>
    variant('sample.pptx', {
      'ppt/slides/slide1.xml': (xml) => xml.replace('</p:spTree>', `${frames}</p:spTree>`),
      'ppt/slides/_rels/slide1.xml.rels': (xml) => xml.replace('</Relationships>', `<Relationship Id="rIdChart1" Type="${R}/chart" Target="../charts/chart1.xml"/><Relationship Id="rIdGone" Type="${R}/chart" Target="../charts/chart9.xml"/></Relationships>`),
      ...(chart ? { 'ppt/charts/chart1.xml': () => chart, '[Content_Types].xml': override('ppt/charts/chart1.xml') } : {}),
    })

  test('a chart takes the place of its frame, at its place on the slide; one whose part is missing stays a placeholder', async () => {
    const engine = new Engine(() => {})
    await engine.open(await pptx(frame('rIdChart1', 901) + frame('rIdGone', 902), chartPart([10, 20, 15, 30])), {})
    const slide = (await engine.call('pptx.slide', [0])) as SlideModel
    const chart = slide.els.find((e) => e.t === 'chart') as Extract<SlideModel['els'][number], { t: 'chart' }>
    expect(chart).toBeDefined()
    expect(chart.box).toMatchObject({ x: 96, y: 96, w: 480, h: 288 })
    expect(chart.top).toBe(901)
    expect(chart.model.groups[0].series[0].values).toEqual([10, 20, 15, 30])
    expect(renderChart(chart.model, chart.box.w, chart.box.h)).toContain('>Q4<')
    expect(slide.els.some((e) => e.t === 'placeholder' && e.label === 'Chart' && e.top === 902)).toBe(true)
  })

  test('a chart part shown by many frames: they share a budget of points, and a slide read again keeps its charts', async () => {
    const engine = new Engine(() => {})
    await engine.open(await pptx([911, 912, 913, 914].map((id) => frame('rIdChart1', id)).join(''), CLAIMED), {})
    const shown = async () => ((await engine.call('pptx.slide', [0])) as SlideModel).els.filter((e) => e.top !== undefined && e.top >= 911).map((e) => [e.top, e.t])
    // 400 000 points a frame: two fit a million
    const first = await shown()
    expect(first).toEqual([[911, 'chart'], [912, 'chart'], [913, 'placeholder'], [914, 'placeholder']])
    // read again (after an edit, for a thumbnail): its own share is not counted twice
    expect(await shown()).toEqual(first)
    expect(await shown()).toEqual(first)
    // the slides share the budget: the sample's fifth slide shows the same part (its own chart's), past what is left;
    // read at once with the first, neither takes the other's points
    const all = (await Promise.all([4, 0, 4, 0].map((i) => engine.call('pptx.slide', [i])))) as SlideModel[]
    for (const k of [1, 3]) expect(all[k].els.filter((e) => e.t === 'chart').length).toBe(2)
    for (const k of [0, 2]) {
      expect(all[k].els.filter((e) => e.t === 'chart').length).toBe(0)
      expect(all[k].els.some((e) => e.t === 'placeholder' && e.label === 'Chart')).toBe(true)
    }
  })
})
