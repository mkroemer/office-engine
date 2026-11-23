import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip } from '../src/core/zipwrite'
import { useBrowser, viewer } from './harness'

useBrowser()

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

/** the sample document with a column chart in its first paragraph, its values saved with it */
async function chartDocx(): Promise<Buffer> {
  const zip = await ZipArchive.open(new BlobSource(new Blob([readFileSync(`${import.meta.dir}/../playground/public/samples/sample.docx`)])))
  const pts = (xs: (string | number)[]) => `<c:ptCount val="${xs.length}"/>${xs.map((x, i) => `<c:pt idx="${i}"><c:v>${x}</c:v></c:pt>`).join('')}`
  const chart =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:plotArea><c:layout/>` +
    `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>${pts(['Sales'])}</c:strCache></c:strRef></c:tx>` +
    `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache>${pts(['Q1', 'Q2', 'Q3'])}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache>${pts([3, 7, 5])}</c:numCache></c:numRef></c:val></c:ser>` +
    `<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`
  const inline =
    `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="77" name="Chart 1"/>` +
    `<a:graphic xmlns:a="${A}"><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
  const files = new Map([
    ['word/document.xml', (await zip.text('word/document.xml')).replace(/<w:body>/, `<w:body>${inline}`)],
    ['word/_rels/document.xml.rels', (await zip.text('word/_rels/document.xml.rels')).replace('</Relationships>', `<Relationship Id="rIdChart1" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`)],
    ['word/charts/chart1.xml', chart],
    ['[Content_Types].xml', (await zip.text('[Content_Types].xml')).replace('</Types>', '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>')],
  ])
  return Buffer.from(await (await rebuildZip(zip, files, DOCX)).arrayBuffer())
}

/** a picture has loaded: an SVG picture has its size once it is read */
const loaded = (img: HTMLImageElement) => img.complete && img.naturalWidth > 0

describe('charts drawn in the views', () => {
  test('a chart on a slide is a picture at its frame', async () => {
    const v = await viewer({ src: 'samples/sample.pptx' })
    // the fifth slide holds the chart: its thumbnail brings it into view
    await v.page.locator('.de-thumb').nth(4).click()
    const frame = v.page.locator('.de-slide-frame[data-slide="5"]')
    const img = frame.locator('[data-top] img[src^="blob:"]').first()
    await img.waitFor()
    await v.page.waitForFunction((el) => (el as HTMLImageElement).complete, await img.elementHandle())
    expect(await img.evaluate(loaded)).toBe(true)
    expect(await frame.locator('.de-ph', { hasText: 'Chart' }).count()).toBe(0)
  })

  test('a chart in a Word document is a picture in its line', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    await v.chooseFile(() => v.menu('Open…'), { name: 'chart.docx', mimeType: DOCX, buffer: await chartDocx() })
    const img = v.page.locator('.de-page-inner img[src^="blob:"]').first()
    await img.waitFor()
    await v.page.waitForFunction((el) => (el as HTMLImageElement).complete, await img.elementHandle())
    expect(await img.evaluate(loaded)).toBe(true)
    // drawn at the inline's size: 4572000 × 2743200 EMU is 360 × 216 pt
    const box = await img.evaluate((el) => (el as HTMLElement).style.width)
    expect(box).toBe('360pt')
  })
})
