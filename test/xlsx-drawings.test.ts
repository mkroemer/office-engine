import { expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { rowsToXlsx } from '../src/core/convert'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'
import type { SheetDrawing } from '../src/core/xlsx/drawings'
import { drawingBox, drawingExtent } from '../src/react/xlsx/drawings'
import { Axis } from '../src/react/xlsx/axis'

const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const XDR = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const EMU = 9525

const corner = (tag: string, c: number, r: number, dx = 0, dy = 0) =>
  `<xdr:${tag}><xdr:col>${c}</xdr:col><xdr:colOff>${dx * EMU}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${dy * EMU}</xdr:rowOff></xdr:${tag}>`
const pic = (rid: string, descr: string) =>
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="${descr}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/></xdr:blipFill><xdr:spPr/></xdr:pic>`

/** a workbook whose sheet has a picture between two cells, a column chart by one cell and a size, an EMF picture at a position, and a text box */
async function withDrawings(): Promise<Blob> {
  const base = await rowsToXlsx([['a', 'b'], ['c', 'd']], 'Data')
  const zip = await ZipArchive.open(new BlobSource(base))
  const files = new Map<string, string | Uint8Array>()
  files.set('xl/worksheets/_rels/sheet1.xml.rels', `${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
  files.set(
    'xl/drawings/_rels/drawing1.xml.rels',
    `${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${R}/image" Target="../media/image1.png"/><Relationship Id="rId2" Type="${R}/chart" Target="../charts/chart1.xml"/><Relationship Id="rId3" Type="${R}/image" Target="../media/image2.emf"/></Relationships>`,
  )
  files.set(
    'xl/drawings/drawing1.xml',
    `${HEAD}<xdr:wsDr ${XDR}>` +
      `<xdr:twoCellAnchor>${corner('from', 1, 2, 10, 5)}${corner('to', 4, 8)}${pic('rId1', 'Logo')}<xdr:clientData/></xdr:twoCellAnchor>` +
      `<xdr:oneCellAnchor>${corner('from', 6, 1)}<xdr:ext cx="${300 * EMU}" cy="${200 * EMU}"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>` +
      `<xdr:absoluteAnchor><xdr:pos x="${50 * EMU}" y="${600 * EMU}"/><xdr:ext cx="${80 * EMU}" cy="${40 * EMU}"/>${pic('rId3', '')}<xdr:clientData/></xdr:absoluteAnchor>` +
      `<xdr:twoCellAnchor>${corner('from', 0, 0)}${corner('to', 1, 1)}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="5" name="TextBox 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>` +
      `</xdr:wsDr>`,
  )
  files.set(
    'xl/charts/chart1.xml',
    `${HEAD}<c:chartSpace ${XDR}><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales </a:t></a:r><a:r><a:t>2026</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
  )
  files.set('xl/media/image1.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47]))
  files.set('xl/media/image2.emf', new Uint8Array([1, 0, 0, 0]))
  const ct = (await zip.text('[Content_Types].xml')).replace(
    '</Types>',
    '<Default Extension="png" ContentType="image/png"/><Default Extension="emf" ContentType="image/x-emf"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
  )
  files.set('[Content_Types].xml', ct)
  return rebuildZip(zip, files as Map<string, string>, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}

test('a sheet lists its pictures and charts with their anchors, leaving shapes out', async () => {
  const engine = new Engine(() => {})
  await engine.open(await withDrawings(), {})
  const list = (await engine.call('xlsx.drawings', [0])) as SheetDrawing[]
  expect(list.length).toBe(3)
  const [logo, chart, emf] = list
  expect(logo).toMatchObject({ kind: 'picture', title: 'Logo', target: 'xl/media/image1.png', from: { c: 1, r: 2, dx: 10, dy: 5 }, to: { c: 4, r: 8, dx: 0, dy: 0 } })
  expect(chart).toMatchObject({ kind: 'chart', chart: 'column', title: 'Sales 2026', from: { c: 6, r: 1 }, size: { w: 300, h: 200 } })
  // a format the browser cannot draw keeps its place but no picture
  expect(emf).toMatchObject({ kind: 'picture', at: { x: 50, y: 600 }, size: { w: 80, h: 40 } })
  expect(emf.target).toBeUndefined()
  expect(await engine.call('xlsx.drawings', [5])).toEqual([])
})

test('a drawing is placed by its cells, so it moves and grows with them', () => {
  const cols = new Axis(64)
  const rows = new Axis(20)
  cols.count = rows.count = 50
  const d: SheetDrawing = { kind: 'picture', title: '', from: { c: 1, r: 2, dx: 10, dy: 5 }, to: { c: 4, r: 8, dx: 0, dy: 0 } }
  expect(drawingBox(d, cols, rows)).toEqual({ x: 74, y: 45, w: 182, h: 115 })
  cols.set(2, 100)
  rows.set(3, 0)
  expect(drawingBox(d, cols, rows)).toEqual({ x: 74, y: 45, w: 218, h: 95 })
  // by one cell and a size: the size stays
  expect(drawingBox({ kind: 'chart', title: '', from: { c: 6, r: 1, dx: 0, dy: 0 }, size: { w: 300, h: 200 } }, cols, rows)).toMatchObject({ w: 300, h: 200 })
  // the sheet reaches past its last drawing
  expect(drawingExtent([d, { kind: 'chart', title: '', at: { x: 0, y: 600 }, size: { w: 80, h: 40 } }], new Axis(64), new Axis(20))).toEqual({ rows: 33, cols: 5 })
})
