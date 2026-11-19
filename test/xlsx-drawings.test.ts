import { expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { rowsToXlsx } from '../src/core/convert'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'
import type { SheetDrawing } from '../src/core/xlsx/drawings'
import { drawingBox, drawingExtent } from '../src/react/xlsx/drawings'
import { Axis } from '../src/react/xlsx/axis'
import { MAX_COL, MAX_ROWS, MAX_SHEET_HEIGHT_PX, MAX_SHEET_WIDTH_PX } from '../src/core/xlsx/limits'

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
async function withDrawings(anchors?: string): Promise<Blob> {
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
      (anchors ?? (`<xdr:twoCellAnchor>${corner('from', 1, 2, 10, 5)}${corner('to', 4, 8)}${pic('rId1', 'Logo')}<xdr:clientData/></xdr:twoCellAnchor>` +
      `<xdr:oneCellAnchor>${corner('from', 6, 1)}<xdr:ext cx="${300 * EMU}" cy="${200 * EMU}"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>` +
      `<xdr:absoluteAnchor><xdr:pos x="${50 * EMU}" y="${600 * EMU}"/><xdr:ext cx="${80 * EMU}" cy="${40 * EMU}"/>${pic('rId3', '')}<xdr:clientData/></xdr:absoluteAnchor>` +
      `<xdr:twoCellAnchor>${corner('from', 0, 0)}${corner('to', 1, 1)}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="5" name="TextBox 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`)) +
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

test('an anchor corner without its offsets is at the cell\'s own corner', async () => {
  const bare = (name: string, c: number, r: number) => `<xdr:${name}><xdr:col>${c}</xdr:col><xdr:row>${r}</xdr:row></xdr:${name}>`
  const engine = new Engine(() => {})
  await engine.open(await withDrawings(`<xdr:twoCellAnchor>${bare('from', 1, 1)}${bare('to', 3, 6)}${pic('rId1', 'loose')}</xdr:twoCellAnchor>`), {})
  const list = await engine.call('xlsx.drawings', [0]) as SheetDrawing[]
  expect(list).toHaveLength(1)
  expect(list[0].from).toEqual({ c: 1, r: 1, dx: 0, dy: 0 })
  expect(list[0].to).toEqual({ c: 3, r: 6, dx: 0, dy: 0 })
})

test('invalid drawing coordinates discard the anchor without stretching the sheet', async () => {
  const two = (from: string, to = corner('to', 2, 3)) => `<xdr:twoCellAnchor>${from}${to}${pic('rId1', 'bad')}</xdr:twoCellAnchor>`
  const invalid: string[] = []
  for (const r of [-1, 0.5, MAX_ROWS, 3_999_999_998, NaN, Infinity]) invalid.push(two(corner('from', 0, r)))
  for (const c of [-1, 0.5, MAX_COL, 4_000_000_000, NaN, Infinity]) invalid.push(two(corner('from', c, 0)))
  invalid.push(two(corner('from', 0, 0), corner('to', 0, 4_000_000_000)))
  for (const offset of [NaN, Infinity, Number.MAX_SAFE_INTEGER, MAX_SHEET_HEIGHT_PX + 1]) {
    invalid.push(two(corner('from', 0, 0, 0, offset)))
  }
  invalid.push(two(corner('from', 0, 0).replace('<xdr:row>0</xdr:row>', '')))
  invalid.push(two(corner('from', 0, 0).replace('<xdr:colOff>0</xdr:colOff>', '<xdr:colOff>0.5</xdr:colOff>')))
  invalid.push(two(corner('from', 0, 0), ''))
  for (const [cx, cy] of [['NaN', '9525'], ['9525', 'Infinity'], ['-1', '9525'], ['0.5', '9525'], ['9007199254740992', '9525'], [String((MAX_SHEET_WIDTH_PX + 1) * EMU), '9525']]) {
    invalid.push(`<xdr:oneCellAnchor>${corner('from', 0, 0)}<xdr:ext cx="${cx}" cy="${cy}"/>${pic('rId1', 'bad')}</xdr:oneCellAnchor>`)
  }
  for (const x of ['NaN', 'Infinity', '0.5', '9007199254740992', String((MAX_SHEET_WIDTH_PX + 1) * EMU)]) {
    invalid.push(`<xdr:absoluteAnchor><xdr:pos x="${x}" y="0"/><xdr:ext cx="9525" cy="9525"/>${pic('rId1', 'bad')}</xdr:absoluteAnchor>`)
  }
  const valid = two(corner('from', MAX_COL - 2, MAX_ROWS - 2, -1, 1 / EMU), corner('to', MAX_COL - 1, MAX_ROWS - 1))
  const engine = new Engine(() => {})
  await engine.open(await withDrawings(invalid.join('') + valid), {})
  const list = await engine.call('xlsx.drawings', [0]) as SheetDrawing[]
  expect(list).toHaveLength(1)
  expect(list[0].from).toEqual({ c: MAX_COL - 2, r: MAX_ROWS - 2, dx: -1, dy: 1 / EMU })
  expect(drawingExtent(list, new Axis(64), new Axis(20))).toEqual({ rows: MAX_ROWS, cols: MAX_COL })
})

test('valid large pixel anchors contribute at most an OOXML worksheet to view extents', async () => {
  const size = `<xdr:ext cx="${MAX_SHEET_WIDTH_PX * EMU}" cy="${MAX_SHEET_HEIGHT_PX * EMU}"/>`
  const engine = new Engine(() => {})
  await engine.open(await withDrawings(
    `<xdr:oneCellAnchor>${corner('from', MAX_COL - 1, MAX_ROWS - 1)}${size}${pic('rId1', 'large cell anchor')}</xdr:oneCellAnchor>` +
    `<xdr:absoluteAnchor><xdr:pos x="${MAX_SHEET_WIDTH_PX * EMU}" y="${MAX_SHEET_HEIGHT_PX * EMU}"/>${size}${pic('rId1', 'large absolute anchor')}</xdr:absoluteAnchor>`,
  ), {})
  const list = await engine.call('xlsx.drawings', [0]) as SheetDrawing[]
  expect(list).toHaveLength(2)
  for (const d of list) {
    expect(d.size).toEqual({ w: MAX_SHEET_WIDTH_PX, h: MAX_SHEET_HEIGHT_PX })
    expect(drawingExtent([d], new Axis(64), new Axis(20))).toEqual({ rows: MAX_ROWS, cols: MAX_COL })
  }
})
