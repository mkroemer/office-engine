/**
 * Builds Visio drawings (.vsdx) part by part, the way Visio writes them: a document with its style sheets, masters,
 * pages of shapes (cells, Geometry / Character / Paragraph sections, text) and a theme. For the tests and the
 * playground sample (scripts/visio-samples.ts). Lengths are inches, angles radians, as in the file.
 */
import { writeZip, type PartData } from '../src/core/zipwrite'
import { png } from './iwork-builder'

type Val = string | number

const NS = "xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xml:space='preserve'"
const HEAD = "<?xml version='1.0' encoding='utf-8' ?>\n"
const REL = 'http://schemas.microsoft.com/visio/2010/relationships'

export const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&apos;')

/** `<Cell>` elements; a value of the form `[value, formula]` writes the formula too */
export const cells = (c: Record<string, Val | [Val, string]> = {}) =>
  Object.entries(c)
    .map(([n, v]) => (Array.isArray(v) ? `<Cell N='${n}' V='${esc(String(v[0]))}' F='${esc(v[1])}'/>` : `<Cell N='${n}' V='${esc(String(v))}'/>`))
    .join('')

export type GeomRow = [type: string, cells?: Record<string, Val>] | { del: true }

/** a Geometry section: rows numbered from 1; `{ del: true }` deletes the row a master gives at that place */
export function geom(ix: number, rows: GeomRow[], flags: Record<string, Val> = {}): string {
  const body = rows.map((r, i) => (Array.isArray(r) ? `<Row T='${r[0]}' IX='${i + 1}'>${cells(r[1])}</Row>` : `<Row T='LineTo' IX='${i + 1}' Del='1'/>`)).join('')
  return `<Section N='Geometry' IX='${ix}'>${cells(flags)}${body}</Section>`
}

/** a section of numbered rows (Character, Paragraph, FillGradient, Layer…) */
export const rows = (name: string, list: Record<string, Val>[]) => `<Section N='${name}'>${list.map((c, i) => `<Row IX='${i}'>${cells(c)}</Row>`).join('')}</Section>`

export const rect = (w: number, h: number, ix = 0, flags: Record<string, Val> = {}) =>
  geom(ix, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: w, Y: 0 }], ['LineTo', { X: w, Y: h }], ['LineTo', { X: 0, Y: h }], ['LineTo', { X: 0, Y: 0 }]], flags)

/** a rectangle in fractions of the shape's box: it follows the shape's size */
export const relRect = (ix = 0, flags: Record<string, Val> = {}) =>
  geom(ix, [['RelMoveTo', { X: 0, Y: 0 }], ['RelLineTo', { X: 1, Y: 0 }], ['RelLineTo', { X: 1, Y: 1 }], ['RelLineTo', { X: 0, Y: 1 }], ['RelLineTo', { X: 0, Y: 0 }]], flags)

export interface ShapeSpec {
  id: number
  type?: 'Shape' | 'Group' | 'Foreign' | 'Guide'
  name?: string
  /** the master this shape is an instance of; `masterShape`: the master's shape a member of an instance follows */
  master?: number
  masterShape?: number
  /** the style (line, fill and text) */
  style?: number
  /** PinX, PinY, Width, Height; the local pin is the center */
  box?: [number, number, number, number]
  cells?: Record<string, Val | [Val, string]>
  /** sections, as XML */
  sections?: string
  /** the inner XML of `<Text>` */
  text?: string
  shapes?: string[]
  /** relationship id of the picture of a Foreign shape */
  foreign?: string
}

export function shape(s: ShapeSpec): string {
  let attrs = ` ID='${s.id}'`
  if (s.name) attrs += ` NameU='${esc(s.name)}' Name='${esc(s.name)}'`
  if (s.type || s.master === undefined) attrs += ` Type='${s.type ?? 'Shape'}'`
  if (s.master !== undefined) attrs += ` Master='${s.master}'`
  if (s.masterShape !== undefined) attrs += ` MasterShape='${s.masterShape}'`
  if (s.style !== undefined) attrs += ` LineStyle='${s.style}' FillStyle='${s.style}' TextStyle='${s.style}'`
  const box = s.box ? cells({ PinX: s.box[0], PinY: s.box[1], Width: s.box[2], Height: s.box[3], LocPinX: [s.box[2] / 2, 'Width*0.5'], LocPinY: [s.box[3] / 2, 'Height*0.5'] }) : ''
  return (
    `<Shape${attrs}>${box}${cells(s.cells)}${s.sections ?? ''}${s.text !== undefined ? `<Text>${s.text}</Text>` : ''}` +
    `${s.foreign ? `<ForeignData ForeignType='Bitmap' CompressionType='PNG'><Rel r:id='${s.foreign}'/></ForeignData>` : ''}` +
    `${s.shapes ? `<Shapes>${s.shapes.join('')}</Shapes>` : ''}</Shape>`
  )
}

/** the defaults every sheet falls back to (Visio's "No Style") */
const NO_STYLE: Record<string, Val> = {
  LineWeight: 0.01041666666666667, LineColor: 0, LinePattern: 1, Rounding: 0, EndArrowSize: 2, BeginArrow: 0, EndArrow: 0, LineCap: 0, BeginArrowSize: 2, LineColorTrans: 0,
  FillForegnd: 1, FillBkgnd: 0, FillPattern: 1, FillForegndTrans: 0, FillBkgndTrans: 0, LeftMargin: 0, RightMargin: 0, TopMargin: 0, BottomMargin: 0, VerticalAlign: 1,
  TextBkgnd: 0, TextBkgndTrans: 0, HideText: 0, DisplayMode: 2, QuickStyleLineColor: 100, QuickStyleFillColor: 100, QuickStyleFontColor: 100, QuickStyleLineMatrix: 100,
  QuickStyleFillMatrix: 100, QuickStyleFontMatrix: 100, QuickStyleVariation: 0,
}
const CHAR: Record<string, Val> = { Font: 'Calibri', Color: 0, Style: 0, Case: 0, Pos: 0, FontScale: 1, Size: 0.1666666666666667, DblUnderline: 0, Overline: 0, Strikethru: 0, ColorTrans: 0 }
const PARA: Record<string, Val> = { IndFirst: 0, IndLeft: 0, IndRight: 0, SpLine: -1.2, SpBefore: 0, SpAfter: 0, HorzAlign: 1, Bullet: 0, BulletStr: '' }
/** every cell a theme sets */
const THEMED = ['LineWeight', 'LineColor', 'LinePattern', 'Rounding', 'FillForegnd', 'FillBkgnd', 'FillPattern', 'FillGradientEnabled']

/** style ids of a new drawing: 0 the defaults, 3 "Normal" (the 4 pt text margins of Visio), 6 "Theme" (every look from the theme), 7 "Themed" (Normal over Theme) */
export const STYLE = { NONE: 0, NORMAL: 3, THEME: 6, THEMED: 7 }

const styleSheet = (id: number, name: string, parent: number | undefined, body: string) =>
  `<StyleSheet ID='${id}' NameU='${name}' Name='${name}'${parent === undefined ? '' : ` LineStyle='${parent}' FillStyle='${parent}' TextStyle='${parent}'`}>${body}</StyleSheet>`

export interface PageSpec {
  /** inches */
  w?: number
  h?: number
  background?: boolean
  /** id of the background page drawn under this one */
  backPage?: number
  /** more page cells (PageScale, DrawingScale, ThemeIndex…) */
  cells?: Record<string, Val>
  /** layers: their Visible cell */
  layers?: boolean[]
  /** media file names: relationship rId1, rId2… of the page */
  images?: string[]
}

/** A sample theme: the Office colors, a variation of 7 colors, three fill, line and font styles. */
export function themeXml(id = 33): string {
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  const VT = "xmlns:vt='http://schemas.microsoft.com/office/visio/2012/theme'"
  const c = (name: string, hex: string) => `<a:${name}><a:srgbClr val="${hex}"/></a:${name}>`
  const ph = (mod = '') => `<a:schemeClr val="phClr">${mod}</a:schemeClr>`
  const vars = ['5B9BD5', '759FCC', '41719C', 'BDD0E9', 'ED7D31', '70AD47', 'FEC000'].map((hex, i) => `<vt:varColor${i + 1}><a:srgbClr val="${hex}"/></vt:varColor${i + 1}>`).join('')
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="${A}" name="Office"><a:themeElements><a:clrScheme name="Office">` +
    c('dk1', '000000') + c('lt1', 'FEFFFF') + c('dk2', '44546A') + c('lt2', 'E7E6E6') + c('accent1', '5B9BD5') + c('accent2', 'ED7D31') + c('accent3', 'A5A5A5') +
    c('accent4', 'FFC000') + c('accent5', '4672C4') + c('accent6', '70AD47') + c('hlink', '0563C1') + c('folHlink', '954F72') +
    `<a:extLst><a:ext uri="{093E89EA-6996-430E-BFF9-83A9FAAAAB73}"><vt:bkgnd ${VT}><a:srgbClr val="FFFFFF"/></vt:bkgnd></a:ext>` +
    `<a:ext uri="{DDD2D869-C2EF-471E-B8FA-914AFA308C9F}"><vt:variationClrSchemeLst ${VT}><vt:variationClrScheme monotone="1">${vars}</vt:variationClrScheme></vt:variationClrSchemeLst></a:ext></a:extLst></a:clrScheme>` +
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme>' +
    `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:solidFill>${ph()}</a:solidFill>` +
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${ph('<a:tint val="20000"/>')}</a:gs><a:gs pos="100000">${ph()}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>` +
    `<a:lnStyleLst><a:ln w="9525"><a:solidFill>${ph()}</a:solidFill><a:prstDash val="solid"/></a:ln>` +
    `<a:ln w="3175"><a:solidFill><a:schemeClr val="lt1"><a:shade val="58000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>` +
    `<a:ln w="19050"><a:solidFill>${ph('<a:shade val="75000"/>')}</a:solidFill><a:prstDash val="dash"/></a:ln></a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>' +
    `<a:extLst><a:ext uri="{D75FF423-9257-4291-A4FE-1B2448832E17}"><vt:themeScheme ${VT}><vt:schemeID schemeEnum="${id}" schemeGUID="{00000000-0000-0000-0000-000000000000}"/></vt:themeScheme></a:ext>` +
    `<a:ext uri="{1342405F-259F-4C95-8CDF-A9DCE2D418A2}"><vt:fmtConnectorScheme ${VT} name="Office"><a:lnStyleLst><a:ln w="12700"><a:solidFill>${ph()}</a:solidFill><a:prstDash val="sysDash"/></a:ln></a:lnStyleLst></vt:fmtConnectorScheme></a:ext>` +
    `<a:ext uri="{6CAB99AB-0A78-4BAB-B597-526D62367CB4}"><vt:lineStyles ${VT}><vt:fmtConnectorSchemeLineStyles><vt:lineStyle><vt:lineEx rndg="0" start="0" startSize="2" end="4" endSize="2" pattern="2"/></vt:lineStyle></vt:fmtConnectorSchemeLineStyles>` +
    '<vt:fmtSchemeLineStyles><vt:lineStyle><vt:lineEx rndg="0" start="0" startSize="2" end="0" endSize="2" pattern="1"/></vt:lineStyle><vt:lineStyle><vt:lineEx rndg="0" start="0" startSize="2" end="0" endSize="2" pattern="1"/></vt:lineStyle>' +
    '<vt:lineStyle><vt:lineEx rndg="0" start="0" startSize="2" end="0" endSize="2" pattern="2"/></vt:lineStyle></vt:fmtSchemeLineStyles></vt:lineStyles></a:ext>' +
    `<a:ext uri="{EBE24D50-EC5C-4D6F-A1A3-C5F0A18B936A}"><vt:fontStylesGroup ${VT}><vt:connectorFontStyles><vt:fontProps style="0"><vt:color>${ph('<a:shade val="75000"/>')}</vt:color></vt:fontProps></vt:connectorFontStyles>` +
    `<vt:fontStyles><vt:fontProps style="0"><vt:color>${ph()}</vt:color></vt:fontProps><vt:fontProps style="1"><vt:color><a:schemeClr val="lt1"/></vt:color></vt:fontProps>` +
    '<vt:fontProps style="2"><vt:color><a:schemeClr val="dk1"/></vt:color></vt:fontProps></vt:fontStyles></vt:fontStylesGroup></a:ext>' +
    `<a:ext uri="{494CE47F-D151-47DC-95E8-85652EA8A67E}"><vt:variationStyleSchemeLst ${VT}><vt:variationStyleScheme embellishment="1"><vt:varStyle fillIdx="2" lineIdx="2" effectIdx="1" fontIdx="2"/>` +
    '<vt:varStyle fillIdx="1" lineIdx="1" effectIdx="1" fontIdx="1"/><vt:varStyle fillIdx="3" lineIdx="3" effectIdx="1" fontIdx="3"/><vt:varStyle fillIdx="2" lineIdx="1" effectIdx="1" fontIdx="2"/>' +
    '</vt:variationStyleScheme></vt:variationStyleSchemeLst></a:ext></a:extLst></a:themeElements></a:theme>'
  )
}

export class VisioBuilder {
  /** extra style sheets, as XML */
  styles = ''
  /** `<ColorEntry>` list: the document's palette */
  colors = ''
  theme: string | undefined
  private masters: { id: number; name: string; shapes: string; w: number; h: number; images: string[] }[] = []
  private pages: { id: number; name: string; shapes: string; spec: PageSpec }[] = []
  private media = new Map<string, Uint8Array>()

  /** a master of one or more shapes; returns its id */
  master(name: string, shapes: string, opts: { w?: number; h?: number; images?: string[] } = {}): number {
    const id = this.masters.length + 1
    this.masters.push({ id, name, shapes, w: opts.w ?? 2, h: opts.h ?? 2, images: opts.images ?? [] })
    return id
  }

  /** a page of shapes (back to front); returns its id */
  page(name: string, shapes: string, spec: PageSpec = {}): number {
    const id = this.pages.length * 4
    this.pages.push({ id, name, shapes, spec })
    return id
  }

  /** a picture of the package, named in a page's or master's `images` */
  image(name: string, bytes: Uint8Array): string {
    this.media.set(name, bytes)
    return name
  }

  async zip(ext: 'drawing' | 'stencil' | 'template' = 'drawing'): Promise<Uint8Array<ArrayBuffer>> {
    const files = new Map<string, PartData>()
    const rels = (list: [string, string][]) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`
    const imageRels = (images: string[]): [string, string][] => images.map((n) => ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', `../media/${n}`])
    const pageSheet = (w: number, h: number, extra: Record<string, Val> = {}, layers: boolean[] = []) =>
      `<PageSheet LineStyle='0' FillStyle='0' TextStyle='0'>${cells({ PageWidth: w, PageHeight: h, PageScale: 1, DrawingScale: 1, ...extra })}` +
      `${layers.length ? rows('Layer', layers.map((v, i) => ({ Name: `Layer ${i + 1}`, Visible: v ? 1 : 0, Print: 1 }))) : ''}</PageSheet>`

    const themed = cells(Object.fromEntries(THEMED.map((n) => [n, ['Themed', 'THEMEVAL()'] as [Val, string]])))
    files.set(
      'visio/document.xml',
      `${HEAD}<VisioDocument ${NS}><DocumentSettings TopPage='0' DefaultTextStyle='3' DefaultLineStyle='3' DefaultFillStyle='3'/>` +
        `<Colors>${this.colors}</Colors><FaceNames><FaceName NameU='Calibri'/></FaceNames><StyleSheets>` +
        styleSheet(STYLE.NONE, 'No Style', undefined, cells(NO_STYLE) + rows('Character', [CHAR]) + rows('Paragraph', [PARA])) +
        styleSheet(STYLE.NORMAL, 'Normal', STYLE.NONE, cells({ LeftMargin: 0.05555555555555555, RightMargin: 0.05555555555555555, TopMargin: 0.05555555555555555, BottomMargin: 0.05555555555555555 })) +
        styleSheet(STYLE.THEME, 'Theme', STYLE.NONE, themed + rows('Character', [{ Font: 'Themed', Color: 'Themed', Style: 'Themed' }])) +
        styleSheet(STYLE.THEMED, 'Themed', STYLE.THEME, cells({ LeftMargin: 0.05555555555555555, RightMargin: 0.05555555555555555 })) +
        `${this.styles}</StyleSheets></VisioDocument>`,
    )
    const docRels: [string, string][] = [
      [`${REL}/masters`, 'masters/masters.xml'],
      [`${REL}/pages`, 'pages/pages.xml'],
    ]
    if (this.theme) {
      docRels.push(['http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml'])
      files.set('visio/theme/theme1.xml', this.theme)
    }
    files.set('visio/_rels/document.xml.rels', rels(docRels))

    files.set(
      'visio/masters/masters.xml',
      `${HEAD}<Masters ${NS}>${this.masters.map((m, i) => `<Master ID='${m.id}' NameU='${esc(m.name)}' Name='${esc(m.name)}'>${pageSheet(m.w, m.h)}<Rel r:id='rId${i + 1}'/></Master>`).join('')}</Masters>`,
    )
    files.set('visio/masters/_rels/masters.xml.rels', rels(this.masters.map((m) => [`${REL}/master`, `master${m.id}.xml`])))
    for (const m of this.masters) {
      files.set(`visio/masters/master${m.id}.xml`, `${HEAD}<MasterContents ${NS}><Shapes>${m.shapes}</Shapes></MasterContents>`)
      if (m.images.length) files.set(`visio/masters/_rels/master${m.id}.xml.rels`, rels(imageRels(m.images)))
    }

    files.set(
      'visio/pages/pages.xml',
      `${HEAD}<Pages ${NS}>${this.pages
        .map((p, i) => {
          const attrs = (p.spec.background ? " Background='1'" : '') + (p.spec.backPage !== undefined ? ` BackPage='${p.spec.backPage}'` : '')
          return `<Page ID='${p.id}' NameU='${esc(p.name)}' Name='${esc(p.name)}'${attrs}>${pageSheet(p.spec.w ?? 11, p.spec.h ?? 8.5, p.spec.cells, p.spec.layers)}<Rel r:id='rId${i + 1}'/></Page>`
        })
        .join('')}</Pages>`,
    )
    files.set('visio/pages/_rels/pages.xml.rels', rels(this.pages.map((_, i) => [`${REL}/page`, `page${i + 1}.xml`])))
    this.pages.forEach((p, i) => {
      files.set(`visio/pages/page${i + 1}.xml`, `${HEAD}<PageContents ${NS}><Shapes>${p.shapes}</Shapes></PageContents>`)
      if (p.spec.images?.length) files.set(`visio/pages/_rels/page${i + 1}.xml.rels`, rels(imageRels(p.spec.images)))
    })
    for (const [name, bytes] of this.media) files.set(`visio/media/${name}`, bytes)

    const main = `application/vnd.ms-visio.${ext}.main+xml`
    files.set(
      '[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        `<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/visio/document.xml" ContentType="${main}"/>` +
        '<Override PartName="/visio/masters/masters.xml" ContentType="application/vnd.ms-visio.masters+xml"/><Override PartName="/visio/pages/pages.xml" ContentType="application/vnd.ms-visio.pages+xml"/>' +
        this.masters.map((m) => `<Override PartName="/visio/masters/master${m.id}.xml" ContentType="application/vnd.ms-visio.master+xml"/>`).join('') +
        this.pages.map((_, i) => `<Override PartName="/visio/pages/page${i + 1}.xml" ContentType="application/vnd.ms-visio.page+xml"/>`).join('') +
        (this.theme ? '<Override PartName="/visio/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' : '') +
        '</Types>',
    )
    files.set('_rels/.rels', rels([[`${REL}/document`, 'visio/document.xml']]))
    const blob = await writeZip(files, 'application/vnd.ms-visio.drawing')
    return new Uint8Array(await blob.arrayBuffer())
  }
}

/** a straight 1-D shape from (x1, y1) to (x2, y2): its box lies along the line, turned to its direction */
export function line(id: number, x1: number, y1: number, x2: number, y2: number, extra: Partial<ShapeSpec> = {}): string {
  const len = Math.hypot(x2 - x1, y2 - y1)
  const angle = Math.atan2(y2 - y1, x2 - x1)
  return shape({
    id,
    style: STYLE.NORMAL,
    ...extra,
    cells: {
      PinX: (x1 + x2) / 2, PinY: (y1 + y2) / 2, Width: len, Height: 0, LocPinX: len / 2, LocPinY: 0, Angle: angle, BeginX: x1, BeginY: y1, EndX: x2, EndY: y2,
      ...extra.cells,
    },
    sections: (extra.sections ?? '') + (extra.master === undefined ? geom(0, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: len, Y: 0 }]], { NoFill: 1 }) : geom(0, [['MoveTo', {}], ['LineTo', { X: len }]])),
  })
}

/**
 * The demo drawing: a flowchart made of masters (terminator, process, decision, connector) over a background page
 * with the title, and a page of shapes drawn by hand (curves, a group, a picture, formatted text).
 */
export async function sampleVisio(): Promise<Uint8Array<ArrayBuffer>> {
  const b = new VisioBuilder()
  b.theme = themeXml()
  const text = (size: number, color = '#1f2937', style = 0) => rows('Character', [{ Size: size / 72, Color: color, Style: style, Font: 'Calibri' }])

  // ---- masters
  const process = b.master(
    'Process',
    shape({ id: 5, name: 'Process', style: STYLE.NORMAL, box: [1, 1, 1.75, 0.75], cells: { FillForegnd: '#dbeafe', LineColor: '#1d4ed8', LineWeight: 1.25 / 72, Rounding: 0.08 }, sections: text(11) + relRect() }),
  )
  const decision = b.master(
    'Decision',
    shape({
      id: 5, name: 'Decision', style: STYLE.NORMAL, box: [1, 1, 1.7, 1.1], cells: { FillForegnd: '#fef3c7', LineColor: '#b45309', LineWeight: 1.25 / 72 },
      sections: text(11) + geom(0, [['RelMoveTo', { X: 0, Y: 0.5 }], ['RelLineTo', { X: 0.5, Y: 0 }], ['RelLineTo', { X: 1, Y: 0.5 }], ['RelLineTo', { X: 0.5, Y: 1 }], ['RelLineTo', { X: 0, Y: 0.5 }]]),
    }),
  )
  // a stadium: two lines joined by half circles (ArcTo: how far the arc bows from its chord)
  const terminator = b.master(
    'Terminator',
    shape({
      id: 5, name: 'Terminator', style: STYLE.NORMAL, box: [1, 1, 1.5, 0.6], cells: { FillForegnd: '#dcfce7', LineColor: '#15803d', LineWeight: 1.25 / 72 },
      sections: text(11, '#14532d', 1) + geom(0, [['MoveTo', { X: 0.3, Y: 0 }], ['LineTo', { X: 1.2, Y: 0 }], ['ArcTo', { X: 1.2, Y: 0.6, A: 0.3 }], ['LineTo', { X: 0.3, Y: 0.6 }], ['ArcTo', { X: 0.3, Y: 0, A: 0.3 }]]),
    }),
  )
  const connector = b.master(
    'Dynamic connector',
    shape({
      id: 5, name: 'Dynamic connector', style: STYLE.NORMAL,
      cells: { PinX: 1, PinY: 1, Width: 1, Height: 0, LocPinX: 0.5, LocPinY: 0, BeginX: 0.5, BeginY: 1, EndX: 1.5, EndY: 1, LineColor: '#475569', LineWeight: 1 / 72, EndArrow: 4, ObjType: 2, TextBkgnd: '#ffffff' },
      sections: text(9, '#475569') + geom(0, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: 1, Y: 0 }]], { NoFill: 1 }),
    }),
  )

  // ---- the background page: a title band and a frame
  const back = b.page(
    'Background',
    shape({ id: 1, style: STYLE.NORMAL, box: [5.5, 7.95, 11, 1.1], cells: { FillForegnd: '#1e3a8a', LinePattern: 0, LeftMargin: 0.4, VerticalAlign: 1 }, sections: rows('Character', [{ Size: 22 / 72, Color: '#ffffff', Style: 1 }]) + rows('Paragraph', [{ HorzAlign: 0 }]) + relRect(), text: 'Order handling\n' }) +
      shape({ id: 2, style: STYLE.NORMAL, box: [5.5, 0.3, 11, 0.6], cells: { FillPattern: 0, LinePattern: 0, RightMargin: 0.4 }, sections: rows('Character', [{ Size: 9 / 72, Color: '#64748b' }]) + rows('Paragraph', [{ HorzAlign: 2 }]) + relRect(), text: 'Northwind Field Lab · process map\n' }),
    { background: true },
  )

  // ---- page 1: the flowchart, every shape an instance of a master
  let id = 0
  const inst = (master: number, x: number, y: number, label: string, extra: Partial<ShapeSpec> = {}) => shape({ id: ++id, master, cells: { PinX: x, PinY: y, ...extra.cells }, sections: extra.sections, text: `${esc(label)}\n` })
  const link = (x1: number, y1: number, x2: number, y2: number, label?: string, extra: Record<string, Val> = {}) => {
    const len = Math.hypot(x2 - x1, y2 - y1)
    const angle = Math.atan2(y2 - y1, x2 - x1)
    return line(++id, x1, y1, x2, y2, {
      master: connector,
      style: undefined,
      // the label stays level, whatever way the connector goes
      cells: { ...(label ? { TxtPinX: len / 2, TxtPinY: 0, TxtWidth: 0.5, TxtHeight: 0.22, TxtLocPinX: 0.25, TxtLocPinY: 0.11, TxtAngle: -angle } : {}), ...extra },
      text: label ? `${esc(label)}\n` : undefined,
    })
  }
  const y1 = 5.6
  const y2 = 3.5
  const flow =
    inst(terminator, 1.4, y1, 'Start') +
    inst(process, 3.65, y1, 'Receive order') +
    inst(decision, 6.1, y1, 'In stock?') +
    // wider than its master: the size is the instance's own, the geometry follows it
    inst(process, 8.9, y1, 'Pick, pack and ship', { cells: { Width: 2.2, LocPinX: 1.1 } }) +
    inst(process, 6.1, y2, 'Back-order', { cells: { FillForegnd: '#fee2e2', LineColor: '#b91c1c' } }) +
    inst(process, 8.9, y2, 'Notify customer') +
    inst(terminator, 8.9, 1.9, 'Done') +
    link(2.15, y1, 2.775, y1) +
    link(4.525, y1, 5.25, y1) +
    link(6.95, y1, 7.8, y1, 'yes') +
    link(6.1, y1 - 0.55, 6.1, y2 + 0.375, 'no', { LinePattern: 2, LineColor: '#b91c1c' }) +
    link(8.9, y1 - 0.375, 8.9, y2 + 0.375) +
    link(6.975, y2, 8.025, y2) +
    link(8.9, y2 - 0.375, 8.9, 2.2)
  b.page('Flowchart', flow, { backPage: back, cells: { ThemeIndex: 33 } })

  // ---- page 2: shapes drawn by hand
  b.image('photo.png', png(240, 160))
  id = 0
  const shapes =
    // an ellipse with a gradient
    shape({
      id: ++id, name: 'Ellipse', style: STYLE.NORMAL, box: [1.6, 6.45, 2.2, 1.3], cells: { FillGradientEnabled: 1, FillGradientAngle: Math.PI / 2, LineColor: '#0f766e' },
      sections: rows('FillGradient', [{ GradientStopColor: '#ccfbf1', GradientStopColorTrans: 0, GradientStopPosition: 0 }, { GradientStopColor: '#14b8a6', GradientStopColorTrans: 0, GradientStopPosition: 1 }]) + text(12, '#134e4a') + geom(0, [['Ellipse', { X: 1.1, Y: 0.65, A: 2.2, B: 0.65, C: 1.1, D: 1.3 }]]),
      text: 'Ellipse\n',
    }) +
    // turned 20°
    shape({ id: ++id, name: 'Turned', style: STYLE.NORMAL, box: [4.3, 6.45, 1.9, 0.9], cells: { Angle: (20 * Math.PI) / 180, FillForegnd: '#ede9fe', LineColor: '#6d28d9', LinePattern: 2 }, sections: text(12, '#4c1d95') + relRect(), text: 'Turned 20°\n' }) +
    // a cylinder: elliptical arcs, a second section for the rim that is drawn but not filled
    shape({
      id: ++id, name: 'Database', style: STYLE.NORMAL, box: [6.9, 6.45, 1.3, 1.5], cells: { FillForegnd: '#e0f2fe', LineColor: '#0369a1' },
      sections:
        text(11, '#0c4a6e') +
        geom(0, [['MoveTo', { X: 0, Y: 1.3 }], ['LineTo', { X: 0, Y: 0.2 }], ['EllipticalArcTo', { X: 1.3, Y: 0.2, A: 0.65, B: 0, C: 0, D: 3.25 }], ['LineTo', { X: 1.3, Y: 1.3 }], ['EllipticalArcTo', { X: 0, Y: 1.3, A: 0.65, B: 1.5, C: 0, D: 3.25 }]]) +
        geom(1, [['MoveTo', { X: 0, Y: 1.3 }], ['EllipticalArcTo', { X: 1.3, Y: 1.3, A: 0.65, B: 1.1, C: 0, D: 3.25 }]], { NoFill: 1 }),
      text: 'Orders\n',
    }) +
    // a five-pointed star: one polyline
    shape({
      id: ++id, name: 'Star', style: STYLE.NORMAL, box: [9.3, 6.45, 1.4, 1.4], cells: { FillForegnd: '#fde68a', LineColor: '#b45309' },
      sections: geom(0, [['RelMoveTo', { X: 0.5, Y: 1 }], ['PolylineTo', { X: 0.7, Y: 1.4, A: 'POLYLINE(0, 0, 0.62,0.64, 1,0.62, 0.7,0.38, 0.81,0, 0.5,0.22, 0.19,0, 0.3,0.38, 0,0.62, 0.38,0.64)' }]]),
    }) +
    // a group: its own frame and title behind two members, one of them mirrored
    shape({
      id: ++id, type: 'Group', name: 'Group', style: STYLE.NORMAL, box: [2.3, 4.3, 3.6, 1.8], cells: { FillForegnd: '#f1f5f9', LineColor: '#94a3b8', DisplayMode: 1, VerticalAlign: 0, TopMargin: 0.08 },
      sections: text(10, '#475569', 1) + relRect(),
      text: 'Group\n',
      shapes: [
        shape({ id: ++id, style: STYLE.NORMAL, box: [0.95, 0.7, 1.3, 0.8], cells: { FillForegnd: '#bbf7d0', LineColor: '#15803d' }, sections: text(10) + geom(0, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: 1.3, Y: 0.4 }], ['LineTo', { X: 0, Y: 0.8 }], ['LineTo', { X: 0, Y: 0 }]]), text: 'A\n' }),
        shape({ id: ++id, style: STYLE.NORMAL, box: [2.65, 0.7, 1.3, 0.8], cells: { FlipX: 1, FillForegnd: '#bfdbfe', LineColor: '#1d4ed8' }, sections: text(10) + geom(0, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: 1.3, Y: 0.4 }], ['LineTo', { X: 0, Y: 0.8 }], ['LineTo', { X: 0, Y: 0 }]]), text: 'B\n' }),
      ],
    }) +
    // a picture
    shape({ id: ++id, type: 'Foreign', name: 'Photo', style: STYLE.NORMAL, box: [5.8, 4.3, 2.4, 1.6], cells: { LinePattern: 0, FillPattern: 0 }, foreign: 'rId1' }) +
    // a curve with a line end at both ends
    shape({
      id: ++id, name: 'Curve', style: STYLE.NORMAL, box: [9.1, 4.3, 2.6, 1.4], cells: { LineColor: '#be123c', LineWeight: 2 / 72, BeginArrow: 10, EndArrow: 4, EndArrowSize: 3 },
      sections: geom(0, [['MoveTo', { X: 0, Y: 0.2 }], ['NURBSTo', { X: 2.6, Y: 1.2, A: 0, B: 1, C: 0, D: 1, E: 'NURBS(1, 3, 0, 0, 0.3,1.2,0,1, 0.7,-0.2,0,1)' }]], { NoFill: 1 }),
    }) +
    // formatted text: runs of characters, paragraphs, bullets, a field
    shape({
      id: ++id, name: 'Notes', style: STYLE.NORMAL, box: [3.1, 1.95, 5.2, 2.1], cells: { FillForegnd: '#fffbeb', LineColor: '#d97706', VerticalAlign: 0, LeftMargin: 0.15, TopMargin: 0.12 },
      sections:
        rows('Character', [{ Size: 14 / 72, Color: '#92400e', Style: 1 }, { Size: 10.5 / 72, Color: '#1f2937', Style: 0 }, { Size: 10.5 / 72, Color: '#b91c1c', Style: 3 }, { Size: 10.5 / 72, Color: '#1f2937', Style: 4 }]) +
        rows('Paragraph', [{ HorzAlign: 0, SpAfter: 4 / 72 }, { HorzAlign: 0, Bullet: 1, IndLeft: 0.1 }]) +
        relRect(),
      text: "<cp IX='0'/><pp IX='0'/>Notes\n<cp IX='1'/><pp IX='1'/>Shapes keep their <cp IX='2'/>fills, lines<cp IX='1'/> and text\nThe page is <cp IX='3'/>Page-2<cp IX='1'/>, drawn <fld IX='0'>2 June 2026</fld>\nMasters, groups and themes are followed\n",
    }) +
    // every look from the theme: its fill, its line, its text color and font
    shape({ id: ++id, name: 'Themed', style: STYLE.THEMED, box: [8.6, 2.4, 2.2, 1], sections: rows('Character', [{ Size: 12 / 72 }]) + relRect(), text: 'Themed\n' }) +
    // on a hidden layer: not shown
    shape({ id: ++id, name: 'Hidden', style: STYLE.NORMAL, box: [8.6, 1.2, 2.2, 0.6], cells: { LayerMember: '0', FillForegnd: '#ef4444' }, sections: relRect(), text: 'Hidden layer\n' })
  b.page('Shapes', shapes, { backPage: back, images: ['photo.png'], layers: [false], cells: { ThemeIndex: 33 } })
  return b.zip()
}
