/**
 * Minimal workbooks for tests, written part by part: what a test needs to say (cells, defined names, a chart sheet,
 * what follows the rows) and nothing else.
 */
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { writeZip } from '../src/core/zipwrite'

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

export interface SheetSpec {
  name: string
  /** rows of the sheet: `<row>` elements (see `row`), or the whole `<sheetData>` content */
  rows?: string
  /** elements before `<sheetData>` (sheetViews, cols…) */
  head?: string
  /** elements after `<sheetData>` (autoFilter, mergeCells…) */
  tail?: string
  /** a chart sheet or a dialog sheet: listed in the workbook, not a worksheet */
  kind?: 'worksheet' | 'chartsheet' | 'dialogsheet'
  state?: string
}

export interface BookSpec {
  sheets: SheetSpec[]
  /** content of `<definedNames>` */
  names?: string
  /** attributes of `<workbookView>` */
  view?: string
  /** attributes of `<workbookPr>` */
  pr?: string
  /** the styles part */
  styles?: string
  /** content type of the workbook part (a macro-enabled one, a template) */
  mainType?: string
  /** more parts: path → content */
  parts?: Record<string, string | Uint8Array>
  /** more `<Override>` / `<Default>` elements of the content types */
  types?: string
  /** more relationships of the workbook part */
  rels?: string
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/** a cell: a number, a text (inline), or `{ f, v }` a formula with its cached value */
export type CellSpec = number | string | boolean | { f: string; v?: number | string; shared?: { si: number; ref?: string }; s?: number } | { s: number; v?: number | string }

const colName = (c: number) => {
  let s = ''
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}

export function cell(ref: string, v: CellSpec): string {
  if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`
  if (typeof v === 'boolean') return `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`
  if (typeof v === 'string') return `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`
  const s = v.s !== undefined ? ` s="${v.s}"` : ''
  if (!('f' in v)) return v.v === undefined ? `<c r="${ref}"${s}/>` : typeof v.v === 'number' ? `<c r="${ref}"${s}><v>${v.v}</v></c>` : `<c r="${ref}"${s} t="inlineStr"><is><t>${esc(v.v)}</t></is></c>`
  const fa = v.shared ? ` t="shared" si="${v.shared.si}"${v.shared.ref ? ` ref="${v.shared.ref}"` : ''}` : ''
  const f = v.f ? `<f${fa}>${esc(v.f)}</f>` : `<f${fa}/>`
  if (v.v === undefined) return `<c r="${ref}"${s}>${f}</c>`
  return typeof v.v === 'number' ? `<c r="${ref}"${s}>${f}<v>${v.v}</v></c>` : `<c r="${ref}"${s} t="str">${f}<v>${esc(v.v)}</v></c>`
}

/** rows from a grid: `grid[r][c]`, `null` for no cell */
export function rows(grid: (CellSpec | null)[][], from = 0): string {
  return grid
    .map((cells, i) => {
      const r = from + i + 1
      const xml = cells.map((v, c) => (v === null ? '' : cell(colName(c) + r, v))).join('')
      return xml ? `<row r="${r}">${xml}</row>` : ''
    })
    .join('')
}

export function sheetXml(s: SheetSpec): string {
  if (s.kind === 'chartsheet') return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<chartsheet xmlns="${NS}" xmlns:r="${NS_R}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>`
  if (s.kind === 'dialogsheet') return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dialogsheet xmlns="${NS}" xmlns:r="${NS_R}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></dialogsheet>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS}" xmlns:r="${NS_R}">${s.head ?? ''}<sheetData>${s.rows ?? ''}</sheetData>${s.tail ?? ''}</worksheet>`
}

export function book(spec: BookSpec): Promise<Blob> {
  const parts = new Map<string, string | Uint8Array>()
  const dir = { worksheet: 'worksheets', chartsheet: 'chartsheets', dialogsheet: 'dialogsheets' }
  const paths = spec.sheets.map((s, i) => `xl/${dir[s.kind ?? 'worksheet']}/sheet${i + 1}.xml`)
  let types = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
  types += `<Override PartName="/xl/workbook.xml" ContentType="${spec.mainType ?? CT + '.sheet.main+xml'}"/>`
  let rels = ''
  let list = ''
  spec.sheets.forEach((s, i) => {
    const kind = s.kind ?? 'worksheet'
    types += `<Override PartName="/${paths[i]}" ContentType="${CT}.${kind}+xml"/>`
    rels += `<Relationship Id="rId${i + 1}" Type="${REL}/${kind}" Target="${paths[i].slice(3)}"/>`
    list += `<sheet name="${esc(s.name)}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`
    parts.set(paths[i], sheetXml(s))
  })
  if (spec.styles) {
    types += `<Override PartName="/xl/styles.xml" ContentType="${CT}.styles+xml"/>`
    rels += `<Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/>`
    parts.set('xl/styles.xml', spec.styles)
  }
  types += spec.types ?? ''
  rels += spec.rels ?? ''
  parts.set('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${types}</Types>`)
  parts.set('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`)
  parts.set('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`)
  parts.set(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS}" xmlns:r="${NS_R}">${spec.pr !== undefined ? `<workbookPr ${spec.pr}/>` : ''}` +
      `${spec.view !== undefined ? `<bookViews><workbookView ${spec.view}/></bookViews>` : ''}<sheets>${list}</sheets>${spec.names ? `<definedNames>${spec.names}</definedNames>` : ''}</workbook>`,
  )
  for (const [k, v] of Object.entries(spec.parts ?? {})) parts.set(k, v)
  return writeZip(parts, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}

export interface Opened {
  engine: Engine
  info: any
  events: { event: string; data: any }[]
  /** loads a sheet (as showing its tab does) */
  load(sheet: number): Promise<void>
  /** the text shown in a cell ('' when there is none) */
  shown(sheet: number, r: number, c: number): Promise<string>
  /** formula of a cell (with its "="), undefined when it has none */
  formula(sheet: number, r: number, c: number): Promise<string | undefined>
  set(sheet: number, ref: string, input: string): Promise<any>
  /** saves; the parts of the saved file by path */
  save(): Promise<{ blob: Blob; text(path: string): Promise<string>; has(path: string): boolean; paths: string[] }>
}

export function at(ref: string): { r: number; c: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  let c = 0
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64)
  return { r: Number(m[2]) - 1, c: c - 1 }
}

export async function unzip(blob: Blob) {
  const zip = await ZipArchive.open(new BlobSource(blob))
  const paths = zip.list().map((e) => e.name)
  return { blob, text: (path: string) => zip.text(path), has: (path: string) => paths.includes(path), paths }
}

/** opens a workbook in an engine; sheet 0 is loaded (the others when asked, like tabs) */
export async function openBook(src: Blob, first = 0): Promise<Opened> {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, {})
  const load = async (sheet: number) => {
    const from = events.length
    await engine.call('xlsx.load', [sheet])
    while (!events.slice(from).some((e) => e.event === 'xlsx:progress' && e.data.sheet === sheet && e.data.done)) await new Promise((r) => setTimeout(r, 5))
  }
  await load(first)
  return {
    engine, info, events, load,
    shown: async (sheet, r, c) => {
      const b: any = await engine.call('xlsx.block', [sheet, r, r + 1])
      return b.rows[0]?.cells.find((x: any) => x[0] === c)?.[1] ?? ''
    },
    formula: async (sheet, r, c) => ((await engine.call('xlsx.cell', [sheet, r, c])) as any)?.formula,
    set: (sheet, ref, input) => engine.call('xlsx.set', [sheet, [{ ...at(ref), input }]]),
    save: async () => unzip((await engine.call('save', [])) as Blob),
  }
}
