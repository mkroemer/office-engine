/**
 * The smallest Word documents for tests: document.xml from a body, styles.xml from style elements and document
 * defaults. Paragraphs and tables are written as WordprocessingML strings by the helpers below.
 */
import { writeZip } from '../src/core/zipwrite'

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

export interface DocxParts {
  /** <w:style> elements */
  styles?: string
  /** the inside of <w:pPrDefault><w:pPr> */
  pPrDefault?: string
  /** the inside of <w:rPrDefault><w:rPr> */
  rPrDefault?: string
  /** the body's section properties (page size, margins): a letter page with 1 inch margins when absent */
  sectPr?: string
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** a paragraph: `pPr` is the inside of <w:pPr>, the text one run (`\n`: a line break) */
export const para = (text: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${text ? `<w:r>${text.split('\n').map((t, i) => `${i ? '<w:br/>' : ''}<w:t xml:space="preserve">${esc(t)}</w:t>`).join('')}</w:r>` : ''}</w:p>`

/** a table with a grid line on every cell: `rows[r][c]` the text of a cell; `header` leading rows repeat on each page */
export function table(rows: string[][], opts: { widths?: number[]; header?: number; trPr?: (r: number) => string; tcPr?: (r: number, c: number) => string } = {}) {
  const n = rows[0]?.length ?? 1
  const widths = opts.widths ?? Array<number>(n).fill(Math.floor(9360 / n))
  const line = '<w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/>'
  const trs = rows.map((cells, r) => {
    const trPr = `${r < (opts.header ?? 0) ? '<w:tblHeader/>' : ''}${opts.trPr?.(r) ?? ''}`
    const tcs = cells.map((t, c) => `<w:tc><w:tcPr><w:tcW w:w="${widths[c]}" w:type="dxa"/>${opts.tcPr?.(r, c) ?? ''}</w:tcPr>${para(t)}</w:tc>`)
    return `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${tcs.join('')}</w:tr>`
  })
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${line}</w:tblBorders></w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${trs.join('')}</w:tbl>`
}

const LETTER = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'

/** a .docx holding `body` (paragraphs and tables) */
export function bareDocx(body: string, parts: DocxParts = {}): Promise<Blob> {
  const defaults = `<w:docDefaults><w:rPrDefault><w:rPr>${parts.rPrDefault ?? '<w:sz w:val="22"/>'}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>${parts.pPrDefault ?? ''}</w:pPr></w:pPrDefault></w:docDefaults>`
  const normal = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
  return writeZip(
    new Map(
      Object.entries({
        '[Content_Types].xml': `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
        '_rels/.rels': `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
        'word/_rels/document.xml.rels': `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
        'word/styles.xml': `${XML}<w:styles xmlns:w="${W}">${defaults}${normal}${parts.styles ?? ''}</w:styles>`,
        'word/document.xml': `${XML}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr>${parts.sectPr ?? LETTER}</w:sectPr></w:body></w:document>`,
      }),
    ),
    DOCX,
  )
}
