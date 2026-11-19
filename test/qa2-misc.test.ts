/**
 * QA 2026-11-19 (CSV, converters, file detection, error messages): N3-004 / P2-027, P2-026, P2-028, P2-029, N3-005 /
 * P3-022 / P3-029, N3-009, N3-011, N3-012 / P2-039, P2-038, P2-045, P3-010, P3-021, P3-039.
 */
import { describe, expect, test } from 'bun:test'
import { decodeText } from '../src/core/plain'
import { Engine } from '../src/core/worker/engine'
import type { OpenParams } from '../src/core/worker/protocol'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { rebuildZip, writeZip } from '../src/core/zipwrite'
import type { BlockOut, SheetProgress } from '../src/core/xlsx/types'
import { convertOdfZip } from '../src/core/legacy'
import { OdfStyles, parseOdfXml } from '../src/core/legacy/odf/common'
import { excelFormat } from '../src/core/legacy/odf/ods-format'
import { Cfb } from '../src/core/legacy/cfb'
import { docToDocx } from '../src/core/legacy/doc'
import { rtfToDocx } from '../src/core/legacy/rtf'
import { writeCfb } from './cfb-writer'
import { DocBuilder } from './doc-builder'
import { relRect, shape as vshape, STYLE, VisioBuilder } from './visio-builder'
import type { SlideEl } from '../src/core/pptx/types'
import { boolErr, cat, formula, label, ptg, rec, stringRec, u16, u32, u8, xlsFile } from './xls-builder'

const enc = (s: string) => new TextEncoder().encode(s)

/** the bytes of `text` in a single-byte Windows code page; a letter the page has only as base + mark is written as the two */
function encodeIn(label: string, text: string): Uint8Array {
  const dec = new TextDecoder(label)
  const map = new Map<string, number>()
  for (let i = 0; i < 256; i++) {
    const ch = dec.decode(Uint8Array.of(i))
    if (!map.has(ch)) map.set(ch, i)
  }
  const out: number[] = []
  const put = (ch: string) => {
    if (map.has(ch)) return out.push(map.get(ch)!)
    const d = ch.normalize('NFD')
    let base = d[0]
    let marks = [...d.slice(1)]
    // fold in the marks the page has the letter with (o + circumflex = ô), the others are written after it
    for (const m of [...marks]) {
      const c = (base + m).normalize('NFC')
      if (c.length === 1 && map.has(c)) {
        base = c
        marks = marks.filter((x) => x !== m)
      }
    }
    for (const x of [base, ...marks]) if (!map.has(x)) throw new Error(`not in ${label}: ${ch}`)
    for (const x of [base, ...marks]) out.push(map.get(x)!)
  }
  for (const ch of text.normalize('NFC')) put(ch)
  return Uint8Array.from(out)
}

async function openCsv(bytes: Uint8Array | string, name = 'a.csv', params: OpenParams = {}) {
  const events: { event: string; data: any }[] = []
  const eng = new Engine((event, data) => events.push({ event, data }))
  await eng.open(new Blob([typeof bytes === 'string' ? enc(bytes) : (bytes as Uint8Array<ArrayBuffer>)]), { name, ...params })
  await (eng as any).xlsx.loadSheet(0, () => {})
  return { eng, events }
}
const cellText = async (eng: Engine, r: number, c: number) => ((await eng.call('xlsx.cell', [0, r, c])) as { value: string } | undefined)?.value
const exported = async (eng: Engine, sep: string) => new Uint8Array(await ((await eng.call('xlsx.text', [0, sep])) as Blob).arrayBuffer())
const exportedText = async (eng: Engine, sep: string) => new TextDecoder().decode(await exported(eng, sep)).replace(/^﻿/, '')

describe('N3-004 / P2-027: Windows-1258 (Vietnamese) CSV', () => {
  const csv = 'Họ tên,Địa chỉ\r\nNguyễn Văn Đức,Hà Nội\r\nTrần Thị Ánh,Đà Nẵng\r\n'

  test('the test encoder gives what Excel writes (1252 reads it as the QA saw)', () => {
    expect(new TextDecoder('windows-1252').decode(encodeIn('windows-1258', 'Nguyễn Văn Đức'))).toContain('NguyêÞn')
  })

  test('a CSV read as 1258, in NFC', async () => {
    const { eng } = await openCsv(encodeIn('windows-1258', csv))
    expect(await cellText(eng, 0, 0)).toBe('Họ tên')
    expect(await cellText(eng, 1, 0)).toBe('Nguyễn Văn Đức')
    expect(await cellText(eng, 1, 1)).toBe('Hà Nội')
    expect(await cellText(eng, 2, 1)).toBe('Đà Nẵng')
  })

  test('a text file the same (decodeText), and one that is not Vietnamese stays Windows-1252', () => {
    expect(decodeText(encodeIn('windows-1258', csv))).toBe(csv)
    const it = 'Così è già più bello, però: ì ò Þ ð ã õ ý'
    expect(decodeText(encodeIn('windows-1252', it))).toBe(it)
    const pt = 'Não há ação, informação e coração'
    expect(decodeText(encodeIn('windows-1252', pt))).toBe(pt)
  })

  test('an Italian / Portuguese CSV is not taken for Vietnamese', async () => {
    const { eng } = await openCsv(encodeIn('windows-1252', 'città,però\r\nfunzionalità,così\r\n'))
    expect(await cellText(eng, 0, 0)).toBe('città')
    expect(await cellText(eng, 1, 1)).toBe('così')
  })

  test('the encoding can be asked for: `encoding` in the open parameters', async () => {
    const bytes = encodeIn('windows-1258', csv)
    const forced = await openCsv(bytes, 'a.csv', { encoding: 'windows-1252' })
    expect(await cellText(forced.eng, 1, 0)).toContain('NguyêÞn')
    const vi = await openCsv(bytes, 'a.csv', { encoding: 'windows-1258' })
    expect(await cellText(vi.eng, 1, 0)).toBe('Nguyễn Văn Đức')
    // text files too
    const e = new Engine(() => {})
    await e.open(new Blob([encodeIn('windows-1258', 'Hà Nội\n') as Uint8Array<ArrayBuffer>]), { name: 'a.txt', encoding: 'windows-1258' })
    // an unknown label is refused, saying so
    await expect(openCsv(bytes, 'a.csv', { encoding: 'klingon' })).rejects.toThrow(/Unknown text encoding/)
  })
})

describe('P2-026: more columns than a sheet holds', () => {
  test('a warning says the extra columns were left out', async () => {
    const wide = Array.from({ length: 17_000 }, (_, i) => i).join(',')
    const { eng, events } = await openCsv(`${wide}\r\n1,2\r\n`)
    const w = events.filter((e) => e.event === 'warning')
    expect(w.length).toBe(1)
    expect(w[0].data.text).toMatch(/17,000 columns.*16,384/)
    expect(await cellText(eng, 0, 16383)).toBe('16383')
  })

  test('no warning for an ordinary file, nor for exactly 16 384 columns', async () => {
    const full = Array.from({ length: 16_384 }, (_, i) => i).join(',')
    expect((await openCsv(`${full}\r\n`)).events.filter((e) => e.event === 'warning')).toEqual([])
    expect((await openCsv('a,b\r\n1,2\r\n')).events.filter((e) => e.event === 'warning')).toEqual([])
  })
})

describe('P2-029 / P2-028: CSV export', () => {
  test('texts a spreadsheet would run as formulas are neutralized, numbers stay numbers', async () => {
    const { eng } = await openCsv('"=2+2","+cmd","@SUM(A1)","-b",-5,+7,plain,"=x"\r\n')
    expect(await cellText(eng, 0, 4)).toBe('-5')
    expect(await exportedText(eng, ',')).toBe(`'=2+2,'+cmd,'@SUM(A1),'-b,-5,7,plain,'=x\r\n`)
  })

  test('a formula result that is text starting with = is neutralized too', async () => {
    const { eng } = await openCsv('x\r\n')
    await eng.call('xlsx.set', [0, [{ r: 0, c: 1, input: '="="&"2+2"' }]])
    expect(await exportedText(eng, ',')).toBe(`x,'=2+2\r\n`)
  })

  test('TSV and ; CSV start with a byte order mark like the comma one', async () => {
    const { eng } = await openCsv('a,b\r\nNguyễn,2\r\n')
    for (const sep of [',', '\t', ';']) expect([...(await exported(eng, sep)).slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf])
    expect(await exportedText(eng, ';')).toBe('a;b\r\nNguyễn;2\r\n')
  })

  test('P2-028: a General number keeps its digits, a formatted one what is shown', async () => {
    const { eng } = await openCsv('123456789012,1234.5678912345,0.1,1234567.891\r\n')
    expect(await exportedText(eng, ',')).toBe('123456789012,1234.5678912345,0.1,1234567.891\r\n')
    await eng.call('xlsx.format', [0, { r0: 0, c0: 1, r1: 0, c1: 1 }, { numFmt: '#,##0.00' }])
    expect(await exportedText(eng, '\t')).toBe('123456789012\t"1,234.57"\t0.1\t1234567.891\r\n'.replace('"1,234.57"', '1,234.57'))
  })

  test('P2-028: 15 significant digits, no exponent for plain forms up to 21 characters', async () => {
    const { eng } = await openCsv('1E+20,x,123456789012345\r\n')
    await eng.call('xlsx.set', [0, [{ r: 0, c: 1, input: '=0.1+0.2' }]])
    expect(await exportedText(eng, ',')).toBe('100000000000000000000,0.3,123456789012345\r\n')
  })
})

// ---------------------------------------------------------------- file detection and messages (N3-005, P3-022, P3-029)

const SAMPLES = `${import.meta.dir}/../playground/public/samples`
const sample = async (n: string) => new Uint8Array(await Bun.file(`${SAMPLES}/${n}`).arrayBuffer())
async function openAs(bytes: Uint8Array | string, name: string) {
  const eng = new Engine(() => {})
  const info = await eng.open(new Blob([typeof bytes === 'string' ? enc(bytes) : (bytes as Uint8Array<ArrayBuffer>)]), { name })
  return { eng, info }
}
const failure = async (bytes: Uint8Array | string, name: string) => {
  try {
    await openAs(bytes, name)
  } catch (e) {
    return (e as Error).message
  }
  throw new Error(`${name} opened`)
}
const RAW = /inflate|zlib|Trailing junk|unexpected end of file|TypeError|end of central directory/i

describe('P3-029 / N3-005: a file that is not what its name says', () => {
  const table = '<html><body><table><tr><td>Name</td><td>Amount</td></tr><tr><td>Café</td><td>12.5</td></tr></table></body></html>'

  test('a web page named .xls / .doc opens as a web page', async () => {
    for (const name of ['report.xls', 'report.doc', 'report.xlt']) {
      const { info } = await openAs(table, name)
      expect(info.imported).toBe('html')
      expect(info.kind).toBe('docx')
    }
  })

  test('CSV and TSV named .xls open as a sheet, text named .doc as a document', async () => {
    const csv = await openAs('Name,Amount\r\nCafé,12.5\r\nTea,3\r\n', 'csv-as.xls')
    expect(csv.info).toMatchObject({ kind: 'xlsx', imported: 'csv' })
    await (csv.eng as any).xlsx.loadSheet(0, () => {})
    expect(await cellText(csv.eng, 1, 0)).toBe('Café')
    expect(await cellText(csv.eng, 1, 1)).toBe('12.5')
    expect((await openAs('Name\tAmount\r\nCafé\t12.5\r\n', 'tsv-as.xls')).info).toMatchObject({ kind: 'xlsx', imported: 'tsv' })
    expect((await openAs('Dear team,\r\nthe meeting moved.\r\n', 'letter.doc')).info).toMatchObject({ kind: 'docx', imported: 'text' })
  })

  test('an MHTML page named .doc opens as its HTML', async () => {
    const html = '<html><body><p>Word HTML report</p></body></html>'
    const mht = ['MIME-Version: 1.0', 'Content-Type: multipart/related; boundary="----=_NextPart_01"; type="text/html"', '', 'This is a multi-part message.', '', '------=_NextPart_01', 'Content-Type: text/html; charset="utf-8"', 'Content-Transfer-Encoding: quoted-printable', '', html.replace('report', 'r=65port'), '', '------=_NextPart_01', 'Content-Type: image/png', 'Content-Transfer-Encoding: base64', '', 'iVBORw0KGgo=', '------=_NextPart_01--', ''].join('\r\n')
    const { eng, info } = await openAs(mht, 'mht-as.doc')
    expect(info).toMatchObject({ kind: 'docx', imported: 'html' })
    expect(await eng.call('plain.source', [])).toContain('Word HTML report')
  })

  test('XML 2003 says what it is and what to do', async () => {
    const xml = '<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet/></Workbook>'
    const msg = await failure(xml, 'xml2003-as.xls')
    expect(msg).toMatch(/Excel 2003 XML spreadsheet.*save it as \.xlsx/)
    expect(await failure(xml, 'x.xlsx')).toMatch(/Excel 2003 XML/)
  })

  test('a real .doc / .xls named .docx / .xlsx is converted', async () => {
    const doc = await openAs(await sample('sample.doc'), 'doc-as.docx')
    expect(doc.info).toMatchObject({ kind: 'docx', converted: 'doc' })
    const xls = await openAs(await sample('sample.xls'), 'xls-as.xlsx')
    expect(xls.info).toMatchObject({ kind: 'xlsx', converted: 'xls' })
    // and a package named .doc / no name
    expect((await openAs(await sample('sample.docx'), 'x.doc')).info.kind).toBe('docx')
  })

  test('a PDF, an image, an empty file: the message says what the file is', async () => {
    const pdf = enc('%PDF-1.7\n%âãÏÓ\n1 0 obj\n<< /Type /Catalog >>\nendobj\n')
    for (const name of ['congvan.pdf.docx', 'fake.docx', 'scan.xls', 'x.pptx', 'noname']) {
      const msg = await failure(pdf, name)
      expect(msg).toMatch(/^This is a PDF file/)
      expect(msg).not.toMatch(RAW)
    }
    expect(await failure(Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10, ...new Array(40).fill(0)]), 'logo.xlsx')).toMatch(/^This is a PNG image/)
    expect(await failure(new Uint8Array(0), 'e.docx')).toBe('The file is empty')
    expect(await failure(Uint8Array.from({ length: 300 }, (_, i) => (i * 7 + 200) & 0xff), 'random.docx')).toMatch(/not in a format that can be opened here/)
  })

  test('an HTML error page served as .docx is still an error, saying so', async () => {
    expect(await failure('<!doctype html><title>404</title><p>Not found</p>', 'report.docx')).toMatch(/web page.*error page/)
  })

  test('an encrypted OOXML package says so, whatever its name', async () => {
    const enc2 = writeCfb({ EncryptionInfo: new Uint8Array(40), EncryptedPackage: new Uint8Array(5000) })
    for (const name of ['password.docx', 'password.xlsx', 'password.pptx', 'password.doc', 'password']) {
      expect(await failure(enc2, name)).toBe('This file is password-protected (encrypted): open it in Office and save it without a password')
    }
  })
})

describe('P3-022: damaged files are told in words', () => {
  /** the entry's data bytes (as stored in the file) */
  async function entryAt(bytes: Uint8Array, name: string) {
    const zip = await ZipArchive.open(new BlobSource(new Blob([bytes as Uint8Array<ArrayBuffer>])))
    const e = zip.get(name)!
    return { e, at: (await (zip as any).dataOffset(e)) as number }
  }

  test('a part that does not unpack: no zlib error', async () => {
    const bytes = await sample('sample.docx')
    const { at, e } = await entryAt(bytes, 'word/document.xml')
    const bad = bytes.slice()
    for (let i = 2; i < Math.min(40, e.compSize); i++) bad[at + i] = 0xff
    // the document part is streamed: the failure comes with its text
    const events: { event: string; data: any }[] = []
    const e2 = new Engine((event, data) => events.push({ event, data }))
    let msg = ''
    try {
      await e2.open(new Blob([bad as Uint8Array<ArrayBuffer>]), { name: 'broken.docx' })
      await e2.call('docx.start', [])
      msg = events.find((x) => x.event === 'docx:done')?.data.error ?? ''
    } catch (err) {
      msg = (err as Error).message
    }
    expect(msg).toMatch(/^The file is damaged: the part “word\/document\.xml” does not unpack correctly/)
    expect(msg).not.toMatch(RAW)
  })

  test('a part whose checksum is wrong is damaged, and a file cut short says so', async () => {
    const bytes = await sample('sample.docx')
    const bad = bytes.slice()
    const dv = new DataView(bad.buffer)
    // central directory record of word/document.xml: flip its CRC
    for (let p = 0; p + 46 < bad.length; p++) {
      if (dv.getUint32(p, true) !== 0x02014b50) continue
      const n = dv.getUint16(p + 28, true)
      if (new TextDecoder().decode(bad.subarray(p + 46, p + 46 + n)) === 'word/document.xml') dv.setUint32(p + 16, 0xdeadbeef, true)
    }
    const events: { event: string; data: any }[] = []
    const eng = new Engine((event, data) => events.push({ event, data }))
    await eng.open(new Blob([bad]), { name: 'crc.docx' })
    await eng.call('docx.start', [])
    expect(events.find((x) => x.event === 'docx:done')?.data.error).toMatch(/^The file is damaged: the part “word\/document\.xml”/)
    expect(await failure(bytes.slice(0, Math.floor(bytes.length / 2)), 'cut.docx')).toMatch(/^The file is damaged: it is a ZIP archive that ends too soon/)
  })

  test('a Word 97 file whose streams cannot be read is damaged, not "Word 6.0 / 95"', async () => {
    const cfb = writeCfb({ WordDocument: new Uint8Array(2048), '1Table': new Uint8Array(512) })
    const msg = await failure(cfb, 'cfb-dir.doc')
    expect(msg).toMatch(/^The file is damaged/)
    expect(msg).not.toMatch(/Word 6/)
  })
})

describe('P3-021: broken packages are repaired on save or refused', () => {
  /** the package with the names of its parts rewritten in place (same length) in the local headers and the directory */
  function renamed(bytes: Uint8Array, map: (n: string) => string): Uint8Array {
    const out = bytes.slice()
    const dv = new DataView(out.buffer)
    for (let p = 0; p + 30 < out.length; p++) {
      const sig = dv.getUint32(p, true)
      const [head, lenAt] = sig === 0x04034b50 ? [30, 26] : sig === 0x02014b50 ? [46, 28] : [0, 0]
      if (!head) continue
      const n = dv.getUint16(p + lenAt, true)
      const name = new TextDecoder().decode(out.subarray(p + head, p + head + n))
      if (!/^[\w./\[\]\\-]+$/.test(name)) continue
      out.set(enc(map(name)), p + head)
    }
    return out
  }
  const withoutParts = async (name: string, drop: string[], extra: Record<string, string> = {}) => {
    const zip = await ZipArchive.open(new BlobSource(new Blob([(await sample(name)) as Uint8Array<ArrayBuffer>])))
    const changes = new Map<string, string | null>(drop.map((d) => [d, null]))
    for (const [k, v] of Object.entries(extra)) changes.set(k, v)
    return new Uint8Array(await (await rebuildZip(zip, changes as Map<string, string | null>)).arrayBuffer())
  }
  const names = async (blob: Blob) => (await ZipArchive.open(new BlobSource(blob))).list().map((e) => e.name)

  test('no [Content_Types].xml: written on save, for the three formats', async () => {
    for (const [file, kind] of [['sample.docx', 'docx'], ['sample.xlsx', 'xlsx'], ['sample.pptx', 'pptx']] as const) {
      const { eng, info } = await openAs(await withoutParts(file, ['[Content_Types].xml']), `no-ct.${kind}`)
      expect(info.kind).toBe(kind)
      if (kind === 'docx') await eng.call('docx.start', [])
      const saved = (await eng.call('save', [])) as Blob
      const zip = await ZipArchive.open(new BlobSource(saved))
      expect(zip.has('[Content_Types].xml')).toBe(true)
      const ct = await zip.text('[Content_Types].xml')
      // the saved file is one the engine reads back by its content types (the main part's type says the kind)
      expect(ct).toContain(kind === 'docx' ? 'wordprocessingml.document.main+xml' : kind === 'xlsx' ? 'spreadsheetml.sheet.main+xml' : 'presentationml.presentation.main+xml')
      expect(ct).toContain('Extension="rels"')
      const again = await openAs(new Uint8Array(await saved.arrayBuffer()), `again.${kind}`)
      expect(again.info.kind).toBe(kind)
      if (kind === 'xlsx') expect(ct).toContain('spreadsheetml.worksheet+xml')
      if (kind === 'docx') expect(ct).toContain('wordprocessingml.styles+xml')
    }
  })

  test('part names with backslashes: normalized on save, the main relationship kept', async () => {
    const bytes = renamed(await sample('sample.docx'), (n) => n.replace(/\//g, '\\'))
    const raw = await ZipArchive.open(new BlobSource(new Blob([bytes as Uint8Array<ArrayBuffer>])))
    expect(raw.list().some((e) => e.name.includes('\\'))).toBe(true)
    const { eng, info } = await openAs(bytes, 'backslash.docx')
    expect(info.kind).toBe('docx')
    await eng.call('docx.start', [])
    const saved = (await eng.call('save', [])) as Blob
    const list = await names(saved)
    expect(list.some((n) => n.includes('\\'))).toBe(false)
    expect(list).toContain('word/document.xml')
    expect(list).toContain('_rels/.rels')
    expect((await openAs(new Uint8Array(await saved.arrayBuffer()), 'fixed.docx')).info.kind).toBe('docx')
  })

  test('a main part that is not XML is refused, not shown blank', async () => {
    const bytes = await withoutParts('sample.docx', [], { 'word/document.xml': 'this is not xml, it is a plain text note' })
    const msg = await failure(bytes, 'doc-not-xml.docx')
    expect(msg).toBe('The file is damaged: its main part “word/document.xml” is not XML')
  })
})

// ---------------------------------------------------------------- converters

async function xlsSheet(bytes: Uint8Array<ArrayBuffer>, name = 'book.xls') {
  const progress: SheetProgress[] = []
  const eng = new Engine((event, data) => {
    if (event === 'xlsx:progress') progress.push(data as SheetProgress)
  })
  await eng.open(new Blob([bytes]), { name })
  await eng.call('xlsx.load', [0])
  await new Promise((r) => setTimeout(r, 50))
  const block = (await eng.call('xlsx.block', [0, 0, 60])) as BlockOut
  const text = (r: number, c: number) => block.rows.find((x) => x.r === r)?.cells.find((x) => x[0] === c)?.[1] ?? ''
  const formulaAt = async (r: number, c: number) => ((await eng.call('xlsx.cell', [0, r, c])) as { formula?: string; value?: string } | undefined)?.formula
  const saved = new Uint8Array(await ((await eng.call('save', [])) as Blob).arrayBuffer())
  const zip = await ZipArchive.open(new BlobSource(new Blob([saved])))
  return { eng, text, formulaAt, xml: await zip.text('xl/worksheets/sheet1.xml') }
}

describe('P2-038 / N3-012 / P2-039: Excel 97-2003 results and formulas', () => {
  test('an array formula returning text keeps its results', async () => {
    const rgce = cat(ptg.area(0, 2, 0, 0, false), ptg.str('!'), ptg.op(0x08))
    const array = rec(0x0221, u16(0), u16(2), u8(2), u8(2), u16(0), u32(0), u16(rgce.length), rgce)
    const file = xlsFile({
      sheets: [
        {
          name: 'S',
          records: [
            label(0, 0, 'b'), label(1, 0, 'a'), label(2, 0, 'c'),
            formula(0, 2, ptg.exp(0, 2), [0, 0]), array, stringRec('b!'),
            formula(1, 2, ptg.exp(0, 2), [0, 0]), stringRec('a!'),
            formula(2, 2, ptg.exp(0, 2), [0, 0]), stringRec('c!'),
          ],
        },
      ],
    })
    const s = await xlsSheet(file)
    expect([s.text(0, 2), s.text(1, 2), s.text(2, 2)]).toEqual(['b!', 'a!', 'c!'])
  })

  test('booleans are boolean cells, whatever the file wrote: a BOOLERR record, TRUE / FALSE constants, TRUE() / FALSE()', async () => {
    const file = xlsFile({
      sheets: [
        {
          name: 'S',
          records: [
            boolErr(0, 0, 1, false),
            formula(1, 0, ptg.bool(true), [1, 1]),
            formula(2, 0, ptg.bool(false), [1, 0]),
            formula(3, 0, ptg.func(34), [1, 1]),
            formula(4, 0, ptg.func(35), [1, 0]),
            // a real formula with a boolean result stays one
            formula(5, 0, cat(ptg.int(1), ptg.int(2), ptg.op(0x09)), [1, 1]),
          ],
        },
      ],
    })
    const s = await xlsSheet(file)
    expect([0, 1, 2, 3, 4, 5].map((r) => s.text(r, 0))).toEqual(['TRUE', 'TRUE', 'FALSE', 'TRUE', 'FALSE', 'TRUE'])
    for (const r of [0, 1, 2, 3, 4]) {
      expect(await s.formulaAt(r, 0)).toBeUndefined()
      expect(s.xml).toMatch(new RegExp(`<c r="A${r + 1}"[^>]*t="b"><v>[01]</v></c>`))
    }
    expect(await s.formulaAt(5, 0)).toBe('=1<2')
  })
})

// ---- OpenDocument spreadsheets

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ' +
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"'

async function odsFile(body: string, auto = '', styles = ''): Promise<Uint8Array<ArrayBuffer>> {
  const files = new Map<string, string>([
    ['mimetype', 'application/vnd.oasis.opendocument.spreadsheet'],
    ['content.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles>${auto}</office:automatic-styles><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`],
    ['styles.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS}><office:styles>${styles}</office:styles></office:document-styles>`],
  ])
  return new Uint8Array(await (await writeZip(files)).arrayBuffer())
}
const ocell = (attrs: string, text?: string) => `<table:table-cell ${attrs}>${text !== undefined ? `<text:p>${text}</text:p>` : ''}</table:table-cell>`
const orow = (...cells: string[]) => `<table:table-row>${cells.join('')}</table:table-row>`

async function odsSheet(body: string, auto = '', styles = '') {
  const bytes = await odsFile(body, auto, styles)
  const eng = new Engine(() => {})
  await eng.open(new Blob([bytes]), { name: 'book.ods' })
  await eng.call('xlsx.load', [0])
  await new Promise((r) => setTimeout(r, 50))
  const block = (await eng.call('xlsx.block', [0, 0, 60])) as BlockOut
  const text = (r: number, c: number) => block.rows.find((x) => x.r === r)?.cells.find((x) => x[0] === c)?.[1] ?? ''
  const detail = async (r: number, c: number) => (await eng.call('xlsx.cell', [0, r, c])) as { formula?: string; value?: string } | undefined
  const conv = await convertOdfZip(await ZipArchive.open(new BlobSource(new Blob([bytes]))))
  const zip = await ZipArchive.open(new BlobSource(conv!.file))
  return { text, detail, xml: await zip.text('xl/worksheets/sheet1.xml'), styles: await zip.text('xl/styles.xml') }
}

describe('N3-012 / P2-039: OpenDocument booleans and formulas', () => {
  test('booleans are boolean cells, with or without a TRUE() formula; a formula with a boolean result stays one', async () => {
    const body =
      '<table:table table:name="B">' +
      orow(
        ocell('office:value-type="boolean" office:boolean-value="true"', 'TRUE'),
        ocell('table:formula="of:=TRUE()" office:value-type="boolean" office:boolean-value="true"', 'TRUE'),
        ocell('table:formula="of:=FALSE()" office:value-type="boolean" office:boolean-value="false"', 'FALSE'),
        ocell('table:formula="of:=[.A2]&gt;1" office:value-type="boolean" office:boolean-value="true"', 'TRUE'),
      ) +
      orow(ocell('office:value-type="float" office:value="5"', '5')) +
      '</table:table>'
    const s = await odsSheet(body)
    expect([0, 1, 2, 3].map((c) => s.text(0, c))).toEqual(['TRUE', 'TRUE', 'FALSE', 'TRUE'])
    for (const c of [0, 1, 2]) expect((await s.detail(0, c))!.formula).toBeUndefined()
    expect(s.xml).toMatch(/<c r="B1"[^>]*t="b"><v>1<\/v><\/c>/)
    expect(s.xml).toMatch(/<c r="C1"[^>]*t="b"><v>0<\/v><\/c>/)
    expect((await s.detail(0, 3))!.formula).toBe('=A2>1')
  })

  test('a formula the converter cannot translate keeps its value', async () => {
    const body =
      '<table:table table:name="I">' +
      orow(ocell('office:value-type="float" office:value="1"', '1'), ocell('office:value-type="float" office:value="2"', '2'), ocell('office:value-type="float" office:value="3"', '3')) +
      orow(ocell('office:value-type="float" office:value="4"', '4'), ocell('table:formula="of:=[.A1:.C1]![.B1:.B2]" office:value-type="float" office:value="2"', '2')) +
      '</table:table>'
    const s = await odsSheet(body)
    expect(s.text(1, 1)).toBe('2')
  })
})

describe('P2-045: OpenDocument number formats', () => {
  const XMLNS = `${NS} xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"`
  const formats = (auto: string) => new OdfStyles(undefined, parseOdfXml(`<office:document-content ${XMLNS}><office:automatic-styles>${auto}</office:automatic-styles></office:document-content>`))

  test(';;; (hide the values) stays three empty sections, not General', () => {
    // as LibreOffice writes it: styles holding only an empty text
    const st = formats(
      '<number:number-style style:name="N10P0" style:volatile="true"><number:text/></number:number-style>' +
        '<number:number-style style:name="N10P1" style:volatile="true"><number:text/></number:number-style>' +
        '<number:number-style style:name="N10"><number:text/><style:map style:condition="value()&gt;0" style:apply-style-name="N10P0"/><style:map style:condition="value()&lt;0" style:apply-style-name="N10P1"/></number:number-style>',
    )
    expect(excelFormat(st, 'N10')).toBe(';;;')
    // a section that is only text stays text
    const t = formats('<number:number-style style:name="T"><number:text>n/a</number:text></number:number-style>')
    expect(excelFormat(t, 'T')).toBe('"n/a"')
  })

  test('currency with a symbol, two decimals, red negatives in parentheses', () => {
    const st = formats(
      '<number:currency-style style:name="N105P0" style:volatile="true"><number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol><number:number number:decimal-places="2" loext:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><number:text> </number:text></number:currency-style>' +
        '<number:currency-style style:name="N105"><style:text-properties fo:color="#ff0000"/><number:text>(</number:text><number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol><number:number number:decimal-places="2" loext:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><number:text>)</number:text><style:map style:condition="value()&gt;=0" style:apply-style-name="N105P0"/></number:currency-style>',
    )
    expect(excelFormat(st, 'N105')).toBe('[$$]#,##0.00 ;[Red]([$$]#,##0.00)')
  })

  test('a cell with ;;; is written with that format and keeps its value (what the grid shows for it is numfmt.ts)', async () => {
    const auto =
      '<number:number-style style:name="N10P0" style:volatile="true"><number:text/></number:number-style>' +
      '<number:number-style style:name="N10P1" style:volatile="true"><number:text/></number:number-style>' +
      '<number:number-style style:name="N10"><number:text/><style:map style:condition="value()&gt;0" style:apply-style-name="N10P0"/><style:map style:condition="value()&lt;0" style:apply-style-name="N10P1"/></number:number-style>' +
      '<style:style style:name="ce1" style:family="table-cell" style:data-style-name="N10"/>'
    const s = await odsSheet(`<table:table table:name="H"><table:table-row>${ocell('table:style-name="ce1" office:value-type="float" office:value="99"', '99')}</table:table-row></table:table>`, auto)
    expect((await s.detail(0, 0))!.value).toBe('99')
    expect(s.styles).toContain('formatCode=";;;"')
  })
})

describe('P3-010: .doc first-page header and footer', () => {
  test('the first-page footer is a footer reference, the first-page header a header reference', async () => {
    const d = new DocBuilder()
    d.p('Body text')
    d.headers(['EVEN HEADER', 'ODD HEADER', 'EVEN FOOTER', [{ text: 'ODD FOOTER page ' }, ...d.field('PAGE', '1')], 'FIRST HEADER', [{ text: 'FIRST FOOTER page ' }, ...d.field('PAGE', '1')]])
    const cfbBytes = d.build()
    const zip = await ZipArchive.open(new BlobSource(await docToDocx(Cfb.open(cfbBytes))))
    const doc = await zip.text('word/document.xml')
    expect(doc.match(/<w:headerReference w:type="first"/g)?.length).toBe(1)
    expect(doc.match(/<w:footerReference w:type="first"/g)?.length).toBe(1)
    const target = async (kind: 'header' | 'footer') => {
      const rid = new RegExp(`<w:${kind}Reference w:type="first" r:id="([^"]+)"`).exec(doc)![1]
      const rels = await zip.text('word/_rels/document.xml.rels')
      const t = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`).exec(rels)!
      return zip.text(`word/${t[1] ?? t[2]}`)
    }
    expect(await target('header')).toContain('FIRST HEADER')
    expect(await target('footer')).toContain('FIRST FOOTER page')
    expect(await target('footer')).toContain('<w:ftr ')
  })
})

describe('P3-039: Visio groups nested deep', () => {
  for (const depth of [40, 50, 60, 200]) {
    test(`${depth} levels of groups keep their shapes and text`, async () => {
      let inner = vshape({ id: 9000, style: STYLE.NORMAL, box: [0.5, 0.5, 1, 1], sections: relRect(), text: `deep ${depth}\n` })
      for (let i = 0; i < depth; i++) inner = vshape({ id: 1000 + i, type: 'Group', style: STYLE.NORMAL, box: [1, 1, 2, 2], cells: { FillPattern: 0, LinePattern: 0 }, shapes: [inner] })
      const b = new VisioBuilder()
      b.page('Page-1', inner, {})
      const eng = new Engine(() => {})
      const info = await eng.open(new Blob([await b.zip()]), { name: 'nested.vsdx' })
      const slide = (await eng.call('pptx.slide', [0])) as { els: SlideEl[] }
      const texts = slide.els.flatMap((e) => (e.t === 'shape' && e.text ? [e.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n')] : []))
      expect(info.pptx!.count).toBe(1)
      expect(texts).toEqual([`deep ${depth}`])
    })
  }
})

describe('P3-033 (converters): different even-page headers and footers', () => {
  const evenRef = async (read: (path: string) => Promise<string>, kind: 'header' | 'footer') => {
    const doc = await read('word/document.xml')
    const rid = new RegExp(`<w:${kind}Reference w:type="even" r:id="([^"]+)"`).exec(doc)?.[1]
    expect(rid).toBeDefined()
    const rels = await read('word/_rels/document.xml.rels')
    const t = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`).exec(rels)!
    return read(`word/${t[1] ?? t[2]}`)
  }

  test('.doc: the facing-pages flag of the DOP and the even stories', async () => {
    const d = new DocBuilder()
    d.facingPages = true
    d.p('Body text')
    d.headers(['EVEN HEADER', 'ODD HEADER', 'EVEN FOOTER', 'ODD FOOTER', undefined, undefined])
    const zip = await ZipArchive.open(new BlobSource(await docToDocx(Cfb.open(d.build()))))
    const read = async (p: string) => (zip.has(p) ? zip.text(p) : '')
    expect(await read('word/settings.xml')).toContain('<w:evenAndOddHeaders/>')
    expect(await evenRef(read, 'header')).toContain('EVEN HEADER')
    expect(await evenRef(read, 'footer')).toContain('EVEN FOOTER')
    // without the flag Word shows the odd ones on every page: no even references
    const e = new DocBuilder()
    e.p('Body text')
    e.headers(['EVEN HEADER', 'ODD HEADER', 'EVEN FOOTER', 'ODD FOOTER', undefined, undefined])
    const z2 = await ZipArchive.open(new BlobSource(await docToDocx(Cfb.open(e.build()))))
    expect(await z2.text('word/settings.xml')).not.toContain('evenAndOddHeaders')
    expect(await z2.text('word/document.xml')).not.toContain('w:type="even"')
  })

  test('.rtf: \\facingp with \\headerl / \\headerr / \\footerl / \\footerr', async () => {
    const rtf = String.raw`{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}\facingp\sectd{\headerl\pard EVEN HEADER\par}{\headerr\pard ODD HEADER\par}{\footerl\pard EVEN FOOTER\par}{\footerr\pard ODD FOOTER\par} body\par}`
    const conv = await rtfToDocx(Uint8Array.from(rtf, (c) => c.charCodeAt(0) & 0xff))
    const zip = await ZipArchive.open(new BlobSource(conv))
    const read = async (p: string) => (zip.has(p) ? zip.text(p) : '')
    expect(await read('word/settings.xml')).toContain('<w:evenAndOddHeaders/>')
    expect(await evenRef(read, 'header')).toContain('EVEN HEADER')
    expect(await evenRef(read, 'footer')).toContain('EVEN FOOTER')
  })

  test('.odt: style:header-left / style:footer-left of the master page', async () => {
    const files = new Map<string, string>([
      ['mimetype', 'application/vnd.oasis.opendocument.text'],
      ['content.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles/><office:body><office:text><text:p>body</text:p></office:text></office:body></office:document-content>`],
      ['styles.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3"><office:styles/><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>ODD HEADER</text:p></style:header><style:header-left><text:p>EVEN HEADER</text:p></style:header-left><style:footer><text:p>ODD FOOTER</text:p></style:footer><style:footer-left><text:p>EVEN FOOTER</text:p></style:footer-left></style:master-page></office:master-styles></office:document-styles>`],
    ])
    const bytes = new Uint8Array(await (await writeZip(files, 'application/vnd.oasis.opendocument.text')).arrayBuffer())
    const conv = await convertOdfZip(await ZipArchive.open(new BlobSource(new Blob([bytes]))))
    const zip = await ZipArchive.open(new BlobSource(conv!.file))
    const read = async (p: string) => (zip.has(p) ? zip.text(p) : '')
    expect(await read('word/settings.xml')).toContain('<w:evenAndOddHeaders/>')
    expect(await evenRef(read, 'header')).toContain('EVEN HEADER')
    expect(await evenRef(read, 'footer')).toContain('EVEN FOOTER')
  })
})
