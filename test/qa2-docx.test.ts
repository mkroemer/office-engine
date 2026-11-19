import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { paraSpec } from '../src/core/docx/spec'
import type { Block, DocxMeta, EditResult, Inline, Paragraph } from '../src/core/docx/types'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'
import { parseXml, parseXmlStream } from '../src/core/xml'
import { formatNumber } from '../src/core/docx/numbering'
import { pageHF } from '../src/react/docx/paginate'

const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** a minimal package: `body` goes inside w:body, `extra` adds / replaces parts (a Uint8Array is written as it is) */
function mkDocx(body: string, extra: Record<string, string | Uint8Array> = {}, rels = ''): Promise<Blob> {
  const files = new Map<string, string | Uint8Array>()
  files.set('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${extra['word/numbering.xml'] ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ''}</Types>`)
  files.set('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)
  files.set('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${extra['word/numbering.xml'] ? `<Relationship Id="rIdN" Type="${REL}/numbering" Target="numbering.xml"/>` : ''}${extra['word/styles.xml'] ? `<Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/>` : ''}${rels}</Relationships>`)
  files.set('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${WNS}><w:body>${body}<w:sectPr/></w:body></w:document>`)
  for (const [k, v] of Object.entries(extra)) files.set(k, v)
  return rebuildZip(null, files as Map<string, string>, MIME)
}

async function openDocx(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(src, {})
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  const meta = events.filter((e) => e.event === 'docx:meta').pop()?.data as DocxMeta
  return { engine, blocks, events, meta }
}

const paras = (blocks: Block[]) => blocks.filter((b): b is Paragraph => b.t === 'p')
/** the text of the text runs only */
const shown = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' || r.t === 'ref' ? r.text : '')).join('')
async function partText(blob: Blob, path: string) {
  return (await ZipArchive.open(new BlobSource(blob))).text(path)
}
const bodyOf = (xml: string) => xml.slice(xml.indexOf('<w:body>'))
async function editPara(engine: Engine, p: Paragraph, runs?: Paragraph['runs']): Promise<EditResult> {
  const edited: Paragraph = { ...p, edited: true, runs: runs ?? p.runs }
  return (await engine.call('docx.edit', [{ anchor: p.pid, list: [paraSpec([edited])], want: true }])) as EditResult
}
/** the first text run of the first paragraph is changed (the way the editor does it) and the file saved */
async function editFirst(file: Blob, from = 'Hello', to = 'EDITED') {
  const { engine, blocks } = await openDocx(file)
  const p = paras(blocks)[0]
  const i = p.runs.findIndex((r) => r.t === 'text')
  const first = p.runs[i] as Extract<Inline, { t: 'text' }>
  const runs = [...p.runs]
  runs[i] = { ...first, text: first.text.replace(from, to) }
  await editPara(engine, p, runs)
  const saved = (await engine.call('save', [])) as Blob
  // (a run of text is written with xml:space="preserve": not what these tests look at)
  return { xml: bodyOf(await partText(saved, 'word/document.xml')).replace(/ xml:space="preserve"/g, ''), saved, p, engine }
}
const head = '<w:r><w:t xml:space="preserve">Hello </w:t></w:r>'
const edit = async (inner: string, extra: Record<string, string> = {}) => editFirst(await mkDocx(`<w:p>${head}${inner}</w:p>`, extra))
const fld = (instr: string, res: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${res}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`

describe('P3-003 a nested field shows the result of the outermost one', () => {
  const inner = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> MERGEFIELD Gender </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«Gender»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
  const ifField = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF </w:instrText></w:r>${inner}<w:r><w:instrText> = "M" "Mr." "Ms." </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Mr.</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
  test('the instruction, with its nested fields, is not shown', async () => {
    const { blocks } = await openDocx(await mkDocx(`<w:p><w:r><w:t xml:space="preserve">Dear </w:t></w:r>${ifField}<w:r><w:t xml:space="preserve"> Smith</w:t></w:r></w:p>`))
    expect(shown(paras(blocks)[0])).toBe('Dear Mr. Smith')
  })
  test('a field in the instruction of another is not an object, and survives editing', async () => {
    const page = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    const calc = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> = </w:instrText></w:r>${page}<w:r><w:instrText> + 1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>5</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
    const file = await mkDocx(`<w:p>${head}${calc}</w:p>`)
    const { blocks } = await openDocx(file)
    const p = paras(blocks)[0]
    expect(shown(p)).toBe('Hello 5')
    expect(p.runs.some((r) => r.t === 'field')).toBe(false)
    const { xml } = await editFirst(file)
    expect(xml).toContain('<w:instrText> PAGE </w:instrText>')
    expect(xml).toContain('<w:instrText> + 1 </w:instrText>')
    expect((xml.match(/fldCharType="begin"/g) ?? []).length).toBe(2)
    expect((xml.match(/fldCharType="end"/g) ?? []).length).toBe(2)
  })
  test('a field in the result of another is shown', async () => {
    const f = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF 1 = 1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">a </w:t></w:r>${fld('DATE', 'today')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
    const { blocks } = await openDocx(await mkDocx(`<w:p>${f}</w:p>`))
    expect(shown(paras(blocks)[0])).toBe('a today')
  })
})

describe('P3-006 equations are shown in a linear form', () => {
  const r = (t: string) => `<m:r><m:t>${t}</m:t></m:r>`
  const shownMath = async (math: string) => {
    const { blocks } = await openDocx(await mkDocx(`<w:p><m:oMath>${math}</m:oMath></w:p>`))
    return shown(paras(blocks)[0])
  }
  test('fraction, superscript, subscript', async () => {
    expect(await shownMath(`<m:f><m:num>${r('a')}</m:num><m:den>${r('b')}</m:den></m:f>`)).toBe('a/b')
    expect(await shownMath(`${r('E=m')}<m:sSup><m:e>${r('c')}</m:e><m:sup>${r('2')}</m:sup></m:sSup>`)).toBe('E=mc²')
    expect(await shownMath(`<m:sSub><m:e>${r('x')}</m:e><m:sub>${r('1')}</m:sub></m:sSub>`)).toBe('x₁')
    expect(await shownMath(`<m:sSup><m:e>${r('x')}</m:e><m:sup>${r('a+b')}</m:sup></m:sSup>`)).toBe('x^(a+b)')
    expect(await shownMath(`<m:sSub><m:e>${r('x')}</m:e><m:sub>${r('ij')}</m:sub></m:sSub>`)).toBe('x_ij')
  })
  test('parentheses where a fraction part is more than one token', async () => {
    expect(await shownMath(`<m:f><m:num>${r('a+b')}</m:num><m:den>${r('c')}</m:den></m:f>`)).toBe('(a+b)/c')
    expect(await shownMath(`<m:f><m:num>${r('1')}</m:num><m:den>${r('x+y')}</m:den></m:f>`)).toBe('1/(x+y)')
  })
  test('radicals, n-ary operators, delimiters, functions', async () => {
    expect(await shownMath(`<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${r('x+1')}</m:e></m:rad>`)).toBe('√(x+1)')
    expect(await shownMath(`<m:rad><m:deg>${r('3')}</m:deg><m:e>${r('x')}</m:e></m:rad>`)).toBe('∛(x)')
    expect(await shownMath(`<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${r('i=1')}</m:sub><m:sup>${r('n')}</m:sup><m:e>${r('i')}</m:e></m:nary>`)).toBe('∑_(i=1)^n i')
    expect(await shownMath(`<m:nary><m:naryPr/><m:sub>${r('0')}</m:sub><m:sup>${r('1')}</m:sup><m:e>${r('f')}</m:e></m:nary>`)).toBe('∫_0^1 f')
    expect(await shownMath(`<m:d><m:e>${r('a')}</m:e><m:e>${r('b')}</m:e></m:d>`)).toBe('(a,b)')
    expect(await shownMath(`<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>${r('x')}</m:e></m:d>`)).toBe('[x]')
    expect(await shownMath(`<m:func><m:fName>${r('sin')}</m:fName><m:e>${r('x')}</m:e></m:func>`)).toBe('sin x')
  })
  test('still one object, kept as it is on edit', async () => {
    const math = `<m:oMath><m:f><m:num>${r('a')}</m:num><m:den>${r('b')}</m:den></m:f></m:oMath>`
    const { xml } = await edit(math)
    expect(xml).toContain(math)
  })
})

describe('P3-007 inline content controls, customXml and smartTag stay around their text', () => {
  const sdt = '<w:sdt><w:sdtPr><w:alias w:val="Agree"/><w:id w:val="42"/><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>'
  test('a checkbox is kept, with its text editable', async () => {
    const file = await mkDocx(`<w:p>${head}${sdt}<w:r><w:t> tail</w:t></w:r></w:p>`)
    const { blocks } = await openDocx(file)
    expect(shown(paras(blocks)[0])).toBe('Hello ☒ tail')
    const { xml } = await editFirst(file)
    expect(xml).toContain('EDITED ')
    expect(xml).toContain(sdt)
    expect(xml.indexOf('EDITED')).toBeLessThan(xml.indexOf('<w:sdt>'))
    expect(xml.indexOf('</w:sdt>')).toBeLessThan(xml.indexOf(' tail'))
  })
  test('the text inside a content control, edited, stays inside it', async () => {
    const file = await mkDocx(`<w:p><w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:r><w:t>Hello name</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> tail</w:t></w:r></w:p>`)
    const { xml } = await editFirst(file)
    expect(xml).toContain('<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:r><w:t>EDITED name</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> tail</w:t></w:r>')
  })
  test('smartTag and customXml wrappers, nested', async () => {
    const st = '<w:smartTag w:uri="urn:x" w:element="place"><w:smartTagPr><w:attr w:name="a" w:val="b"/></w:smartTagPr><w:r><w:t>Paris</w:t></w:r></w:smartTag>'
    const cx = '<w:customXml w:uri="urn:y" w:element="person"><w:customXmlPr><w:attr w:name="n" w:val="1"/></w:customXmlPr><w:r><w:t>Ann</w:t></w:r></w:customXml>'
    const { xml } = await edit(st + '<w:r><w:t> and </w:t></w:r>' + cx)
    expect(xml).toContain(st)
    expect(xml).toContain(cx)
    expect(xml).toContain('EDITED')
  })
  test('one wrapper around several runs is written once', async () => {
    const st = '<w:smartTag w:uri="urn:x" w:element="place"><w:r><w:t>Pa</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>ris</w:t></w:r></w:smartTag>'
    const { xml } = await edit(st)
    expect((xml.match(/<w:smartTag /g) ?? []).length).toBe(1)
    expect((xml.match(/<\/w:smartTag>/g) ?? []).length).toBe(1)
  })
})

describe('P3-008 breaks, tabs and direction of an edited paragraph', () => {
  test('a column break and a textWrapping clear stay what they are', async () => {
    const col = '<w:r><w:br w:type="column"/></w:r>'
    const clr = '<w:r><w:br w:type="textWrapping" w:clear="all"/></w:r>'
    const { xml, p } = await edit(col + '<w:r><w:t>x</w:t></w:r>' + clr)
    expect(xml).toContain(col)
    expect(xml).toContain(clr)
    expect(shown(p)).toBe('Hello \nx\n')
  })
  test('a plain break is still a plain break', async () => {
    const { xml } = await edit('<w:r><w:t>a</w:t><w:br/><w:t>b</w:t></w:r>')
    expect(xml).toContain('a</w:t><w:br/><w:t>b</w:t>')
  })
  test('an absolute position tab stays', async () => {
    const pt = '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>'
    const { xml, p } = await edit(pt + '<w:r><w:t>right</w:t></w:r>')
    expect(xml).toContain(pt)
    expect(shown(p)).toBe('Hello \tright')
  })
  test('w:dir and w:bdo stay around their runs', async () => {
    const dir = '<w:dir w:val="rtl"><w:r><w:t>שלום</w:t></w:r></w:dir>'
    const bdo = '<w:bdo w:val="ltr"><w:r><w:t>abc</w:t></w:r></w:bdo>'
    const { xml } = await edit(dir + bdo)
    expect(xml).toContain(dir)
    expect(xml).toContain(bdo)
  })
})

describe('P3-009 tracked changes of the file stay when the paragraph is edited', () => {
  test('rPrChange stays on its run, as the last child of rPr', async () => {
    const run = '<w:r><w:rPr><w:b/><w:rPrChange w:id="3" w:author="Lan" w:date="2026-11-01T09:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>bold now</w:t></w:r>'
    const { xml } = await edit(run)
    expect(xml).toContain('<w:rPr><w:b/><w:rPrChange w:id="3" w:author="Lan" w:date="2026-11-01T09:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>bold now</w:t>')
  })
  test('pPrChange stays', async () => {
    const ppr = '<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="4" w:author="Lan"><w:pPr/></w:pPrChange></w:pPr>'
    const { xml } = await editFirst(await mkDocx(`<w:p>${ppr}${head}</w:p>`))
    expect(xml).toContain(ppr)
  })
  test('moveFrom / moveTo stay moves, with their ranges', async () => {
    const from = '<w:moveFromRangeStart w:id="1" w:name="move1"/><w:moveFrom w:id="5" w:author="Lan" w:date="2026-11-01T09:00:00Z"><w:r><w:delText>moved</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/>'
    const to = '<w:moveToRangeStart w:id="2" w:name="move1"/><w:moveTo w:id="6" w:author="Lan" w:date="2026-11-01T09:00:00Z"><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="2"/>'
    const { xml } = await edit(from + to)
    expect(xml).toContain('<w:moveFrom ')
    expect(xml).toContain('<w:moveTo ')
    expect(xml).not.toContain('<w:ins ')
    expect(xml).not.toContain('<w:del ')
    expect(xml).toContain('>moved</w:delText>')
    for (const k of ['moveFromRangeStart w:id="1" w:name="move1"', 'moveFromRangeEnd w:id="1"', 'moveToRangeStart w:id="2" w:name="move1"', 'moveToRangeEnd w:id="2"']) expect(xml).toContain(k)
  })
})

describe('P3-011 note references show their ordinal', () => {
  test('footnotes 1, 2 and endnotes i, ii whatever their w:id', async () => {
    const fn = (id: number) => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`
    const en = (id: number) => `<w:r><w:endnoteReference w:id="${id}"/></w:r>`
    const file = await mkDocx(`<w:p><w:r><w:t>a</w:t></w:r>${fn(5)}</w:p><w:p>${fn(9)}${en(1)}${en(4)}</w:p>`)
    const { blocks } = await openDocx(file)
    const refs = paras(blocks).flatMap((p) => p.runs.filter((r) => r.t === 'ref').map((r) => (r as Extract<Inline, { t: 'ref' }>).text))
    expect(refs).toEqual(['1', '2', 'i', 'ii'])
  })
  test('still the same after the paragraph is edited', async () => {
    const fn = (id: number) => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`
    const file = await mkDocx(`<w:p>${head}${fn(5)}</w:p><w:p>${head}${fn(9)}</w:p>`)
    const { engine, blocks } = await openDocx(file)
    const p = paras(blocks)[1]
    const res = await editPara(engine, p, p.runs.map((r) => (r.t === 'text' ? { ...r, text: 'x ' } : r)))
    const got = JSON.stringify(res)
    expect(got).toContain('"text":"2"')
  })
})

describe('P3-012 list number formats', () => {
  test.each([
    ['chineseCounting', 3, '三'], ['chineseCounting', 11, '十一'], ['chineseCounting', 20, '二十'], ['chineseCounting', 105, '一百零五'],
    ['ideographTraditional', 3, '丙'], ['ideographZodiac', 3, '寅'], ['decimalEnclosedCircle', 3, '③'], ['decimalEnclosedCircle', 25, '25'],
    ['aiueo', 3, 'ｳ'], ['aiueoFullWidth', 3, 'ウ'], ['iroha', 3, 'ﾊ'], ['irohaFullWidth', 3, 'ハ'], ['koreanDigital', 3, '삼'], ['koreanCounting', 3, '삼'], ['ganada', 3, '다'], ['chosung', 3, 'ㄷ'],
    ['hebrew1', 3, 'ג'], ['hebrew1', 15, 'טו'], ['hebrew2', 3, 'ג'], ['arabicAlpha', 3, 'ت'], ['arabicAbjad', 3, 'ج'], ['thaiLetters', 3, 'ค'], ['thaiNumbers', 12, '๑๒'],
    ['russianLower', 3, 'в'], ['russianUpper', 3, 'В'], ['cardinalText', 3, 'Three'], ['cardinalText', 21, 'Twenty-one'], ['ordinalText', 3, 'Third'], ['ordinalText', 21, 'Twenty-first'],
    ['decimalZero', 5, '05'], ['decimalZero', 12, '12'], ['numberInDash', 3, '-3-'], ['decimalFullWidth', 12, '１２'], ['upperLetter', 28, 'BB'], ['hex', 255, 'FF'],
  ])('%s %d', (fmt, n, want) => expect(formatNumber(n as number, fmt as string)).toBe(want))
  test('every format is bounded', () => {
    for (const fmt of ['chineseCounting', 'chineseLegalSimplified', 'ideographTraditional', 'decimalEnclosedCircle', 'aiueo', 'iroha', 'koreanDigital', 'koreanCounting', 'ganada', 'chosung', 'hebrew1', 'hebrew2', 'arabicAlpha', 'arabicAbjad', 'thaiLetters', 'thaiNumbers', 'russianLower', 'cardinalText', 'ordinalText', 'numberInDash', 'decimalFullWidth', 'hex', 'hindiVowels', 'vietnameseCounting']) {
      for (const n of [1, 27, 999, 2600, 9999, 99999, 2147483647]) {
        const t = performance.now()
        const s = formatNumber(n, fmt)
        expect(s.length).toBeLessThan(120)
        expect(performance.now() - t).toBeLessThan(200)
      }
    }
  })
})

/** the whole body of the file written as UTF-16 */
function utf16(xml: string, le: boolean, bom: boolean): Uint8Array {
  const out = new Uint8Array((xml.length + (bom ? 1 : 0)) * 2)
  const v = new DataView(out.buffer)
  let o = 0
  if (bom) v.setUint16((o++) * 2, 0xfeff, le)
  for (let i = 0; i < xml.length; i++) v.setUint16((o++) * 2, xml.charCodeAt(i), le)
  return out
}

describe('P3-018 a part in UTF-16', () => {
  const doc = (decl: string) => `${decl}<w:document ${WNS}><w:body><w:p><w:r><w:t>utf16 text Tiếng Việt 😀</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
  test.each([
    ['LE with a BOM', true, true, '<?xml version="1.0" encoding="UTF-16"?>'],
    ['BE with a BOM', false, true, '<?xml version="1.0" encoding="UTF-16"?>'],
    ['LE without a BOM', true, false, '<?xml version="1.0" encoding="UTF-16"?>'],
    ['BE without a BOM', false, false, '<?xml version="1.0" encoding="utf-16"?>'],
    ['LE with a BOM, no declaration', true, true, ''],
  ])('%s', async (_n, le, bom, decl) => {
    const file = await mkDocx('', { 'word/document.xml': utf16(doc(decl), le, bom) })
    const { blocks } = await openDocx(file)
    expect(shown(paras(blocks)[0])).toBe('utf16 text Tiếng Việt 😀')
  })
  test('edited and saved: UTF-8 with a declaration that says so', async () => {
    const file = await mkDocx('', { 'word/document.xml': utf16(doc('<?xml version="1.0" encoding="UTF-16"?>'), true, true) })
    const { saved } = await editFirst(file, 'utf16', 'EDITED')
    const xml = await partText(saved, 'word/document.xml')
    expect(xml).toContain('EDITED text Tiếng Việt 😀')
    expect(xml).not.toMatch(/encoding="UTF-16"/i)
    const { blocks } = await openDocx(saved)
    expect(shown(paras(blocks)[0])).toBe('EDITED text Tiếng Việt 😀')
  })
  test('another part (styles) and the stream parser, in tiny chunks', async () => {
    const xml = '<?xml version="1.0" encoding="UTF-16"?><a x="é"><b>漢字😀</b></a>'
    for (const [le, bom] of [[true, true], [false, true], [true, false], [false, false]] as const) {
      const bytes = utf16(xml, le, bom)
      for (const size of [1, 3, 7, 1000]) {
        let got = ''
        await parseXmlStream(
          new ReadableStream({
            start(c) {
              for (let i = 0; i < bytes.length; i += size) c.enqueue(bytes.slice(i, i + size))
              c.close()
            },
          }),
          { open: (n, a) => void (got += `<${n}${a.x ?? ''}>`), close: () => {}, text: (t) => void (got += t) },
        )
        expect(got).toBe('<aé><b>漢字😀')
      }
    }
    const file = await mkDocx('<w:p><w:r><w:t>x</w:t></w:r></w:p>', { 'word/styles.xml': utf16(`<?xml version="1.0" encoding="UTF-16"?><w:styles ${WNS}><w:style w:type="paragraph" w:styleId="Weird"><w:name w:val="Weird"/></w:style></w:styles>`, true, true) })
    const z = await ZipArchive.open(new BlobSource(file))
    expect(await z.text('word/styles.xml')).toContain('w:styleId="Weird"')
  })
  test('UTF-8 stays UTF-8 (BOM, declared single-byte encodings)', async () => {
    const file = await mkDocx('', { 'word/document.xml': new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(doc('<?xml version="1.0" encoding="UTF-8"?>'))]) })
    expect(shown(paras((await openDocx(file)).blocks)[0])).toBe('utf16 text Tiếng Việt 😀')
    const latin = Uint8Array.from(Buffer.from('<?xml version="1.0" encoding="ISO-8859-1"?><w:document ' + WNS + '><w:body><w:p><w:r><w:t>café</w:t></w:r></w:p></w:body></w:document>', 'latin1'))
    expect(shown(paras((await openDocx(await mkDocx('', { 'word/document.xml': latin }))).blocks)[0])).toBe('café')
  })
})

describe('P3-033 even pages take the even header and footer', () => {
  const hdr = (t: string) => `<?xml version="1.0"?><w:hdr ${WNS}><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:hdr>`
  const ftr = (t: string) => `<?xml version="1.0"?><w:ftr ${WNS}><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:ftr>`
  const sect = (titlePg: boolean) =>
    `<w:sectPr><w:headerReference w:type="default" r:id="h1"/><w:headerReference w:type="even" r:id="h2"/><w:headerReference w:type="first" r:id="h3"/><w:footerReference w:type="default" r:id="f1"/><w:footerReference w:type="even" r:id="f2"/>${titlePg ? '<w:titlePg/>' : ''}</w:sectPr>`
  const open = async (evenOdd: boolean, titlePg = true) => {
    const rels = ['h1:header1', 'h2:header2', 'h3:header3', 'f1:footer1', 'f2:footer2'].map((x) => `<Relationship Id="${x.split(':')[0]}" Type="${REL}/${x.includes(':header') ? 'header' : 'footer'}" Target="${x.split(':')[1]}.xml"/>`).join('')
    const document = `<?xml version="1.0"?><w:document ${WNS}><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p>${sect(titlePg)}</w:body></w:document>`
    const settings = `<?xml version="1.0"?><w:settings ${WNS}>${evenOdd ? '<w:evenAndOddHeaders/>' : ''}</w:settings>`
    const file = await mkDocx('', {
      'word/document.xml': document, 'word/settings.xml': settings,
      'word/header1.xml': hdr('DEFAULT H'), 'word/header2.xml': hdr('EVEN H'), 'word/header3.xml': hdr('FIRST H'), 'word/footer1.xml': ftr('DEFAULT F'), 'word/footer2.xml': ftr('EVEN F'),
    }, rels + `<Relationship Id="rIdSet" Type="${REL}/settings" Target="settings.xml"/>`)
    return (await openDocx(file)).meta
  }
  const text = (b?: Block[]) => (b?.[0] as Paragraph | undefined)?.runs.map((r) => (r.t === 'text' ? r.text : '')).join('')
  test('with evenAndOddHeaders: first, even, default (odd)', async () => {
    const meta = await open(true)
    expect(meta.evenOdd).toBe(true)
    const at = (i: number) => {
      const hf = pageHF(meta, i)
      return `${text(hf.header)}|${text(hf.footer)}`
    }
    expect([0, 1, 2, 3, 4, 5].map(at)).toEqual(['FIRST H|undefined', 'EVEN H|EVEN F', 'DEFAULT H|DEFAULT F', 'EVEN H|EVEN F', 'DEFAULT H|DEFAULT F', 'EVEN H|EVEN F'])
  })
  test('without it the even parts are not used', async () => {
    const meta = await open(false, false)
    expect(meta.evenOdd).toBeFalsy()
    expect([0, 1, 2].map((i) => text(pageHF(meta, i).header))).toEqual(['DEFAULT H', 'DEFAULT H', 'DEFAULT H'])
  })
})

describe('P3-035 / P3-036 list styles and legal numbering', () => {
  const lvl = (i: number, fmt: string, text: string, extra = '') => `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/>${extra}</w:lvl>`
  const li = (id: number, ilvl: number, t: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${id}"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`
  const marks = (blocks: Block[]) => paras(blocks).map((p) => p.marker?.text ?? '')
  test('numStyleLink takes the levels of the abstract list that has the styleLink', async () => {
    const numbering = `<?xml version="1.0"?><w:numbering ${WNS}>
      <w:abstractNum w:abstractNumId="0"><w:styleLink w:val="MyList"/>${lvl(0, 'upperLetter', '%1)')}${lvl(1, 'lowerRoman', '%2.')}</w:abstractNum>
      <w:abstractNum w:abstractNumId="1"><w:numStyleLink w:val="MyList"/></w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
    const styles = `<?xml version="1.0"?><w:styles ${WNS}><w:style w:type="numbering" w:styleId="MyList"><w:name w:val="MyList"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>`
    const file = await mkDocx(li(2, 0, 'A') + li(2, 1, 'A.i') + li(2, 0, 'B'), { 'word/numbering.xml': numbering, 'word/styles.xml': styles })
    expect(marks((await openDocx(file)).blocks)).toEqual(['A)', 'i.', 'B)'])
  })
  test('isLgl turns every number of the level into an Arabic one', async () => {
    const numbering = `<?xml version="1.0"?><w:numbering ${WNS}><w:abstractNum w:abstractNumId="0">${lvl(0, 'upperRoman', 'Chapter %1')}${lvl(1, 'lowerLetter', '%1.%2', '<w:isLgl/>')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
    const file = await mkDocx(li(1, 0, 'a') + li(1, 1, 'b') + li(1, 0, 'c') + li(1, 1, 'd'), { 'word/numbering.xml': numbering })
    expect(marks((await openDocx(file)).blocks)).toEqual(['Chapter I', '1.1', 'Chapter II', '2.1'])
  })
})

describe('P3-040 editing a header keeps the comment anchored in it', () => {
  test('range and reference are written again, around the text', async () => {
    const hdrXml = `<?xml version="1.0"?><w:hdr ${WNS}><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Company</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p></w:hdr>`
    const comments = `<?xml version="1.0"?><w:comments ${WNS}><w:comment w:id="0" w:author="Lan" w:date="2026-11-01T09:00:00Z" w:initials="L"><w:p><w:r><w:t>check this</w:t></w:r></w:p></w:comment></w:comments>`
    const rels = `<Relationship Id="h1" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rC" Type="${REL}/comments" Target="comments.xml"/>`
    const document = `<?xml version="1.0"?><w:document ${WNS}><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="h1"/></w:sectPr></w:body></w:document>`
    const file = await mkDocx('', { 'word/document.xml': document, 'word/header1.xml': hdrXml, 'word/comments.xml': comments }, rels)
    const { engine } = await openDocx(file)
    await engine.call('docx.hf', [{ parts: { header: [{ text: 'Company ABC' }] } }])
    const saved = (await engine.call('save', [])) as Blob
    const h = await partText(saved, 'word/header1.xml')
    expect(h).toMatch(/<w:commentRangeStart w:id="0"\/>/)
    expect(h).toMatch(/<w:commentRangeEnd w:id="0"\/>/)
    expect(h).toMatch(/<w:commentReference w:id="0"\/>/)
    expect(h.indexOf('commentRangeStart')).toBeLessThan(h.indexOf('Company ABC'))
    expect(h.indexOf('Company ABC')).toBeLessThan(h.indexOf('commentRangeEnd'))
    expect(await partText(saved, 'word/comments.xml')).toContain('check this')
  })
})

void parseXml
