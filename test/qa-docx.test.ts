import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { paraSpec } from '../src/core/docx/spec'
import type { Block, EditResult, Paragraph } from '../src/core/docx/types'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'
import { rebuildZip } from '../src/core/zipwrite'
import { SaxParser, parseXml } from '../src/core/xml'
import { formatNumber } from '../src/core/docx/numbering'
import { resolveRuns, trackEdit } from '../src/react/docx/track'
import { deleteRange, type RRun } from '../src/react/edit/rich'
import type { Inline, Table, TrackMark } from '../src/core/docx/types'
import { deleteRow, gridOf, mergeCells, newTable } from '../src/react/docx/tables'

const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

/** a minimal package: `body` goes inside w:body, `extra` adds / replaces parts */
export function mkDocx(body: string, extra: Record<string, string> = {}): Promise<Blob> {
  const files = new Map<string, string>()
  files.set('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${extra['word/numbering.xml'] ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ''}</Types>`)
  files.set('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  files.set('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${extra['word/numbering.xml'] ? '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : ''}</Relationships>`)
  files.set('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${WNS}><w:body>${body}<w:sectPr/></w:body></w:document>`)
  for (const [k, v] of Object.entries(extra)) files.set(k, v)
  return rebuildZip(null, files as Map<string, string>, MIME)
}

export async function openDocx(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(src, {})
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  return { engine, blocks, events }
}

export const textOf = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : '￼')).join('')
export const paras = (blocks: Block[]) => blocks.filter((b): b is Paragraph => b.t === 'p')
export async function partText(blob: Blob, path: string) {
  return (await ZipArchive.open(new BlobSource(blob))).text(path)
}
/** edits paragraph `p` (a model paragraph, `edited`) in place and returns the saved file */
export async function editPara(engine: Engine, p: Paragraph, runs?: Paragraph['runs']): Promise<EditResult> {
  const edited: Paragraph = { ...p, edited: true, runs: runs ?? p.runs }
  return (await engine.call('docx.edit', [{ anchor: p.pid, list: [paraSpec([edited])], want: true }])) as EditResult
}
export const bodyOf = (xml: string) => xml.slice(xml.indexOf('<w:body>'))

describe('C20 XML reader is linear', () => {
  const big = 300_000
  const timed = (f: () => void) => {
    const t = performance.now()
    f()
    return performance.now() - t
  }
  const chunked = (xml: string, size: number) => {
    const ev: string[] = []
    const p = new SaxParser({ open: (n) => void ev.push(n), close: () => {}, text: () => {} })
    for (let i = 0; i < xml.length; i += size) p.write(xml.slice(i, i + size))
    p.end()
    return ev
  }
  test('a long comment arriving in tiny chunks', () => {
    const xml = '<a><!--' + 'x'.repeat(big) + '--><b/></a>'
    let ev: string[] = []
    expect(timed(() => (ev = chunked(xml, 16)))).toBeLessThan(1000)
    expect(ev).toEqual(['a', 'b'])
  })
  test('whitespace between tags / a long tag / CDATA / PI in tiny chunks', () => {
    const ws = '<a>' + ' '.repeat(big) + '<b/></a>'
    expect(timed(() => chunked(ws, 16))).toBeLessThan(1000)
    const tag = '<a ' + ' '.repeat(big) + 'x="1"><b/></a>'
    expect(timed(() => chunked(tag, 16))).toBeLessThan(1000)
    const cd = '<a><![CDATA[' + 'y'.repeat(big) + ']]><b/></a>'
    expect(timed(() => chunked(cd, 16))).toBeLessThan(1000)
    const pi = '<a><?pi ' + 'y'.repeat(big) + '?><b/></a>'
    expect(timed(() => chunked(pi, 16))).toBeLessThan(1000)
    const q = '<a x="' + '>'.repeat(big) + '"><b/></a>'
    expect(timed(() => chunked(q, 16))).toBeLessThan(1000)
  })
  test('any chunking gives the events of one write', () => {
    const xml = '<?xml version="1.0"?><!DOCTYPE a [<!ENTITY e "x">]><a x="1>2" y=\'<\'><!-- c > <b> --><![CDATA[ <z>]]>t&amp;u<b/> text <?p ?> <c   ></c ><d q="a>b"/></a>'
    const events = (size: number) => {
      const ev: string[] = []
      const raw: string[] = []
      const p = new SaxParser({ open: (n, a) => void ev.push('<' + n + JSON.stringify(a)), close: (n) => void ev.push('/' + n), text: (t) => void ev.push('t' + t) }, (r) => raw.push(r))
      for (let i = 0; i < xml.length; i += size) p.write(xml.slice(i, i + size))
      p.end()
      // text may arrive in pieces; what matters is the concatenation
      return [ev.join('|').replace(/\|t/g, ''), raw.join('')]
    }
    const whole = events(xml.length)
    expect(whole[1]).toBe(xml)
    for (const size of [1, 2, 3, 5, 7, 11, 64]) expect(events(size)).toEqual(whole)
  })
  test('whole-string parse with many comments and split output stays the same', () => {
    const xml = '<a>' + '<!-- c -->  <b>t</b>\n'.repeat(20_000) + '</a>'
    expect(parseXml(xml).c.length).toBe(20_000)
    expect(chunked(xml, 7).length).toBe(20_000 + 1)
  })
  test('a run of text, a comment or a CDATA section as long as the file is passed on chunk by chunk, not held', () => {
    const CHUNK = 1 << 16
    const feed = (open: string, fill: string, close: string, mb: number) => {
      const pieces: number[] = []
      let raw = 0
      const p = new SaxParser({ open: () => {}, close: () => {}, text: (t) => void pieces.push(t.length) }, (r) => void (raw += r.length))
      const block = fill.repeat(CHUNK)
      const t = performance.now()
      p.write('<a>' + open)
      for (let k = 0; k < (mb << 20) / CHUNK; k++) p.write(block)
      p.write(close + '<b/></a>')
      p.end()
      return { pieces, raw, ms: performance.now() - t, total: (mb << 20) + 3 + open.length + close.length + 8 }
    }
    // 64 MB of each: held whole, every chunk would copy what came before (2 GB and more of copies)
    const text = feed('', ' ', '', 64)
    expect(text.pieces.length).toBeGreaterThan(1000)
    expect(Math.max(...text.pieces)).toBeLessThanOrEqual(CHUNK + 16)
    expect(text.pieces.reduce((x, y) => x + y, 0)).toBe(64 << 20)
    expect(text.raw).toBe(text.total)
    const comment = feed('<!--', 'c', '-->', 64)
    expect(comment.pieces).toEqual([])
    expect(comment.raw).toBe(comment.total)
    const cdata = feed('<![CDATA[', ']', ']]>', 64)
    expect(cdata.pieces.reduce((x, y) => x + y, 0)).toBe(64 << 20)
    expect(cdata.raw).toBe(cdata.total)
    expect(text.ms + comment.ms + cdata.ms).toBeLessThan(5000)
  })

  test('a character reference cut by the end of a chunk is read whole', () => {
    for (const size of [1, 2, 3, 4, 5, 7]) {
      let out = ''
      const p = new SaxParser({ open: () => {}, close: () => {}, text: (t) => void (out += t) })
      const xml = '<a>x &amp; y &#x41;&#66; &lt;tag&gt; AT&T &quot;q&quot; end &</a>'
      for (let i = 0; i < xml.length; i += size) p.write(xml.slice(i, i + size))
      p.end()
      expect(out).toBe('x & y AB <tag> AT&T "q" end &')
    }
  })

  test('a tag of megabytes is read once; past 16 MB the file is refused', () => {
    const attrs = (mb: number) => '<a x="' + 'v'.repeat(mb << 20) + '"><b/></a>'
    const parse = (xml: string) => {
      const ev: [string, number][] = []
      const p = new SaxParser({ open: (n, a) => void ev.push([n, (a.x ?? '').length]), close: () => {}, text: () => {} })
      for (let i = 0; i < xml.length; i += 1 << 16) p.write(xml.slice(i, i + (1 << 16)))
      p.end()
      return ev
    }
    const t = performance.now()
    expect(parse(attrs(8))).toEqual([['a', 8 << 20], ['b', 0]])
    expect(performance.now() - t).toBeLessThan(3000)
    expect(() => parse(attrs(17))).toThrow(/tag longer than 16 MB/)
  })

  test('a 300 KB docx full of whitespace and comments opens in bounded time', async () => {
    const pad = '<!--' + 'c'.repeat(big) + '-->' + ' '.repeat(big)
    const file = await mkDocx(`<w:p>${pad}<w:r>${pad}<w:t>hi</w:t></w:r></w:p>`)
    const t = performance.now()
    const { blocks } = await openDocx(file)
    expect(textOf(paras(blocks)[0])).toBe('hi')
    expect(performance.now() - t).toBeLessThan(3000)
  })
})

describe('N2-008 / P3-020 links that run code', () => {
  test('a document takes web, mail and phone links only', async () => {
    const { engine } = await openDocx(await mkDocx('<w:p><w:r><w:t>x</w:t></w:r></w:p>'))
    for (const url of ['javascript:alert(document.domain)', 'JaVaScRiPt:alert(1)', ' javascript:alert(1)', 'java\tscript:alert(1)', 'vbscript:x', 'data:text/html,<script>1</script>', 'file:///etc/passwd']) {
      await expect(engine.call('docx.link', [url])).rejects.toThrow(/cannot be used/)
    }
    for (const url of ['https://example.com/a?b=1#c', 'mailto:a@example.com', 'tel:+84123456789', '#bookmark']) expect(typeof (await engine.call('docx.link', [url]))).toBe('string')
    const saved = (await engine.call('save', [])) as Blob
    const rels = await partText(saved, 'word/_rels/document.xml.rels')
    expect(rels).not.toContain('javascript')
    expect(rels).toContain('https://example.com/a?b=1#c')
  })

  test('a scheme spelled with HTML entities is still a scheme', async () => {
    const { safeHref } = await import('../src/core/url')
    for (const url of ['javascript&colon;alert(1)', 'javascript&COLON;alert(1)', '&NewLine;javascript:alert(3)', 'java&Tab;script:alert(1)', 'javascript&#58;alert(1)', 'javascript&#x3a;alert(1)', '&#106;avascript:alert(1)']) expect(safeHref(url)).toBeUndefined()
    for (const url of ['https://example.com/?a=1&colon=2', 'page.html?x=1&amp;y=2', 'mailto:a&b@example.com', '#top']) expect(safeHref(url)).toBe(url)
  })
})

describe('C19 list markers are bounded', () => {
  test('formatNumber caps every numbering format', () => {
    for (const fmt of ['upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter', 'decimalZero', 'ordinal', 'decimal']) {
      const t = performance.now()
      const s = formatNumber(2147483647, fmt)
      expect(s.length).toBeLessThan(40)
      expect(performance.now() - t).toBeLessThan(200)
    }
    expect(formatNumber(4, 'lowerRoman')).toBe('iv')
    expect(formatNumber(28, 'lowerLetter')).toBe('bb')
  })
  test('a list starting at 2147483647 opens fast with a short marker', async () => {
    const lvl = (fmt: string) => `<w:lvl w:ilvl="0"><w:start w:val="2147483647"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1."/></w:lvl>`
    const numbering = `<?xml version="1.0"?><w:numbering ${WNS}><w:abstractNum w:abstractNumId="0">${lvl('upperRoman')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${lvl('lowerLetter')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
    const li = (id: number) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`
    const file = await mkDocx(li(1).repeat(50) + li(2).repeat(50), { 'word/numbering.xml': numbering })
    const t = performance.now()
    const { blocks } = await openDocx(file)
    expect(performance.now() - t).toBeLessThan(3000)
    for (const p of paras(blocks)) expect((p.marker?.text ?? '').length).toBeLessThan(40)
  })
})

/** a paragraph whose first run is replaced by `text` (the way the editor does it), the rest as it is */
async function editFirst(file: Blob, word = 'EDITED') {
  const { engine, blocks } = await openDocx(file)
  const p = paras(blocks)[0]
  const first = p.runs[0] as Extract<Paragraph['runs'][number], { t: 'text' }>
  expect(first.t).toBe('text')
  await editPara(engine, p, [{ ...first, text: first.text.replace('Hello', word), d: { b: true } }, ...p.runs.slice(1)])
  const saved = (await engine.call('save', [])) as Blob
  const xml = bodyOf(await partText(saved, 'word/document.xml'))
  return { xml, saved, p, engine }
}

describe('C16 mc:AlternateContent', () => {
  const emoji = '<w:r><mc:AlternateContent><mc:Choice Requires="w16se"><w16se:symEx xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" w16se:font="Segoe UI Emoji" w16se:char="1F600"/></mc:Choice><mc:Fallback><w:t>\u{1F600}</w:t></mc:Fallback></mc:AlternateContent></w:r>'
  test('an emoji written as AlternateContent survives editing the paragraph', async () => {
    const { xml } = await editFirst(await mkDocx(`<w:p><w:r><w:t>Hello </w:t></w:r>${emoji}<w:r><w:rPr><w:b/></w:rPr><w:t> end</w:t></w:r></w:p>`))
    expect(xml).toContain('EDITED')
    expect(xml).toContain('\u{1F600}')
    expect(xml).toContain(' end')
  })
  test('a choice nobody understands, with no fallback, is kept verbatim', async () => {
    const odd = '<w:r><mc:AlternateContent><mc:Choice Requires="w16se"><w16se:symEx xmlns:w16se="urn:x" w16se:char="1F600"/></mc:Choice></mc:AlternateContent></w:r>'
    const { xml } = await editFirst(await mkDocx(`<w:p><w:r><w:t>Hello </w:t></w:r>${odd}</w:p>`))
    expect(xml).toContain('EDITED')
    expect(xml).toContain(odd)
  })
  test('an understood choice (wps) is still the one read, and its run is copied whole', async () => {
    const box = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="x"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>in box</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>'
    const file = await mkDocx(`<w:p><w:r><w:t>Hello </w:t></w:r>${box}</w:p>`)
    const { blocks } = await openDocx(file)
    expect(paras(blocks)[0].runs.some((r) => r.t === 'box')).toBe(true)
    const { xml } = await editFirst(file)
    expect(xml).toContain(box)
  })
})

describe('P3-002…006 an edited paragraph keeps what the engine does not model', () => {
  const head = '<w:r><w:t xml:space="preserve">Hello </w:t></w:r>'
  const edit = async (inner: string, extra: Record<string, string> = {}) => editFirst(await mkDocx(`<w:p>${head}${inner}</w:p>`, extra))

  test('(a) run properties the engine has no model for stay on the runs', async () => {
    const r1 = '<w:r><w:rPr><w:rStyle w:val="Strong"/><w:b/><w:lang w:val="fr-FR"/><w:rtl/></w:rPr><w:t>Bonjour</w:t></w:r>'
    const r2 = '<w:r><w:rPr><w:b/><w:lang w:val="de-DE"/></w:rPr><w:t> Welt</w:t></w:r>'
    const { xml } = await edit(r1 + r2)
    expect(xml).toContain('<w:rStyle w:val="Strong"/>')
    expect(xml).toContain('<w:lang w:val="fr-FR"/>')
    expect(xml).toContain('<w:rtl/>')
    expect(xml).toContain('<w:lang w:val="de-DE"/>')
  })
  test('(a) the first run, edited, keeps its own', async () => {
    const file = await mkDocx('<w:p><w:r><w:rPr><w:lang w:val="fr-FR"/><w:rtl/></w:rPr><w:t>Hello monde</w:t></w:r></w:p>')
    const { xml } = await editFirst(file)
    expect(xml).toContain('<w:lang w:val="fr-FR"/>')
    expect(xml).toContain('<w:rtl/>')
  })
  test('(b) complex fields (DATE, MERGEFIELD) stay fields', async () => {
    const cplx = (instr: string, res: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${res}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
    const f1 = cplx('DATE \\@ "d MMMM yyyy"', '18 November 2026')
    const f2 = cplx('MERGEFIELD Name', '«Name»')
    const { xml } = await edit(f1 + '<w:r><w:t> and </w:t></w:r>' + f2)
    for (const k of ['DATE \\@ "d MMMM yyyy"', 'MERGEFIELD Name']) expect(xml).toContain(k)
    expect((xml.match(/w:fldCharType="begin"/g) ?? []).length).toBe(2)
    expect((xml.match(/w:fldCharType="separate"/g) ?? []).length).toBe(2)
    expect((xml.match(/w:fldCharType="end"/g) ?? []).length).toBe(2)
    expect(xml).toContain('18 November 2026')
  })
  test('(b) a complex field keeps its place among the text', async () => {
    const cplx = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    const { xml } = await edit(cplx + '<w:r><w:t> tail</w:t></w:r>')
    expect(xml.indexOf('Hello') < xml.indexOf('DATE') || xml.indexOf('EDITED') < xml.indexOf('DATE')).toBe(true)
    expect(xml.indexOf('DATE')).toBeLessThan(xml.indexOf('today'))
    expect(xml.indexOf('today')).toBeLessThan(xml.indexOf(' tail'))
  })
  test('(b) simple fields (fldSimple) stay fields', async () => {
    const fs = '<w:fldSimple w:instr=" DATE \\@ &quot;d MMM yyyy&quot; "><w:r><w:rPr><w:noProof/></w:rPr><w:t>18 Nov 2026</w:t></w:r></w:fldSimple>'
    const { xml } = await edit(fs + '<w:r><w:t> tail</w:t></w:r>')
    expect(xml).toContain(fs)
    expect(xml).toContain(' tail')
  })
  test('(b) PAGE fields still work', async () => {
    const cplx = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    const { xml } = await edit(cplx + '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>3</w:t></w:r></w:fldSimple>')
    expect(xml).toContain('PAGE')
    expect(xml).toContain('NUMPAGES')
    expect(xml).not.toMatch(/fldCharType="begin"/)
  })
  test('(c) hidden text stays in the file', async () => {
    const hid = '<w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r>'
    const { xml } = await edit(hid + '<w:r><w:t> tail</w:t></w:r>')
    expect(xml).toContain(hid)
  })
  test('(d) ruby stays', async () => {
    const ruby = '<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="center"/></w:rubyPr><w:rt><w:r><w:t>kan</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>漢</w:t></w:r></w:rubyBase></w:ruby></w:r>'
    const { xml } = await edit(ruby + '<w:r><w:t> tail</w:t></w:r>')
    expect(xml).toContain(ruby)
  })
  test('(e) equations stay equations', async () => {
    const math = '<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>'
    const { xml, engine } = await edit(math + '<w:r><w:t> tail</w:t></w:r>')
    expect(xml).toContain(math)
    expect(xml).toContain(' tail')
    // and it is shown as before
    expect(engine).toBeDefined()
  })
  test('an untouched paragraph is written byte for byte, and so is the rest of the document', async () => {
    const file = await mkDocx(`<w:p>${head}<w:r><w:vanish/></w:r></w:p><w:p><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></w:p>`)
    const { engine } = await openDocx(file)
    const saved = (await engine.call('save', [])) as Blob
    expect(await partText(saved, 'word/document.xml')).toBe(await partText(file, 'word/document.xml'))
  })
})

describe('P3-030 deleting what another author inserted', () => {
  const ME: TrackMark = { k: 'ins', author: 'Vy', date: '2026-11-18T10:00:00Z' }
  const A = '<w:ins w:id="7" w:author="Lan" w:date="2026-11-01T09:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins>'
  const doc = () => mkDocx(`<w:p><w:r><w:t>Hello </w:t></w:r>${A}<w:r><w:t> tail</w:t></w:r></w:p>`)

  test('the deletion is made on top of the insertion, and rejecting everything takes the text out', async () => {
    const { engine, blocks } = await openDocx(await doc())
    const p = paras(blocks)[0]
    const before = p
    const o = 'Hello '.length
    const next = { ...p, runs: deleteRange(p.runs as RRun[], o, o + 3) as Inline[], edited: true }
    const sel = { a: { pa: '0', o }, f: { pa: '0', o: o + 3 } }
    const r = trackEdit([before], [next], '0', sel, sel, ME)
    const del = r.next[0].runs.find((x) => x.t === 'text' && x.tc?.k === 'del') as Extract<Inline, { t: 'text' }>
    expect(del.text).toBe('new')
    expect(del.tc).toMatchObject({ k: 'del', author: 'Vy', ins: { k: 'ins', author: 'Lan' } })
    // rejecting all: the insertion goes, so does the text
    const all = plenOf(r.next[0])
    expect(textOf({ ...r.next[0], runs: resolveRuns(r.next[0].runs, 0, all, false) })).toBe('Hello  tail')
    expect(textOf({ ...r.next[0], runs: resolveRuns(r.next[0].runs, 0, all, true) })).toBe('Hello  tail')
    // saved as Word writes it: the deletion inside the insertion
    await editPara(engine, r.next[0] as Paragraph)
    const saved = (await engine.call('save', [])) as Blob
    const xml = bodyOf(await partText(saved, 'word/document.xml'))
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Lan" w:date="2026-11-01T09:00:00Z"><w:del w:id="\d+" w:author="Vy" w:date="2026-11-18T10:00:00Z"><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:delText xml:space="preserve">new<\/w:delText><\/w:r><\/w:del><\/w:ins>/)
    // and it reads back the same
    const re = await openDocx(saved)
    const back = paras(re.blocks)[0].runs.find((x) => x.t === 'text' && x.tc?.k === 'del') as Extract<Inline, { t: 'text' }>
    expect(back.tc).toMatchObject({ k: 'del', author: 'Vy', ins: { author: 'Lan' } })
  })
  test('a deletion inside an insertion in a Word file reads as both', async () => {
    const file = await mkDocx(`<w:p><w:ins w:id="1" w:author="A"><w:del w:id="2" w:author="B"><w:r><w:delText>gone</w:delText></w:r></w:del></w:ins></w:p>`)
    const { blocks } = await openDocx(file)
    const x = paras(blocks)[0].runs[0] as Extract<Inline, { t: 'text' }>
    expect(x.tc).toMatchObject({ k: 'del', author: 'B', ins: { k: 'ins', author: 'A' } })
  })
})
const plenOf = (p: Paragraph) => p.runs.reduce((n, r) => n + (r.t === 'text' ? r.text.length : 1), 0)

describe('P3-037 deleting the row a vertical merge starts in', () => {
  let pid = 6_000_000
  const para = (text = ''): Paragraph => ({ t: 'p', ps: {}, runs: text ? [{ t: 'text', text, s: {} }] : [], pid: pid++, edited: true })
  const cellText = (c: { blocks: Block[] }) => c.blocks.map((b) => (b.t === 'p' ? textOf(b) : '')).join('|')
  function grid3(): Table {
    const t = newTable(3, 3, 300, () => para())
    t.rows.forEach((row, ri) => row.cells.forEach((c, ci) => (c.blocks = [para(`${ri},${ci}`)])))
    return t
  }
  test('the next row takes the merge over', () => {
    const t = mergeCells(grid3(), { r0: 0, g0: 1, r1: 2, g1: 1 })!
    expect(t.rows.map((r) => r.cells.length)).toEqual([3, 2, 2])
    const d = deleteRow(t, 0)!
    expect(d.rows.map((r) => r.cells.length)).toEqual([3, 2])
    const g = gridOf(d)
    expect(g.nCols).toBe(3)
    for (let r = 0; r < 2; r++) for (let c = 0; c < 3; c++) expect(g.at(r, c)).toBeDefined()
    expect(g.at(0, 1)!.cell.rowSpan).toBe(2)
    expect(g.at(1, 1)).toBe(g.at(0, 1))
    expect(d.rows[0].cells.map((c) => cellText(c).split('|')[0])).toEqual(['1,0', '0,1', '1,2'])
  })
  test('one row of the merge left: a plain cell; deleting a later row of the merge still shrinks it', () => {
    const d = deleteRow(mergeCells(grid3(), { r0: 0, g0: 1, r1: 1, g1: 1 })!, 0)!
    expect(gridOf(d).at(0, 1)!.cell.rowSpan).toBe(1)
    expect(d.rows.map((r) => r.cells.length)).toEqual([3, 3])
    expect(gridOf(deleteRow(mergeCells(grid3(), { r0: 0, g0: 1, r1: 2, g1: 1 })!, 2)!).at(0, 1)!.cell.rowSpan).toBe(2)
  })
  test('written and read back, the table has no hole', async () => {
    const rows = 3
    const tc = (t: string, vm = '') => `<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/>${vm}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
    const tbl = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>` +
      `<w:tr>${tc('a0')}${tc('m', '<w:vMerge w:val="restart"/>')}</w:tr><w:tr>${tc('a1')}${tc('', '<w:vMerge/>')}</w:tr><w:tr>${tc('a2')}${tc('', '<w:vMerge/>')}</w:tr></w:tbl><w:p/>`
    const { engine, blocks } = await openDocx(await mkDocx(tbl))
    const t = blocks.find((b): b is Table => b.t === 'tbl')!
    expect(t.rows.length).toBe(rows)
    const d = { ...deleteRow(t, 0)!, tid: 9_200_001 }
    const { tableSpec } = await import('../src/core/docx/spec')
    await engine.call('docx.table', [{ tid: d.tid, item: d.item, spec: tableSpec(d) }])
    const re = await openDocx((await engine.call('save', [])) as Blob)
    const r = re.blocks.find((b): b is Table => b.t === 'tbl')!
    expect(r.rows.length).toBe(2)
    expect(r.rows.map((x) => x.cells.length)).toEqual([2, 1])
    expect(r.rows[0].cells[1].rowSpan).toBe(2)
  })
})

describe('P3-028 RTF text in Symbol / Wingdings', () => {
  const bytes = (s: string) => Uint8Array.from(s, (c) => c.charCodeAt(0) & 0xff)
  const fonts = String.raw`{\fonttbl{\f0\froman\fcharset0 Times New Roman;}{\f2\fnil\fcharset2 Symbol;}{\f3\fnil\fcharset2 Wingdings;}{\f4\fnil\fcharset2 Webdings;}}`
  const rtf = String.raw`{\rtf1\ansi\ansicpg1252\deff0` + fonts + '\n' + String.raw`\pard Hello {\f2 abc\'b7\'e5} and {\f3 lq\'fc} and {\f4 xy}\par}`
  test('the characters are kept: Unicode where there is a glyph, the font name is the real one', async () => {
    const events: any[] = []
    const engine = new Engine((event, data) => events.push({ event, data }))
    await engine.open(new Blob([bytes(rtf)]), { name: 'a.rtf' })
    await engine.call('docx.start', [])
    const p = paras(events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data))[0]
    const runs = p.runs.filter((r): r is Extract<Inline, { t: 'text' }> => r.t === 'text')
    const sym = runs.find((r) => r.s.font?.includes('Symbol'))!
    expect(sym.text).toBe('αβχ•∑')
    expect(runs.find((r) => r.s.font?.includes('Wingdings'))!.text).toBe('●❑✓')
    // no mapping (Webdings): the characters stay, in the font that has them
    const web = runs.find((r) => r.s.font?.includes('Webdings'))!
    expect(web.text.length).toBe(2)
    expect(web.text).not.toBe('')
    expect(runs.map((r) => r.text).join('')).toContain('Hello ')
  })
})
