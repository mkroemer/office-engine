import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import { paraSpec } from '../src/core/docx/spec'
import type { Block, Paragraph } from '../src/core/docx/types'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { bareDocx, para } from './docx-builder'
import { DocBuilder, sprm } from './doc-builder'

async function open(src: Blob) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(src, {})
  await engine.call('docx.start', [])
  const paras = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data as Block[]).filter((b): b is Paragraph => b.t === 'p')
  return { engine, paras }
}

const text = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : '')).join('')
const keep = (p: Paragraph) => ({ keepLines: p.ps.keepLines, widowControl: p.ps.widowControl })

const STYLES =
  '<w:style w:type="paragraph" w:styleId="Kept"><w:name w:val="Kept"/><w:basedOn w:val="Normal"/><w:pPr><w:keepLines/><w:widowControl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="KeptChild"><w:name w:val="Kept child"/><w:basedOn w:val="Kept"/></w:style>'

describe('keepLines / widowControl', () => {
  test('read from the paragraph, from its style (and the styles it is based on), from the document defaults', async () => {
    const body = [
      para('direct', '<w:keepLines/><w:widowControl/>'),
      para('direct off', '<w:keepLines w:val="0"/><w:widowControl w:val="false"/>'),
      para('style', '<w:pStyle w:val="Kept"/>'),
      para('based on', '<w:pStyle w:val="KeptChild"/>'),
      para('style overridden', '<w:pStyle w:val="Kept"/><w:keepLines w:val="0"/><w:widowControl/>'),
      para('plain'),
    ].join('')
    const { paras } = await open(await bareDocx(body, { styles: STYLES }))
    expect(paras.map(text)).toEqual(['direct', 'direct off', 'style', 'based on', 'style overridden', 'plain'])
    expect(paras.map(keep)).toEqual([
      { keepLines: true, widowControl: true },
      { keepLines: false, widowControl: false },
      { keepLines: true, widowControl: false },
      { keepLines: true, widowControl: false },
      { keepLines: false, widowControl: true },
      // nothing says: widow control is Word's default (the view treats undefined as on)
      { keepLines: undefined, widowControl: undefined },
    ])

    const defaults = await open(await bareDocx(para('a') + para('b', '<w:widowControl/><w:keepLines w:val="0"/>'), { pPrDefault: '<w:keepLines/><w:widowControl w:val="0"/>' }))
    expect(defaults.paras.map(keep)).toEqual([
      { keepLines: true, widowControl: false },
      { keepLines: false, widowControl: true },
    ])
  })

  test('kept when an edited paragraph is written back, and after a save and reopen', async () => {
    const body = para('first', '<w:keepLines/><w:widowControl w:val="0"/>') + para('second', '<w:pStyle w:val="Kept"/>') + para('third')
    const { engine, paras } = await open(await bareDocx(body, { styles: STYLES }))
    // the text of each paragraph changes, and the first is split in two (Enter): the new one is made from it
    const [p0, p1] = paras
    const a = { ...p0, runs: [{ t: 'text' as const, text: 'first edited', s: {} }], edited: true }
    const b = { ...p0, runs: [{ t: 'text' as const, text: 'new', s: {} }], pid: 1_000_000_000, from: p0.pid, orig: [], edited: true }
    await engine.call('docx.edit', [{ anchor: p0.pid, list: [paraSpec([a]), paraSpec([b])] }])
    await engine.call('docx.edit', [{ anchor: p1.pid, list: [paraSpec([{ ...p1, runs: [{ t: 'text', text: 'second edited', s: {} }], edited: true }])] }])
    const saved = (await engine.call('save', [])) as Blob
    const xml = await (await ZipArchive.open(new BlobSource(saved))).text('word/document.xml')
    expect(xml.match(/<w:keepLines\/>/g)?.length).toBe(2)
    expect(xml.match(/<w:widowControl w:val="0"\/>/g)?.length).toBe(2)
    const re = await open(saved)
    expect(re.paras.map(text)).toEqual(['first edited', 'new', 'second edited', 'third'])
    expect(re.paras.map(keep)).toEqual([
      { keepLines: true, widowControl: false },
      { keepLines: true, widowControl: false },
      { keepLines: true, widowControl: false },
      { keepLines: undefined, widowControl: undefined },
    ])
  })

  test('RTF: widow control with \\widctlpar, or for the document with \\widowctrl; Word 97: unless the paragraph turns it off', async () => {
    const head = String.raw`{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}`
    const body = String.raw`\pard plain\par\pard\widctlpar kept\par\pard\nowidctlpar off\par}`
    const r = await open(new Blob([head + '\n' + body], { type: 'application/rtf' }))
    expect(r.paras.map(text)).toEqual(['plain', 'kept', 'off'])
    expect(r.paras.map((p) => p.ps.widowControl)).toEqual([false, true, false])
    // \widowctrl: on for the document, so a plain paragraph has it (nothing written: on is the default) and only \nowidctlpar says no
    const all = await open(new Blob([head + String.raw`\widowctrl` + '\n' + body], { type: 'application/rtf' }))
    expect(all.paras.map((p) => p.ps.widowControl)).toEqual([undefined, undefined, false])
    // Word 97: a paragraph has it unless its sprm says 0 (fWidowControl defaults to 1)
    const d = new DocBuilder()
    d.p('plain')
    d.p('off', 0, sprm(0x2431, 0))
    d.p('kept', 0, sprm(0x2431, 1))
    const w = await open(new Blob([d.build() as Uint8Array<ArrayBuffer>]))
    expect(w.paras.map(text)).toEqual(['plain', 'off', 'kept'])
    expect(w.paras.map((p) => p.ps.widowControl)).toEqual([undefined, false, undefined])
  })
})

describe('a run that holds more than text', () => {
  const drawing = (id: number) =>
    `<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="952500" cy="952500"/><wp:docPr id="${id}" name="Chart ${id}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId9${id}"/></a:graphicData></a:graphic></wp:inline></w:drawing>`
  const count = (xml: string, re: RegExp) => xml.match(re)?.length ?? 0
  const texts = (xml: string) => [...xml.matchAll(/<w:t\b[^>]*>([^<]*)</g)].map((m) => m[1])
  /** every paragraph written again (as an edit does), `runs` deciding what is left of each; the saved document.xml */
  async function rewritten(body: string, runs: (p: Paragraph) => Paragraph['runs'] = (p) => p.runs) {
    const { engine, paras } = await open(await bareDocx(body))
    for (const p of paras) await engine.call('docx.edit', [{ anchor: p.pid, list: [paraSpec([{ ...p, runs: runs(p), edited: true }])] }])
    const zip = await ZipArchive.open(new BlobSource((await engine.call('save', [])) as Blob))
    return zip.text('word/document.xml')
  }

  test('two objects in one run are written once each', async () => {
    const xml = await rewritten(`<w:p><w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr>${drawing(1)}${drawing(2)}</w:r><w:r><w:t>Z</w:t></w:r></w:p>`)
    expect([...xml.matchAll(/docPr id="(\d+)"/g)].map((m) => m[1])).toEqual(['1', '2'])
    // each in a run of its own, with the properties of the run they came from
    expect(count(xml, /<w:r><w:rPr><w:noProof\/><\/w:rPr><w:drawing>/g)).toBe(2)
    expect(texts(xml)).toEqual(['A', 'Z'])
  })

  test('text next to an object in one run is not written twice', async () => {
    const xml = await rewritten(`<w:p><w:r><w:t>before</w:t>${drawing(1)}<w:t>after</w:t></w:r></w:p>`)
    expect(count(xml, /<w:drawing>/g)).toBe(1)
    expect(texts(xml)).toEqual(['before', 'after'])
  })

  test('a footnote and an endnote reference in one run stay one of each', async () => {
    const xml = await rewritten('<w:p><w:r><w:t>T</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/><w:endnoteReference w:id="1"/></w:r></w:p>')
    expect(count(xml, /<w:footnoteReference /g)).toBe(1)
    expect(count(xml, /<w:endnoteReference /g)).toBe(1)
  })

  test('one of two objects of a run is deleted: only that one goes', async () => {
    const body = `<w:p><w:r><w:t>A</w:t></w:r><w:r>${drawing(1)}${drawing(2)}</w:r></w:p>`
    const xml = await rewritten(body, (p) => p.runs.filter((_, i) => i !== 2))
    expect([...xml.matchAll(/docPr id="(\d+)"/g)].map((m) => m[1])).toEqual(['1'])
    // an object alone in its run is written back as it was, to the byte
    const alone = `<w:p><w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr>${drawing(7)}</w:r></w:p>`
    expect(await rewritten(alone)).toContain(`<w:r><w:rPr><w:noProof/></w:rPr>${drawing(7)}</w:r>`)
  })
})
