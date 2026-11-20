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

  test('RTF and Word 97 files: widow control only where the file has it (their paragraphs have none unless said)', async () => {
    const rtf = String.raw`{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}` + '\n' + String.raw`\pard plain\par\pard\widctlpar kept\par\pard\nowidctlpar off\par}`
    const r = await open(new Blob([rtf], { type: 'application/rtf' }))
    expect(r.paras.map(text)).toEqual(['plain', 'kept', 'off'])
    expect(r.paras.map((p) => p.ps.widowControl)).toEqual([false, true, false])
    const d = new DocBuilder()
    d.p('plain')
    d.p('kept', 0, sprm(0x2431, 1))
    const w = await open(new Blob([d.build() as Uint8Array<ArrayBuffer>]))
    expect(w.paras.map(text)).toEqual(['plain', 'kept'])
    expect(w.paras.map((p) => p.ps.widowControl)).toEqual([false, true])
  })
})
