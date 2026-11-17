import { expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { DocxEditor, type DocState } from '../src/react/docx/DocxEditor'
import { PptxEditor } from '../src/react/pptx/PptxEditor'
import { applyDocx, applyPptx, cellPatch, cellFormat } from '../src/react/edit/format'

const file = async (kind: string) => new Blob([await Bun.file(`${import.meta.dir}/../playground/public/samples/sample.${kind}`).arrayBuffer()])
const invalid = [NaN, Infinity, -Infinity, 0, -1]
const invalidConverted = [Number.MAX_VALUE, Number.MIN_VALUE, Number.MAX_SAFE_INTEGER]
const xml = async (client: DocumentClient, path: string) => (await ZipArchive.open(new BlobSource(await client.call('save')))).text(path)
async function documentEditor() {
  const client = await DocumentClient.open(await file('docx'), { worker: false })
  const st: DocState = { blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }
  client.on('docx:blocks', (b) => st.blocks.push(...b))
  await client.call('docx.start')
  const ed = new DocxEditor(client, st, () => {})
  const pa = String(st.blocks.findIndex((b) => b.t === 'p' && b.runs.some((r) => r.t === 'text' && r.text.includes('This document'))))
  ed.input.lastSel = { a: { pa, o: 0 }, f: { pa, o: 4 } }
  return { client, ed }
}
async function presentationEditor() {
  const client = await DocumentClient.open(await file('pptx'), { worker: false })
  const model = await client.call('pptx.slide', 0)
  const shape = model.els.find((e) => e.t === 'shape' && e.text)
  if (!shape || shape.t !== 'shape' || !shape.text) throw new Error('No text shape')
  const ed = new PptxEditor(client)
  const sid = shape.sid ?? model.els.indexOf(shape)
  ed.begin(0, sid, shape.text)
  ed.input.lastSel = { a: { pa: '0', o: 0 }, f: { pa: '0', o: 4 } }
  return { client, ed, body: shape.text, sid }
}
const settled = async (ed: DocxEditor) => { while (ed.busy) await new Promise((r) => setTimeout(r, 5)) }

test('DOCX: invalid public sizes leave model, history and saved XML untouched; fractional size survives', async () => {
  const { client, ed } = await documentEditor()
  try {
    const before = await xml(client, 'word/document.xml')
    const state = ed.state()
    for (const pt of [...invalid, ...invalidConverted, 0.24]) applyDocx(ed, { op: 'size', pt })
    expect(ed.state()).toEqual(state)
    expect(ed.canUndo).toBe(false)
    expect(await xml(client, 'word/document.xml')).toBe(before)
    applyDocx(ed, { op: 'size', pt: 12.5 })
    await settled(ed)
    const saved = await xml(client, 'word/document.xml')
    expect(saved).toContain('w:sz w:val="25"')
    expect(saved).not.toMatch(/NaN|Infinity/)
    expect(ed.state()?.sz).toBe(12.5)
  } finally { client.destroy() }
})

test('PPTX: invalid public sizes leave model, history and saved XML untouched; fractional size survives', async () => {
  const { client, ed } = await presentationEditor()
  try {
    const before = await xml(client, 'ppt/slides/slide1.xml')
    const text = ed.text()
    for (const pt of [...invalid, ...invalidConverted, 0.004]) applyPptx(ed, { op: 'size', pt })
    expect(ed.text()).toEqual(text)
    // If an invalid command had entered history, undo/redo would also alter this body.
    ed.undo()
    ed.redo()
    expect(ed.text()).toEqual(text)
    ed.commit()
    expect(await xml(client, 'ppt/slides/slide1.xml')).toBe(before)
    applyPptx(ed, { op: 'size', pt: 12.5 })
    ed.commit()
    const saved = await xml(client, 'ppt/slides/slide1.xml')
    expect(saved).toContain('sz="1250"')
    expect(saved).not.toMatch(/NaN|Infinity/)
  } finally { ed.end(); client.destroy() }
})

test('XLSX: invalid public sizes produce no mutation; fractional size survives saved XML', async () => {
  const client = await DocumentClient.open(await file('xlsx'), { worker: false })
  try {
    const loaded = new Promise<void>((r) => client.on('xlsx:progress', (p) => p.sheet === 0 && p.done && r()))
    await client.call('xlsx.load', 0)
    await loaded
    const before = await xml(client, 'xl/styles.xml')
    const now = cellFormat({}, {})
    for (const pt of invalid) expect(cellPatch({ op: 'size', pt }, now)).toBeNull()
    expect(await xml(client, 'xl/styles.xml')).toBe(before)
    await client.call('xlsx.format', 0, { r0: 1, c0: 1, r1: 1, c1: 1 }, cellPatch({ op: 'size', pt: 12.5 }, now)!)
    const saved = await xml(client, 'xl/styles.xml')
    expect(saved).toContain('val="12.5"')
    expect(saved).not.toMatch(/NaN|Infinity/)
  } finally { client.destroy() }
})

test('DOCX lists toggle collectively across a mixed selection; repetition removes the list', async () => {
  const { client, ed } = await documentEditor()
  try {
    const first = ed.input.lastSel!.a.pa
    const second = ed.next(first)!
    ed.input.lastSel = { a: { pa: first, o: 0 }, f: { pa: first, o: 1 } }
    applyDocx(ed, { op: 'list', list: 'bullet' })
    await settled(ed)
    ed.input.lastSel = { a: { pa: first, o: 0 }, f: { pa: second, o: 1 } }
    applyDocx(ed, { op: 'list', list: 'bullet' })
    await settled(ed)
    expect([ed.para(first)?.lst, ed.para(second)?.lst]).toEqual(['bullet', 'bullet'])
    applyDocx(ed, { op: 'list', list: 'bullet' })
    await settled(ed)
    expect([ed.para(first)?.lst, ed.para(second)?.lst]).toEqual([undefined, undefined])
    // Read the converted model directly: subscription refresh after conversion depends on inherited #19.
    expect(ed.state()?.lst).toBeUndefined()
  } finally { client.destroy() }
})

test('PPTX mixed lists toggle from the first paragraph and repetition removes the list', async () => {
  const { client, ed, body, sid } = await presentationEditor()
  try {
    ed.begin(0, sid, { ...body, paras: [{ ...body.paras[0], buKind: 'bullet' }, { ...body.paras[0], buKind: 'none' }] })
    ed.input.lastSel = { a: { pa: '0', o: 0 }, f: { pa: '1', o: 1 } }
    applyPptx(ed, { op: 'list', list: 'bullet' })
    expect(ed.paras.map((p) => p.pp?.bullet)).toEqual(['none', 'none'])
    applyPptx(ed, { op: 'list', list: 'number' })
    expect(ed.paras.map((p) => p.pp?.bullet)).toEqual(['number', 'number'])
    applyPptx(ed, { op: 'list', list: 'number' })
    expect(ed.paras.map((p) => p.pp?.bullet)).toEqual(['none', 'none'])
  } finally { ed.end(); client.destroy() }
})
