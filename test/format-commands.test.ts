import { expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { DocxEditor, type DocState } from '../src/react/docx/DocxEditor'
import { PptxEditor } from '../src/react/pptx/PptxEditor'
import { applyDocx, applyPptx, cellPatch, cellFormat } from '../src/react/edit/format'

const file = async (kind: string) => new Blob([await Bun.file(`${import.meta.dir}/../playground/public/samples/sample.${kind}`).arrayBuffer()])
const invalid = [NaN, Infinity, -Infinity, 0, -1, 0.25, 0.5, 400.1, 1000]
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
    for (const pt of [1, 400, 12.26]) {
      applyDocx(ed, { op: 'size', pt }); await settled(ed)
      expect(await xml(client, 'word/document.xml')).toContain(`w:sz w:val="${Math.round(pt * 2)}"`)
    }
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
    for (const pt of [1, 400, 12.26]) {
      applyPptx(ed, { op: 'size', pt }); ed.commit()
      expect(await xml(client, 'ppt/slides/slide1.xml')).toContain(`sz="${Math.round(pt * 2) * 50}"`)
    }
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
    for (const pt of [1, 400, 12.26]) {
      await client.call('xlsx.format', 0, { r0: 1, c0: 1, r1: 1, c1: 1 }, cellPatch({ op: 'size', pt }, now)!)
      expect(await xml(client, 'xl/styles.xml')).toContain(`val="${Math.round(pt * 2) / 2}"`)
    }
    expect(saved).not.toMatch(/NaN|Infinity/)
  } finally { client.destroy() }
})

test('DOCX list setters cover mixed selection, are idempotent and explicitly remove lists', async () => {
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
    expect([ed.para(first)?.lst, ed.para(second)?.lst]).toEqual(['bullet', 'bullet'])
    applyDocx(ed, { op: 'list', list: null })
    await settled(ed)
    expect([ed.para(first)?.lst, ed.para(second)?.lst]).toEqual([undefined, undefined])
    // Read the converted model directly: subscription refresh after conversion depends on inherited #19.
    expect(ed.state()?.lst).toBeUndefined()
  } finally { client.destroy() }
})

test('PPTX list setters cover mixed selection, are idempotent and explicitly remove lists', async () => {
  const { client, ed, body, sid } = await presentationEditor()
  try {
    ed.begin(0, sid, { ...body, paras: [{ ...body.paras[0], buKind: 'bullet' }, { ...body.paras[0], buKind: 'none' }] })
    ed.input.lastSel = { a: { pa: '0', o: 0 }, f: { pa: '1', o: 1 } }
    applyPptx(ed, { op: 'list', list: 'bullet' })
    expect(ed.paras.map((p) => p.pp?.bullet)).toEqual(['bullet', 'bullet'])
    applyPptx(ed, { op: 'list', list: 'number' })
    expect(ed.paras.map((p) => p.pp?.bullet)).toEqual(['number', 'number'])
    applyPptx(ed, { op: 'list', list: 'number' })
    expect(ed.paras.map((p) => p.pp?.bullet)).toEqual(['number', 'number'])
    applyPptx(ed, { op: 'list', list: null })
    expect(ed.paras.map((p) => p.pp?.bullet)).toEqual(['none', 'none'])
    ed.toggleBullet('bullet')
    ed.toggleBullet('bullet')
    expect(ed.paras.map((p) => p.pp?.bullet)).toEqual(['none', 'none'])
  } finally { ed.end(); client.destroy() }
})

const invalidCommands = [
  { op: 'bold' }, { op: 'toggle', mark: 'unknown' }, { op: 'align', align: 'unknown' }, { op: 'list', list: 'unknown' },
  { op: 'font', name: '' }, { op: 'font', name: '   ' }, { op: 'font', name: 'Arial\u0000' },
  ...['red', 'rgb(255,0,0)', '#ff000080', '#1234', 'var(--brand)', 'not-a-color'].flatMap((color) => [{ op: 'color', color }, { op: 'highlight', color }]),
] as unknown as import('../src/react/controller').FormatCommand[]

for (const kind of ['docx', 'pptx'] as const) {
  test(`${kind}: invalid JS commands, empty fonts and unsupported CSS colors preserve model, dirty, history and XML`, async () => {
    const { client, ed } = kind === 'docx' ? await documentEditor() : await presentationEditor()
    const apply = (c: import('../src/react/controller').FormatCommand) => kind === 'docx' ? applyDocx(ed as DocxEditor, c) : applyPptx(ed as PptxEditor, c)
    const path = kind === 'docx' ? 'word/document.xml' : 'ppt/slides/slide1.xml'
    try {
      let dirty = 0
      ed.onDirty = () => dirty++
      const before = await xml(client, path)
      const model = JSON.stringify(ed instanceof DocxEditor ? ed.state() : ed.text())
      for (const c of invalidCommands) apply(c)
      expect(dirty).toBe(0)
      expect(JSON.stringify(ed instanceof DocxEditor ? ed.state() : ed.text())).toBe(model)
      if (ed instanceof DocxEditor) expect(ed.canUndo).toBe(false)
      else ed.commit()
      ed.undo(); ed.redo()
      expect(await xml(client, path)).toBe(before)
      // Opaque alpha normalizes to writable RGB; shorthand highlight is also expanded.
      apply({ op: 'color', color: '#123456ff' })
      apply({ op: 'highlight', color: '#abc' })
      apply({ op: 'font', name: ' Arial ' })
      if (ed instanceof DocxEditor) await settled(ed)
      else ed.commit()
      const colored = await xml(client, path)
      expect(colored).toMatch(/123456/i)
      expect(colored).toMatch(/AABBCC/i)
      expect(colored).toContain('Arial')
      // Rejected colors cannot erase an existing color or add an undo step.
      const dirtyBefore = dirty
      for (const c of invalidCommands) apply(c)
      expect(dirty).toBe(dirtyBefore)
      if (ed instanceof PptxEditor) ed.commit()
      expect(await xml(client, path)).toBe(colored)
      ed.undo()
      if (ed instanceof DocxEditor) await settled(ed)
      else ed.commit()
      expect(await xml(client, path)).toMatch(/123456/i)
    } finally { if (ed instanceof PptxEditor) ed.end(); client.destroy() }
  })
}

test('XLSX: invalid commands cannot reach worker history; opaque colors survive actual saved XML', async () => {
  const client = await DocumentClient.open(await file('xlsx'), { worker: false })
  try {
    const loaded = new Promise<void>((r) => client.on('xlsx:progress', (p) => p.sheet === 0 && p.done && r()))
    await client.call('xlsx.load', 0); await loaded
    const range = { r0: 1, c0: 1, r1: 1, c1: 1 }
    const now = cellFormat({}, {})
    for (const c of invalidCommands) expect(cellPatch(c, now)).toBeNull()
    for (const c of [{ op: 'color', color: '#123456ff' }, { op: 'highlight', color: '#abc' }] as const) {
      await client.call('xlsx.format', 0, range, cellPatch(c, now)!)
    }
    const saved = await xml(client, 'xl/styles.xml')
    expect(saved).toMatch(/FF123456/i)
    expect(saved).toMatch(/FFAABBCC/i)
    for (const c of invalidCommands) {
      const patch = cellPatch(c, now)
      if (patch) await client.call('xlsx.format', 0, range, patch)
    }
    expect(await xml(client, 'xl/styles.xml')).toBe(saved)
    await client.call('xlsx.undo')
    expect(await xml(client, 'xl/styles.xml')).toMatch(/FF123456/i)
  } finally { client.destroy() }
})
