/** QA round 2, presentations: speaker notes (P3-013), undo of slide list edits (N2-024), table cell text (N2-026 / N2-027). */
import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import type { DocComment } from '../src/core/comments'
import type { PresentationMeta, SlideEl, SlideModel } from '../src/core/pptx/types'
import { slidesToText } from '../src/react/file/text'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const bytes = (n: string) => new Uint8Array(readFileSync(S(n)))

async function open(src: Uint8Array, name = 'deck.pptx') {
  const engine = new Engine(() => {})
  const info = await engine.open(new Blob([src as BlobPart]), { name })
  const slide = (i: number) => engine.call('pptx.slide', [i]) as Promise<SlideModel>
  const all = async () => {
    const out: SlideModel[] = []
    for (let i = 0; i < info.pptx!.count; i++) out.push(await slide(i))
    return out
  }
  /** the slide list as the engine holds it now */
  const paths = () => info.pptx!.slides.map((s) => s.path)
  return { engine, info, slide, all, paths }
}

const saved = async (engine: Engine) => (await engine.call('save', [])) as Blob
async function parts(blob: Blob) {
  const zip = await ZipArchive.open(new BlobSource(blob))
  return { zip, text: (p: string) => zip.text(p) }
}
const titleOf = (m: SlideModel) => {
  const t = m.els.find((e) => e.t === 'shape' && e.text)
  return t && t.t === 'shape' ? t.text!.paras.map((p) => p.runs.map((r) => r.text).join('')).join('|') : ''
}
type Table = Extract<SlideEl, { t: 'table' }>
const tableOf = (m: SlideModel) => m.els.find((e): e is Table => e.t === 'table')!
const cellText = (t: Table, r: number, c: number) => t.rows[r].cells[c].text?.paras.map((p) => p.runs.map((x) => x.text).join('')).join('|')

/** sample.pptx with speaker notes on slide 2 (the way PowerPoint writes them) */
function withNotes(): Uint8Array {
  const z = unzipSync(bytes('sample.pptx'))
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
  z['ppt/notesSlides/notesSlide1.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' +
      '<a:p><a:r><a:rPr lang="en-US"/><a:t>Speaker notes for slide 2</a:t></a:r></a:p><a:p><a:r><a:t>Second </a:t></a:r><a:r><a:t>line</a:t></a:r><a:br/><a:r><a:t>after a break</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" idx="5"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{1}" type="slidenum"><a:t>2</a:t></a:fld></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:notes>',
  )
  z['ppt/notesSlides/_rels/notesSlide1.xml.rels'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/></Relationships>')
  const rels = strFromU8(z['ppt/slides/_rels/slide2.xml.rels'])
  z['ppt/slides/_rels/slide2.xml.rels'] = strToU8(rels.replace('</Relationships>', '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>'))
  const ct = strFromU8(z['[Content_Types].xml'])
  z['[Content_Types].xml'] = strToU8(ct.replace('</Types>', '<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>'))
  return zipSync(z)
}

describe('P3-013 speaker notes', () => {
  test('a pptx\'s notes are on the slide model and in the text export', async () => {
    const { slide, all } = await open(withNotes())
    expect((await slide(1)).notes).toBe('Speaker notes for slide 2\nSecond line\nafter a break')
    expect((await slide(0)).notes).toBeUndefined()
    const text = slidesToText(await all())
    expect(text).toContain('Notes:\nSpeaker notes for slide 2\nSecond line\nafter a break')
    expect(text.match(/Notes:/g)?.length).toBe(1)
  })

  test('a converted .odp keeps its notes in the model', async () => {
    const { slide } = await open(bytes('sample.odp'), 'sample.odp')
    expect((await slide(0)).notes).toBe('Welcome everyone and introduce the project.')
    expect((await slide(1)).notes).toBe('Mention the lazy ZIP reader first.')
  })

  test('a converted .ppt keeps its notes in the model', async () => {
    const { slide } = await open(bytes('sample.ppt'), 'sample.ppt')
    expect((await slide(1)).notes).toContain('Speaker notes: mention the worker and the lazy ZIP reader.')
  })

  test('a copy of the slide has no notes; the original keeps them; saving keeps the notes part', async () => {
    const { engine, slide } = await open(withNotes())
    await engine.call('pptx.slides', [{ op: 'duplicate', index: 1 }])
    expect((await slide(2)).notes).toBeUndefined()
    expect((await slide(1)).notes).toContain('Speaker notes')
    const out = await saved(engine)
    expect(await (await parts(out)).text('ppt/notesSlides/notesSlide1.xml')).toContain('Speaker notes for slide 2')
  })
})

describe('N2-024 slide list edits undo', () => {
  test('delete: undo brings the slide back (part, relationships, notes, place); redo deletes again', async () => {
    const { engine, slide, info } = await open(withNotes())
    const before = { slides: info.pptx!.slides.map((s) => ({ ...s })) }
    const titles = await Promise.all(before.slides.map((s) => slide(s.index).then(titleOf)))
    const meta = (await engine.call('pptx.slides', [{ op: 'delete', index: 1 }])) as PresentationMeta
    expect(meta.count).toBe(4)
    const u: any = await engine.call('pptx.undo', [])
    expect(u.meta.count).toBe(5)
    expect(u.index).toBe(1)
    expect(u.meta.slides.map((s: any) => s.path)).toEqual(before.slides.map((s) => s.path))
    expect(titleOf(u.model)).toBe(titles[1])
    expect(u.model.notes).toContain('Speaker notes')
    expect(await Promise.all(u.meta.slides.map((s: any) => slide(s.index).then(titleOf)))).toEqual(titles)
    // saved: all five slides and the notes part, the file reopens
    const out = await saved(engine)
    const p = await parts(out)
    expect(await p.text('ppt/notesSlides/notesSlide1.xml')).toContain('Speaker notes')
    expect(await p.text('[Content_Types].xml')).toContain('/ppt/slides/slide2.xml')
    const re = await open(new Uint8Array(await out.arrayBuffer()))
    expect(re.info.pptx!.count).toBe(5)
    expect((await re.slide(1)).notes).toContain('Speaker notes')
    // redo: gone again
    const r: any = await engine.call('pptx.undo', [true])
    expect(r.meta.count).toBe(4)
    expect(r.meta.slides.map((s: any) => s.path)).not.toContain(before.slides[1].path)
    const out2 = await saved(engine)
    const re2 = await open(new Uint8Array(await out2.arrayBuffer()))
    expect(re2.info.pptx!.count).toBe(4)
  })

  test('delete takes the slide\'s comments; undo gives them back', async () => {
    const { engine, info } = await open(bytes('sample.pptx'))
    const slidePath = info.pptx!.slides[2].path
    const c: DocComment = { id: '{99999999-9999-4999-8999-999999999999}', author: { name: 'Ann' }, time: '2026-11-07T09:00:00Z', text: 'Look here', at: { slide: slidePath, x: 20, y: 20 } }
    await engine.call('comments.set', [c.id, c])
    await engine.call('pptx.slides', [{ op: 'delete', index: 2 }])
    expect(((await engine.call('comments.list', [])) as DocComment[]).length).toBe(0)
    await engine.call('pptx.undo', [])
    expect(((await engine.call('comments.list', [])) as DocComment[]).map((x) => x.text)).toEqual(['Look here'])
    const out = await saved(engine)
    const re = new Engine(() => {})
    await re.open(out, { name: 'deck.pptx' })
    const got = (await re.call('comments.list', [])) as DocComment[]
    expect(got.map((x) => x.text)).toEqual(['Look here'])
  })

  test('add, duplicate, move and hide are undoable one by one, in order', async () => {
    const { engine, slide, paths } = await open(bytes('sample.pptx'))
    const list = async () => paths()
    const p0 = await list()
    await engine.call('pptx.slides', [{ op: 'add', after: 0 }])
    const p1 = await list()
    expect(p1.length).toBe(6)
    await engine.call('pptx.slides', [{ op: 'duplicate', index: 3 }])
    const p2 = await list()
    expect(p2.length).toBe(7)
    await engine.call('pptx.slides', [{ op: 'move', from: 0, to: 4 }])
    const p3 = await list()
    expect(p3[4]).toBe(p2[0])
    await engine.call('pptx.slides', [{ op: 'hide', index: 2, hidden: true }])
    expect((await slide(2)).hidden).toBe(true)
    const undo = async () => ((await engine.call('pptx.undo', [])) as any).meta.slides.map((s: any) => s.path)
    await undo()
    expect((await slide(2)).hidden).toBe(false)
    expect(await undo()).toEqual(p2)
    expect(await undo()).toEqual(p1)
    expect(await undo()).toEqual(p0)
    expect(await engine.call('pptx.undo', [])).toBeNull()
    // and forward again
    const redo = async () => ((await engine.call('pptx.undo', [true])) as any).meta.slides.map((s: any) => s.path)
    expect(await redo()).toEqual(p1)
    expect(await redo()).toEqual(p2)
    expect(await redo()).toEqual(p3)
    // the file after undo / redo is the same deck
    const re = await open(new Uint8Array(await (await saved(engine)).arrayBuffer()))
    expect(re.info.pptx!.count).toBe(7)
  })

  test('undo of slide edits and of shape edits share one stack', async () => {
    const { engine, slide } = await open(bytes('sample.pptx'))
    const s2 = await slide(2)
    const top = s2.els.find((e) => e.top !== undefined)!.top!
    await engine.call('pptx.shapes', [2, [{ op: 'delete', top }]])
    await engine.call('pptx.slides', [{ op: 'delete', index: 0 }])
    const a: any = await engine.call('pptx.undo', [])
    expect(a.meta.count).toBe(5)
    const b: any = await engine.call('pptx.undo', [])
    expect(b.meta).toBeUndefined()
    expect(b.model.els.some((e: SlideEl) => e.top === top)).toBe(true)
  })

  test('a slide added, then typed into: undo takes the text back, then the slide; redo brings the slide, then the text', async () => {
    const { engine, slide } = await open(bytes('sample.pptx'))
    await engine.call('pptx.slides', [{ op: 'add', after: 0 }])
    const m = await slide(1)
    const ph = m.els.find((e) => e.t === 'shape' && e.sid !== undefined && e.text) as Extract<SlideEl, { t: 'shape' }>
    await engine.call('pptx.setText', [1, ph.sid, [{ runs: [{ text: 'Typed here' }] }]])
    const t: any = await engine.call('pptx.undo', [])
    expect(t.meta).toBeUndefined()
    expect(JSON.stringify(t.model)).not.toContain('Typed here')
    const u: any = await engine.call('pptx.undo', [])
    expect(u.meta.count).toBe(5)
    const r: any = await engine.call('pptx.undo', [true])
    expect(r.meta.count).toBe(6)
    expect(JSON.stringify(r.model)).not.toContain('Typed here')
    const r2: any = await engine.call('pptx.undo', [true])
    expect(JSON.stringify(r2.model)).toContain('Typed here')
  })

  test('QA 3: text typed in a shape is an undo step; a shape edit before it stays one too', async () => {
    const { engine, slide } = await open(bytes('sample.pptx'))
    const hist = () => engine.call('pptx.history', []) as Promise<{ undo: number; redo: number }>
    const textOf = (m: SlideModel, sid: number) => {
      const e = m.els.find((x) => x.t === 'shape' && x.sid === sid) as Extract<SlideEl, { t: 'shape' }>
      return e.text!.paras.map((p) => p.runs.map((r) => r.text).join('')).join('|')
    }
    const m0 = await slide(0)
    const title = m0.els.find((e) => e.t === 'shape' && e.sid !== undefined && e.text && e.top !== undefined) as Extract<SlideEl, { t: 'shape' }>
    const was = textOf(m0, title.sid!)
    expect(await hist()).toEqual({ undo: 0, redo: 0 })
    // a shape edit, then text typed in two writes of one stay in the shape (the second goes on with the first)
    await engine.call('pptx.shapes', [0, [{ op: 'xfrm', top: title.top!, box: { ...title.box, x: title.box.x + 10 } }]])
    await engine.call('pptx.setText', [0, title.sid, [{ k: 0, runs: [{ text: 'One' }] }], undefined, true])
    await engine.call('pptx.setText', [0, title.sid, [{ k: 0, runs: [{ text: 'One two' }] }]])
    expect(await hist()).toEqual({ undo: 2, redo: 0 })
    // the text entered again: a step of its own
    await engine.call('pptx.setText', [0, title.sid, [{ k: 0, runs: [{ text: 'One two three' }] }], undefined, true])
    expect(await hist()).toEqual({ undo: 3, redo: 0 })
    const u1: any = await engine.call('pptx.undo', [])
    expect(textOf(u1.model, title.sid!)).toBe('One two')
    const u2: any = await engine.call('pptx.undo', [])
    expect(textOf(u2.model, title.sid!)).toBe(was)
    // the shape is still where it was moved to: only the text went back
    const moved = (u2.model as SlideModel).els.find((e) => e.top === title.top)!
    expect(Math.round(moved.box.x)).toBe(Math.round(title.box.x + 10))
    expect(await hist()).toEqual({ undo: 1, redo: 2 })
    // the text as the file had it is edited again from what undo gave: the runs are told from the file's again
    const back = (u2.model as SlideModel).els.find((e) => e.t === 'shape' && e.sid === title.sid) as Extract<SlideEl, { t: 'shape' }>
    const src = back.text!.paras[0].runs[0].src as [number, number]
    const again: any = await engine.call('pptx.setText', [0, title.sid, [{ k: 0, runs: [{ text: 'Kept look', src }] }], undefined, true])
    expect(textOf(again, title.sid!)).toBe('Kept look')
    expect(await hist()).toEqual({ undo: 2, redo: 0 })
    await engine.call('pptx.undo', [])
    const u3: any = await engine.call('pptx.undo', [])
    expect(Math.round((u3.model as SlideModel).els.find((e) => e.top === title.top)!.box.x)).toBe(Math.round(title.box.x))
    expect(await hist()).toEqual({ undo: 0, redo: 2 })
    const r1: any = await engine.call('pptx.undo', [true])
    const r2: any = await engine.call('pptx.undo', [true])
    expect(Math.round((r1.model as SlideModel).els.find((e) => e.top === title.top)!.box.x)).toBe(Math.round(title.box.x + 10))
    expect(textOf(r2.model, title.sid!)).toBe('Kept look')
    // and the file saved holds it
    const re = await open(new Uint8Array(await (await saved(engine)).arrayBuffer()))
    expect(textOf(await re.slide(0), title.sid!)).toBe('Kept look')
  })

  test('a slide list edit after undo drops what was undone', async () => {
    const { engine } = await open(bytes('sample.pptx'))
    await engine.call('pptx.slides', [{ op: 'delete', index: 1 }])
    await engine.call('pptx.undo', [])
    await engine.call('pptx.slides', [{ op: 'delete', index: 3 }])
    expect(await engine.call('pptx.undo', [true])).toBeNull()
  })
})

describe('N2-026 table cell text', () => {
  async function cellOf() {
    const o = await open(bytes('sample.pptx'))
    const m = await o.slide(3)
    const t = tableOf(m)
    return { ...o, t }
  }

  test('the model gives the table an id and every cell a text', async () => {
    const { t } = await cellOf()
    expect(t.sid).toBeGreaterThan(0)
    expect(cellText(t, 1, 0)).toBe('Users')
    expect(t.rows.every((r) => r.cells.every((c) => c.text))).toBe(true)
  })

  test('edit a cell: the table shows it, the saved a:tbl has it, the other cells and the structure are as they were', async () => {
    const { engine, t } = await cellOf()
    const src = t.rows[1].cells[0].text!.paras[0].runs[0].src
    const m: any = await engine.call('pptx.setText', [3, t.sid, [{ k: 0, runs: [{ text: 'Utilisateurs ', src, set: { b: true } }, { text: 'actifs' }] }], [1, 0]])
    const t2 = tableOf(m)
    expect(cellText(t2, 1, 0)).toBe('Utilisateurs actifs')
    expect(cellText(t2, 1, 1)).toBe('380')
    expect(cellText(t2, 0, 0)).toBe('Metric')
    // twice: the second edit still finds its template
    await engine.call('pptx.setText', [3, t.sid, [{ k: 0, runs: [{ text: 'Users 2', src }] }, { k: 0, runs: [{ text: 'second line' }] }], [1, 0]])
    const out = await saved(engine)
    const xml = await (await parts(out)).text('ppt/slides/slide4.xml')
    expect(xml).toContain('Users 2')
    expect(xml).toContain('second line')
    expect(xml).not.toContain('Utilisateurs')
    // schema order inside the cell: txBody, then tcPr; the table keeps its cells
    expect(xml.match(/<a:tc>/g)!.length).toBe(strFromU8(unzipSync(bytes('sample.pptx'))['ppt/slides/slide4.xml']).match(/<a:tc>/g)!.length)
    for (const tc of xml.match(/<a:tc>.*?<\/a:tc>/g)!) expect(tc).toMatch(/^<a:tc><a:txBody>.*<\/a:txBody><a:tcPr\/><\/a:tc>$/)
    const re = await open(new Uint8Array(await out.arrayBuffer()))
    const t3 = tableOf(await re.slide(3))
    expect(cellText(t3, 1, 0)).toBe('Users 2|second line')
    expect(cellText(t3, 3, 3)).toBe(cellText(t, 3, 3))
    expect(cellText(t3, 1, 1)).toBe('380')
  })

  test('clearing a cell leaves an empty cell, not a removed table; an empty cell takes text', async () => {
    const { engine, t } = await cellOf()
    let m: any = await engine.call('pptx.setText', [3, t.sid, [{ k: 0, runs: [] }], [1, 1]])
    expect(cellText(tableOf(m), 1, 1) ?? '').toBe('')
    expect(tableOf(m).rows[1].cells[1].text).toBeTruthy()
    m = await engine.call('pptx.setText', [3, t.sid, [{ runs: [{ text: 'filled' }] }], [1, 1]])
    expect(cellText(tableOf(m), 1, 1)).toBe('filled')
    const xml = await (await parts(await saved(engine))).text('ppt/slides/slide4.xml')
    expect(xml).toContain('filled')
    expect(xml).not.toContain('>380<')
  })

  test('a cell without a text body gets one before its properties; a wrong cell is refused', async () => {
    const z = unzipSync(bytes('sample.pptx'))
    const xml = strFromU8(z['ppt/slides/slide4.xml']).replace('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Metric</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>', '<a:tc><a:tcPr/></a:tc>')
    z['ppt/slides/slide4.xml'] = strToU8(xml)
    const { engine, slide } = await open(zipSync(z))
    const t = tableOf(await slide(3))
    expect(t.rows[0].cells[0].text).toBeTruthy()
    await engine.call('pptx.setText', [3, t.sid, [{ runs: [{ text: 'Head' }] }], [0, 0]])
    const out = await (await parts(await saved(engine))).text('ppt/slides/slide4.xml')
    expect(out).toContain('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>')
    expect(out).toMatch(/<a:t>Head<\/a:t><\/a:r><a:endParaRPr[^>]*\/><\/a:p><\/a:txBody><a:tcPr\/><\/a:tc>/)
    await expect(engine.call('pptx.setText', [3, t.sid, [{ runs: [{ text: 'x' }] }], [9, 9]])).rejects.toThrow()
  })

  test('a table duplicated on its slide is editable too (its template is the copy itself)', async () => {
    const { engine, slide } = await cellOf()
    const m = await slide(3)
    await engine.call('pptx.shapes', [3, [{ op: 'duplicate', top: m.els.find((e) => e.t === 'table')!.top! }]])
    const tables = (await slide(3)).els.filter((e): e is Table => e.t === 'table')
    expect(tables.length).toBe(2)
    expect(tables[1].sid).not.toBe(tables[0].sid)
    const r: any = await engine.call('pptx.setText', [3, tables[1].sid, [{ runs: [{ text: 'copy cell' }] }], [0, 0]])
    const ts = r.els.filter((e: SlideEl) => e.t === 'table') as Table[]
    expect(cellText(ts[1], 0, 0)).toBe('copy cell')
    expect(cellText(ts[0], 0, 0)).toBe('Metric')
  })
})

describe('embedded objects (OLE)', () => {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  /** sample.pptx with an embedded worksheet on slide 1, the way PowerPoint 2007 writes one: its picture is the legacy drawing's */
  function withObject(shapeAttrs: string): Uint8Array {
    const z = unzipSync(bytes('sample.pptx'))
    const frame =
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="900" name="Object 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj spid="_x0000_s1026" name="Worksheet" r:id="rId90" imgW="1828800" imgH="914400" progId="Excel.Sheet.8"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>'
    z['ppt/slides/slide1.xml'] = strToU8(strFromU8(z['ppt/slides/slide1.xml']).replace('</p:spTree>', frame + '</p:spTree>'))
    const rels = 'ppt/slides/_rels/slide1.xml.rels'
    z[rels] = strToU8(
      strFromU8(z[rels]).replace(
        '</Relationships>',
        `<Relationship Id="rId90" Type="${R}/oleObject" Target="../embeddings/oleObject1.bin"/><Relationship Id="rId91" Type="${R}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`,
      ),
    )
    z['ppt/embeddings/oleObject1.bin'] = new Uint8Array([1, 2, 3])
    z['ppt/drawings/vmlDrawing1.vml'] = strToU8(
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shapetype id="_x0000_t75"/><v:shape ${shapeAttrs} type="#_x0000_t75" style='position:absolute'><v:imagedata o:relid="rId1" o:title=""/></v:shape></xml>`,
    )
    z['ppt/drawings/_rels/vmlDrawing1.vml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/image" Target="../media/object1.png"/></Relationships>`)
    z['ppt/media/object1.png'] = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0))
    z['[Content_Types].xml'] = strToU8(strFromU8(z['[Content_Types].xml']).replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>'))
    return zipSync(z)
  }

  test('an object without a picture of its own shows the one of its legacy drawing shape, not a placeholder', async () => {
    // the shape is the one the object names: by its id, or (a named shape) by o:spid
    for (const attrs of ['id="_x0000_s1026"', 'id="Object_x0020_2" o:spid="_x0000_s1026"']) {
      const { slide } = await open(withObject(attrs))
      const els = (await slide(0)).els
      expect(els.filter((e) => e.t === 'placeholder')).toEqual([])
      const pic = els.find((e) => e.t === 'pic' && e.target === 'ppt/media/object1.png')
      expect(pic && Math.round(pic.box.w)).toBe(192)
    }
    // a drawing that does not hold the shape: the placeholder, as before
    const none = await open(withObject('id="_x0000_s2222"'))
    expect((await none.slide(0)).els.filter((e) => e.t === 'placeholder').map((e) => e.t === 'placeholder' && e.label)).toEqual(['Embedded object'])
  })
})
