import { afterEach, describe, expect, spyOn, test } from 'bun:test'
import type { Block, EditOp, EditResult, Paragraph, Table } from '../src/core/docx/types'
import { DocumentClient } from '../src/core/worker/client'
import { DocxEditor, type DocState, type FormatState } from '../src/react/docx/DocxEditor'
import { runsText, type RRun, type Sel } from '../src/react/edit/rich'

const clients: DocumentClient[] = []
afterEach(() => {
  for (const client of clients.splice(0)) client.destroy()
})

async function editor(src: Blob = Bun.file(`${import.meta.dir}/../playground/public/samples/sample.docx`)) {
  const client = await DocumentClient.open(src, { worker: false })
  clients.push(client)
  const st: DocState = { blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }
  client.on('docx:blocks', (blocks) => st.blocks.push(...blocks))
  await client.call('docx.start')
  const ed = new DocxEditor(client, st, () => {})
  const pa = String(st.blocks.findIndex((b) => b.t === 'p' && b.runs.some((r) => r.t === 'text' && r.text.startsWith('This document'))))
  const sel: Sel = { a: { pa, o: 0 }, f: { pa, o: 4 } }
  ed.input.lastSel = sel
  const states: { state: FormatState | null; busy: boolean }[] = []
  ed.onState = (state) => states.push({ state, busy: ed.busy })
  return { client, ed, sel, st, states }
}

async function settle(ed: DocxEditor) {
  for (let i = 0; ed.busy && i < 1000; i++) await Bun.sleep(1)
  expect(ed.busy).toBe(false)
}

/** Let the engine finish normally, but control when the editor receives each response (paragraph and table writes). */
function holdEdits(client: DocumentClient) {
  const pending: { result: Promise<EditResult>; release: () => void }[] = []
  client.interceptor = (method, args) => {
    if (method !== 'docx.edit' && method !== 'docx.table') return
    const result = (client.direct as (m: string, op: unknown) => Promise<EditResult>)(method, args[0] as EditOp)
    let release!: () => void
    const gate = new Promise<void>((r) => { release = r })
    pending.push({ result, release })
    return result.then(async (res) => {
      await gate
      return res
    })
  }
  return pending
}

describe('DOCX conversion state notifications (#19)', () => {
  for (const kind of ['bullet', 'number'] as const) test(`${kind} removal, undo and redo publish the saved formatting without reselection`, async () => {
    const { client, ed, sel, states } = await editor()
    ed.toggleList(kind)
    await settle(ed)
    const added = states.at(-1)

    ed.toggleList(kind)
    await settle(ed)
    expect(ed.para(sel.a.pa)?.lst).toBeUndefined()
    expect(ed.para(sel.a.pa)?.marker).toBeUndefined()
    const reopened = await editor(await client.call('save'))
    expect(reopened.ed.para(reopened.sel.a.pa)?.lst).toBeUndefined()
    expect(reopened.ed.para(reopened.sel.a.pa)?.marker).toBeUndefined()
    expect(states.at(-1)).toMatchObject({ state: { lst: undefined }, busy: false })
    expect(added).toMatchObject({ state: { lst: kind, canUndo: true, canRedo: false }, busy: false })

    ed.undo()
    await settle(ed)
    expect(states.at(-1)).toMatchObject({ state: { lst: kind, canUndo: true, canRedo: true }, busy: false })
    ed.redo()
    await settle(ed)
    expect(states.at(-1)).toMatchObject({ state: { lst: undefined, canUndo: true, canRedo: false }, busy: false })
    expect(ed.input.lastSel).toEqual(sel)
    expect(ed.pendingSel).toEqual(sel)
  })

  for (const selected of [true, false]) test(`completion uses the ${selected ? 'new selection' : 'cleared selection'}`, async () => {
    const { client, ed, states } = await editor()
    const pending = holdEdits(client)
    ed.toggleList('bullet')
    await pending[0].result
    // The first render consumed pendingSel; then the selection moved or was cleared.
    ed.pendingSel = null
    const now: Sel | null = selected ? { a: { pa: '0', o: 0 }, f: { pa: '0', o: 4 } } : null
    ed.input.lastSel = now
    ed.onSelection(now)
    const expected = now ? ed.state(now) : null
    const count = states.length
    pending[0].release()
    await settle(ed)
    expect(states).toHaveLength(count + 1)
    expect(states.at(-1)).toEqual({ state: expected, busy: false })
    expect(ed.input.lastSel).toEqual(now)
    expect(ed.pendingSel).toBeNull()
  })

  test('an obsolete conversion does not republish; the accepted replacement publishes current formatting', async () => {
    const { client, ed, sel, states } = await editor()
    const pending = holdEdits(client)
    ed.setPara({ align: 'center' })
    const old = await pending[0].result
    // Alignment has no marker side effects: the whole obsolete result will be ignored.
    expect(old.markers).toEqual({})
    ed.setPara({ align: 'right' })
    await pending[1].result
    const current = ed.para(sel.a.pa)
    const count = states.length
    pending[0].release()
    // Drain the ignored response without releasing the still-pending replacement.
    await Bun.sleep(0)
    expect(ed.busy).toBe(true)
    expect(states).toHaveLength(count)
    expect(ed.para(sel.a.pa)).toBe(current)
    pending[1].release()
    await settle(ed)
    expect(states).toHaveLength(count + 1)
    expect(states.at(-1)).toMatchObject({ state: { align: 'right' }, busy: false })
    expect(ed.para(sel.a.pa)?.ps.align).toBe('right')
  })

  test('an edit with no returned conversion or marker changes keeps its synchronous notification; its answer only ends busy (#29)', async () => {
    const { ed, states } = await editor()
    ed.toggle('b')
    const count = states.length
    expect(states.at(-1)).toMatchObject({ state: { b: true }, busy: true })
    await settle(ed)
    // the same formatting, told once more so that the toolbar does not stay busy
    expect(states).toHaveLength(count + 1)
    expect(states.at(-1)).toEqual({ state: states.at(-2)!.state, busy: false })
  })
})

/** a caret (or a range up to `f`) in paragraph `pa` */
const at = (pa: string, o = 0, f = o): Sel => ({ a: { pa, o }, f: { pa, o: f } })

/** address of the top-level paragraph starting with `text` */
const paraAt = (st: DocState, text: string) => String(st.blocks.findIndex((b) => b.t === 'p' && runsText(b.runs as RRun[]).startsWith(text)))

/** delivers held response `i` and lets the editor take it in */
async function deliver(pending: ReturnType<typeof holdEdits>, i: number) {
  await pending[i].result
  pending[i].release()
  await Bun.sleep(1)
}

/** the editor's input on a stand-in for the page, so that composition events can be sent to it */
function attachInput(ed: DocxEditor) {
  const on = new Map<string, (e: unknown) => void>()
  const g = globalThis as Record<string, unknown>
  const doc = g.document
  g.document = { addEventListener() {}, removeEventListener() {} }
  try {
    ed.input.attach({
      addEventListener: (type: string, fn: (e: unknown) => void) => on.set(type, fn),
      removeEventListener() {},
      contains: () => false,
      querySelectorAll: () => [],
      getRootNode: () => ({}),
      ownerDocument: { getSelection: () => null },
    } as unknown as HTMLElement)
  } finally {
    g.document = doc
  }
  return {
    start: () => on.get('compositionstart')!({}),
    end: (data: string) => on.get('compositionend')!({ data }),
  }
}

/** text, list kind and marker of every paragraph (table cells too), in document order */
function lists(blocks: Block[]) {
  const out: string[] = []
  const para = (p: Paragraph) => out.push(`${runsText(p.runs as RRun[]).slice(0, 12)}|${p.lst ?? ''}|${p.marker?.text ?? ''}`)
  for (const b of blocks) {
    if (b.t === 'p') para(b)
    else if (b.t === 'tbl') for (const r of b.rows) for (const c of r.cells) for (const x of c.blocks) if (x.t === 'p') para(x)
  }
  return out
}

describe('DOCX conversion answers that cross other edits (#26, #29)', () => {
  test('a second edit sent before the first conversion returns gets its own conversion (#26)', async () => {
    const { client, ed, sel, states } = await editor()
    const pending = holdEdits(client)
    ed.toggleList('bullet')
    ed.setPara({ align: 'center' })
    // the first answer comes while the second edit is still on its way
    await deliver(pending, 0)
    await deliver(pending, 1)
    await settle(ed)
    const p = ed.para(sel.a.pa)!
    expect([p.ps.align, p.lst, p.marker?.text]).toEqual(['center', 'bullet', '•'])
    expect(states.at(-1)).toMatchObject({ state: { align: 'center', lst: 'bullet' }, busy: false })
    // the file says the same
    const reopened = await editor(await client.call('save'))
    const q = reopened.ed.para(reopened.sel.a.pa)!
    expect([q.ps.align, q.lst, q.marker?.text]).toEqual(['center', 'bullet', '•'])
  })

  test('the second edit is converted when the first answer comes after it too (#26)', async () => {
    const { client, ed, sel, states } = await editor()
    const pending = holdEdits(client)
    ed.toggleList('bullet')
    ed.setPara({ align: 'right' })
    await deliver(pending, 1)
    await deliver(pending, 0)
    await settle(ed)
    const p = ed.para(sel.a.pa)!
    expect([p.ps.align, p.lst, p.marker?.text]).toEqual(['right', 'bullet', '•'])
    // published with the second answer (the first one still on its way); the first one, obsolete and the last in,
    // tells the same state once more, no longer busy (#29)
    expect(states.slice(-2)).toEqual([
      { state: expect.objectContaining({ align: 'right', lst: 'bullet' }), busy: true },
      { state: expect.objectContaining({ align: 'right', lst: 'bullet' }), busy: false },
    ])
  })

  test('an obsolete list answer publishes nothing: a list turned on and undone never shows as on (#29)', async () => {
    const { client, ed, sel, states } = await editor()
    const pending = holdEdits(client)
    ed.toggleList('bullet')
    ed.undo()
    await deliver(pending, 0)
    await deliver(pending, 1)
    await settle(ed)
    // the list asked for (not known yet), the undo, the undo's conversion: never a pressed list button
    expect(states.map((s) => ({ lst: s.state?.lst, busy: s.busy }))).toEqual([
      { lst: undefined, busy: true },
      { lst: undefined, busy: true },
      { lst: undefined, busy: false },
    ])
    expect(ed.para(sel.a.pa)?.lst).toBeUndefined()
    expect(ed.para(sel.a.pa)?.marker).toBeUndefined()
  })

  test('the undo answered before the list it undoes: no list shown, and the last notification is not busy (#29)', async () => {
    const { client, ed, sel, states } = await editor()
    const pending = holdEdits(client)
    ed.toggleList('bullet')
    ed.undo()
    await deliver(pending, 1)
    await deliver(pending, 0)
    await settle(ed)
    // the list asked for, the undo, the undo's conversion (the list's answer still out), then that obsolete answer,
    // the last in: the model as it stands, no longer busy
    expect(states.map((s) => ({ lst: s.state?.lst, busy: s.busy }))).toEqual([
      { lst: undefined, busy: true },
      { lst: undefined, busy: true },
      { lst: undefined, busy: true },
      { lst: undefined, busy: false },
    ])
    expect(states.at(-1)).toEqual({ state: ed.state(sel), busy: false })
    expect(ed.para(sel.a.pa)?.lst).toBeUndefined()
    expect(ed.para(sel.a.pa)?.marker).toBeUndefined()
  })

  test('the numbering of other paragraphs in an obsolete answer is still shown (#29)', async () => {
    const { client, ed, st } = await editor()
    const hire = paraAt(st, 'Hire the platform team')
    const ship = paraAt(st, 'Ship the viewer engine')
    const pending = holdEdits(client)
    ed.input.lastSel = at(hire, 0, 4)
    // the item leaves the list (the next ones count from 1), and is typed in before the answer comes
    ed.toggleList('number')
    ed.input.lastSel = at(hire, 4)
    ed.insertText('!')
    await deliver(pending, 0)
    await deliver(pending, 1)
    await settle(ed)
    expect(ed.para(hire)?.lst).toBeUndefined()
    expect(ed.para(ship)?.marker?.text).toBe('1.')
    expect(lists(st.blocks)).toEqual(lists((await editor(await client.call('save'))).st.blocks))
  })

  test('a throwing state listener does not make a conversion an unhandled rejection (#29)', async () => {
    const { client, ed } = await editor()
    const pending = holdEdits(client)
    ed.toggleList('bullet')
    const unhandled: unknown[] = []
    const onUnhandled = (e: unknown) => unhandled.push(e)
    process.on('unhandledRejection', onUnhandled)
    const logged = spyOn(console, 'error').mockImplementation(() => {})
    try {
      ed.onState = () => {
        throw new Error('listener failed')
      }
      await deliver(pending, 0)
      await settle(ed)
      await Bun.sleep(20)
      expect(unhandled).toEqual([])
      expect(logged).toHaveBeenCalled()
    } finally {
      process.off('unhandledRejection', onUnhandled)
      logged.mockRestore()
    }
  })
})

describe('DOCX conversion answers during an IME composition (#28)', () => {
  test('an answer that comes during an IME composition is applied when it ends (#28)', async () => {
    const { client, ed, st, states } = await editor()
    const ime = attachInput(ed)
    const hire = paraAt(st, 'Hire the platform team')
    const ship = paraAt(st, 'Ship the viewer engine')
    const apac = paraAt(st, 'Expand to APAC')
    const intro = paraAt(st, 'This document')
    expect([ship, apac].map((pa) => ed.para(pa)?.marker?.text)).toEqual(['2.', '3.'])
    const pending = holdEdits(client)
    ed.input.lastSel = at(hire, 0, 4)
    // the item leaves the list: the ones after it count from 1
    ed.toggleList('number')
    // a composition starts in another paragraph before the answer comes
    ed.input.lastSel = at(intro)
    ime.start()
    const count = states.length
    await deliver(pending, 0)
    expect(ed.input.isComposing).toBe(true)
    // nothing rebuilt under the composition, nothing published (not even "no longer busy": the answer is not taken in)
    expect(ed.busy).toBe(false)
    expect(ed.para(hire)?.lst).toBe('number')
    expect(states).toHaveLength(count)
    ime.end('ă')
    expect(runsText(ed.para(intro)!.runs as RRun[])).toStartWith('ăThis')
    expect(ed.para(hire)?.lst).toBeUndefined()
    expect(ed.para(hire)?.marker).toBeUndefined()
    expect([ship, apac].map((pa) => ed.para(pa)?.marker?.text)).toEqual(['1.', '2.'])
    // the composed text, then the answer applied at its end
    expect(states).toHaveLength(count + 2)
    expect(states.at(-1)).toEqual({ state: ed.state(ed.input.sel()), busy: true })
    // the composed text's own answer, the last in, ends busy
    await deliver(pending, 1)
    await settle(ed)
    expect(states).toHaveLength(count + 3)
    expect(states.at(-1)).toEqual({ state: ed.state(ed.input.sel()), busy: false })
    expect(lists(st.blocks)).toEqual(lists((await editor(await client.call('save'))).st.blocks))
  })
})

describe('DOCX lists in tables the editor regenerates (#27)', () => {
  test('a list in a table inserted this session is shown, published, counted and removed (#27)', async () => {
    const { client, ed, sel, st, states } = await editor()
    const saved = async () => (await editor(await client.call('save'))).st.blocks
    ed.input.lastSel = at(sel.a.pa)
    ed.insertTable(2, 1)
    const bi = Number(sel.a.pa) + 1
    expect(st.blocks[bi].t).toBe('tbl')
    const a = `${bi}.0.0.0`
    const b = `${bi}.1.0.0`
    for (const [pa, text] of [[a, 'One'], [b, 'Two']]) {
      ed.input.lastSel = at(pa)
      ed.insertText(text)
    }
    await settle(ed)

    ed.input.lastSel = at(a, 0, 3)
    ed.toggleList('bullet')
    await settle(ed)
    expect(ed.para(a)?.lst).toBe('bullet')
    expect(ed.para(a)?.marker?.text).toBe('•')
    expect(states.at(-1)).toMatchObject({ state: { lst: 'bullet', inTable: true }, busy: false })
    let file = await saved()
    expect((file[bi] as Table).rows[0].cells[0].blocks[0]).toMatchObject({ lst: 'bullet', marker: { text: '•' } })
    expect(lists(st.blocks)).toEqual(lists(file))

    // the second click takes it away again
    ed.toggleList('bullet')
    await settle(ed)
    expect(ed.para(a)?.lst).toBeUndefined()
    expect(ed.para(a)?.marker).toBeUndefined()
    expect(states.at(-1)).toMatchObject({ state: { lst: undefined, inTable: true }, busy: false })
    file = await saved()
    expect((file[bi] as Table).rows[0].cells[0].blocks[0]).not.toHaveProperty('lst')
    expect(lists(st.blocks)).toEqual(lists(file))

    // numbered items in its cells count in document order, with the numbered paragraphs after the table
    for (const pa of [a, b, String(bi + 1)]) {
      ed.input.lastSel = at(pa, 0, 1)
      ed.toggleList('number')
    }
    await settle(ed)
    expect(ed.para(a)?.lst).toBe('number')
    const hire = () => ed.para(paraAt(st, 'Hire the platform team'))?.marker?.text
    // (the sample's numbered list goes on from them: one list for the numbered paragraphs the editor makes)
    expect([a, b, String(bi + 1)].map((pa) => ed.para(pa)?.marker?.text)).toEqual(['1.', '2.', '3.'])
    expect(hire()).toBe('4.')
    file = await saved()
    expect(lists(st.blocks)).toEqual(lists(file))

    // and with the table taken out, what came after it counts without its items
    ed.input.lastSel = at(a)
    ed.tableOp('delTable')
    await settle(ed)
    expect(st.blocks[bi].t).toBe('p')
    expect(ed.para(String(bi))?.marker?.text).toBe('1.')
    expect(hire()).toBe('2.')
    file = await saved()
    expect(lists(st.blocks)).toEqual(lists(file))
  })

  test('a list in a table of the file restructured this session is shown and removed (#27)', async () => {
    const { client, ed, st, states } = await editor()
    const saved = async () => (await editor(await client.call('save'))).st.blocks
    const ti = st.blocks.findIndex((b) => b.t === 'tbl')
    const north = `${ti}.1.0.0`
    expect(runsText(ed.para(north)!.runs as RRun[])).toBe('North')
    // a row added: the table is written whole from now on
    ed.input.lastSel = at(north)
    ed.tableOp('rowBelow')
    await settle(ed)
    expect((st.blocks[ti] as Table).tid).toBeDefined()

    ed.input.lastSel = at(north, 0, 5)
    ed.toggleList('bullet')
    await settle(ed)
    expect(ed.para(north)).toMatchObject({ lst: 'bullet', marker: { text: '•' } })
    expect(states.at(-1)).toMatchObject({ state: { lst: 'bullet', inTable: true }, busy: false })
    expect(lists(st.blocks)).toEqual(lists(await saved()))

    ed.toggleList('bullet')
    await settle(ed)
    expect(ed.para(north)?.lst).toBeUndefined()
    expect(ed.para(north)?.marker).toBeUndefined()
    expect(states.at(-1)).toMatchObject({ state: { lst: undefined, inTable: true }, busy: false })
    expect(lists(st.blocks)).toEqual(lists(await saved()))
  })

  test('a list turned on in a table cell, then typing in another cell before the answer: both kept (#26, #27)', async () => {
    const { client, ed, sel, st } = await editor()
    ed.input.lastSel = at(sel.a.pa)
    ed.insertTable(1, 2)
    const bi = Number(sel.a.pa) + 1
    const a = `${bi}.0.0.0`
    const b = `${bi}.0.1.0`
    ed.input.lastSel = at(a)
    ed.insertText('One')
    await settle(ed)
    const pending = holdEdits(client)
    ed.input.lastSel = at(a, 0, 3)
    ed.toggleList('bullet')
    ed.input.lastSel = at(b)
    ed.insertText('Two')
    await deliver(pending, 0)
    await deliver(pending, 1)
    await settle(ed)
    expect(ed.para(a)?.lst).toBe('bullet')
    expect(ed.para(a)?.marker?.text).toBe('•')
    expect(runsText(ed.para(b)!.runs as RRun[])).toBe('Two')
    expect(lists(st.blocks)).toEqual(lists((await editor(await client.call('save'))).st.blocks))
  })
})
