import { afterEach, describe, expect, spyOn, test } from 'bun:test'
import type { Block, EditOp, EditResult, Paragraph } from '../src/core/docx/types'
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

  test('an edit with no returned conversion or marker changes keeps its synchronous notification', async () => {
    const { ed, states } = await editor()
    ed.toggle('b')
    const count = states.length
    await settle(ed)
    expect(states).toHaveLength(count)
    expect(states.at(-1)?.state?.b).toBe(true)
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
    // published with the second answer (the first one still on its way); the first one, obsolete, publishes nothing
    expect(states.at(-1)).toMatchObject({ state: { align: 'right', lst: 'bullet' }, busy: true })
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
