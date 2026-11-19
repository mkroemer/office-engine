import { afterEach, describe, expect, test } from 'bun:test'
import type { EditOp, EditResult } from '../src/core/docx/types'
import { DocumentClient } from '../src/core/worker/client'
import { DocxEditor, type DocState, type FormatState } from '../src/react/docx/DocxEditor'
import type { Sel } from '../src/react/edit/rich'

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
  return { client, ed, sel, states }
}

async function settle(ed: DocxEditor) {
  for (let i = 0; ed.busy && i < 1000; i++) await Bun.sleep(1)
  expect(ed.busy).toBe(false)
}

/** Let the engine finish normally, but control when the editor receives each response. */
function holdEdits(client: DocumentClient) {
  const pending: { result: Promise<EditResult>; release: () => void }[] = []
  client.interceptor = (method, args) => {
    if (method !== 'docx.edit') return
    const result = client.direct('docx.edit', args[0] as EditOp)
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
