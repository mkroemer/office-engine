import { afterEach, beforeEach, expect, test } from 'bun:test'
import { RichInput } from '../src/react/edit/input'

// RichInput follows the selection through one animation frame at a time. Detaching cancels a pending frame; if the handle is not
// cleared too, every later selection change is ignored: the editor stops knowing where the caret is, the page holding the caret is
// no longer kept mounted while scrolling, and typing after a scroll lands at the top of the document.

type Fn = () => void
const g = globalThis as Record<string, unknown>
const saved: Record<string, unknown> = {}
let frames: Fn[] = []
let selChange: Fn | null = null

beforeEach(() => {
  for (const k of ['document', 'requestAnimationFrame', 'cancelAnimationFrame']) saved[k] = g[k]
  frames = []
  selChange = null
  g.document = {
    addEventListener: (type: string, fn: Fn) => type === 'selectionchange' && (selChange = fn),
    removeEventListener: (type: string, fn: Fn) => type === 'selectionchange' && selChange === fn && (selChange = null),
  }
  g.requestAnimationFrame = (fn: Fn) => frames.push(fn)
  g.cancelAnimationFrame = (id: number) => void (frames[id - 1] = () => {})
})
afterEach(() => {
  for (const k of Object.keys(saved)) g[k] = saved[k]
})

// a root with no selection in it: reading the selection gives null, which the host is still told about
const root = () => ({ addEventListener() {}, removeEventListener() {}, contains: () => false, getRootNode: () => ({}), ownerDocument: { getSelection: () => null } }) as unknown as HTMLElement

test('a selection change after a re-attach is still followed (a cancelled frame does not block the next one)', () => {
  const seen: unknown[] = []
  const input = new RichInput({ onSelection: (s: unknown) => seen.push(s) } as never)
  input.attach(root())
  selChange!() // a frame is now pending
  expect(frames.length).toBe(1)
  input.attach(root()) // another root: detach (the pending frame is cancelled), then attach
  selChange!()
  expect(frames.length).toBe(2) // the new change asked for its own frame
  frames[1]()
  expect(seen.length).toBe(1) // and the host heard about it
})

test('a selection change whose frame a re-render cancelled is still followed (detach and attach of the same root)', () => {
  const seen: unknown[] = []
  const input = new RichInput({ onSelection: (s: unknown) => seen.push(s) } as never)
  const el = root()
  input.attach(el)
  selChange!() // the selection moved (focus went to the page around the editor): a frame is pending
  input.detach() // a re-render runs the editing effect again before that frame
  input.attach(el)
  for (const f of [...frames]) f()
  expect(seen.length).toBe(1) // the host still heard where the selection is now
})
