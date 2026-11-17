import { describe, expect, test } from 'bun:test'
import { createViewerController, internalController, InternalViewerController, type ViewerCommands, type ViewerController } from '../src/react/controller'

const emptyCommands: ViewerCommands = { setEditing() {}, async save() {}, fileAction() {}, toggleComments() {} }
function mount(ctl: InternalViewerController, commands = emptyCommands) {
  const viewer = ctl.createBinding()
  const release = viewer.attachViewer(commands)
  const view = viewer.createView()
  const detach = view.activate()
  return { viewer, view, release, detach }
}

describe('viewer controller', () => {
  test('state changes count once each; equal updates preserve the snapshot', () => {
    const ctl = new InternalViewerController()
    const { viewer, view } = mount(ctl)
    let calls = 0
    ctl.subscribe(() => calls++)
    viewer.update({ kind: 'docx' })
    view.update({ zoom: { value: 'fit', effective: 0.8 } })
    const before = ctl.snapshot()
    viewer.update({ kind: 'docx' })
    view.update({ zoom: { value: 'fit', effective: 0.8 } })
    expect(ctl.snapshot()).toBe(before)
    expect(calls).toBe(2)
    expect(before.rev).toBe(2)
  })

  test('all host commands remain bound when destructured; absent commands are harmless', async () => {
    const ctl = createViewerController()
    const { setEditing, save, undo, redo, setZoom, zoomIn, zoomOut, goto, gotoPage, present, print, fileAction, toggleComments } = ctl
    undo(); redo(); setEditing(true); await save(); setZoom(1); zoomIn(); zoomOut(); goto('1'); gotoPage(1); present(); print(); fileAction('pdf'); toggleComments()
    const log: unknown[] = []
    const { view } = mount(internalController(ctl), {
      setEditing: (v) => log.push(v), save: async () => void log.push('save'), fileAction: (id) => log.push(id), toggleComments: (v) => log.push(v),
    })
    view.attachView({ undo: () => log.push('undo'), redo: () => log.push('redo'), setZoom: (v) => log.push(v), goto: (v) => log.push(v), gotoPage: (v) => log.push(v), present: () => log.push('present') })
    view.update({ zoom: { value: 'fit', effective: 0.8 } })
    undo(); redo(); setEditing(true); await save(); setZoom(1.25); zoomIn(); zoomOut(); goto('sheet'); gotoPage(2); present(); print(); fileAction('pdf'); toggleComments(false)
    expect(log).toEqual(['undo', 'redo', true, 'save', 1.25, 0.9, 0.75, 'sheet', 2, 'present', 'print', 'pdf', false])
  })

  test('zoom accepts fit and finite factors clamped to the built-in range', () => {
    const ctl = new InternalViewerController()
    const { view } = mount(ctl)
    const values: unknown[] = []
    view.attachView({ setZoom: (z) => values.push(z) })
    for (const z of ['fit', 0, 150, NaN, Infinity, -Infinity, -1, 0.8] as const) ctl.setZoom(z)
    expect(values).toEqual(['fit', 0.25, 3, 0.25, 0.8])
    view.update({ zoom: { value: 0.25, effective: 0.25 } }); ctl.zoomOut()
    view.update({ zoom: null }); ctl.zoomIn()
    expect(values).toHaveLength(5)
  })

  for (const kind of ['docx', 'xlsx', 'pptx'] as const) {
    test(`a slow older view cannot steal a newer ${kind} viewer's state or commands`, () => {
      const ctl = new InternalViewerController()
      const old = mount(ctl)
      old.viewer.update({ loading: true })
      old.view.update({ zoom: { value: 1, effective: 1 }, page: { current: 3, count: 7 } })
      const log: string[] = []
      const current = mount(ctl, { ...emptyCommands, setEditing: () => log.push('edit') })
      current.viewer.update({ kind, loading: false })
      current.view.attachView({ goto: () => log.push('goto'), undo: () => log.push('undo'), setZoom: () => log.push('zoom') })
      current.view.update({ outline: { kind: 'sheets', items: [] }, status: 'new' })
      const snapshot = ctl.snapshot()
      const lateView = old.viewer.createView()
      const off = lateView.activate()
      lateView.attachView({ goto: () => log.push('stale') })
      lateView.update({ page: { current: 3, count: 7 }, status: 'late' })
      old.viewer.update({ kind: 'docx', loading: false, dirty: true })
      off(); old.detach(); old.release()
      ctl.goto('1'); ctl.undo(); ctl.setZoom(1); ctl.setEditing(true)
      expect(log).toEqual(['goto', 'undo', 'zoom', 'edit'])
      expect(ctl.snapshot()).toBe(snapshot)
      expect(snapshot).toMatchObject({ kind, page: null, zoom: null, dirty: false, status: 'new' })
    })
  }

  test('document replacement clears view fields and rejects late publications and cleanup', () => {
    const ctl = new InternalViewerController()
    const { viewer, view, detach } = mount(ctl)
    viewer.update({ kind: 'docx' })
    view.update({ zoom: { value: 1, effective: 1 }, page: { current: 3, count: 7 }, canUndo: true })
    const next = viewer.createView()
    next.activate()
    viewer.update({ kind: 'xlsx' })
    next.update({ status: 'sheet' })
    view.update({ canUndo: true, status: 'stale' }); detach()
    expect(ctl.snapshot()).toMatchObject({ kind: 'xlsx', zoom: null, page: null, canUndo: false, status: 'sheet' })
  })

  test('StrictMode setup-cleanup-setup restores the same scope and cleanup clears all state', () => {
    const ctl = new InternalViewerController()
    const viewer = ctl.createBinding(), view = viewer.createView()
    for (let i = 0; i < 2; i++) {
      const release = viewer.attachViewer(emptyCommands), detach = view.activate()
      const off = view.attachView({})
      viewer.update({ kind: 'pptx', editing: true, dirty: true })
      view.update({ canUndo: true, page: { current: 2, count: 9 }, status: 'Slide 2' })
      expect(ctl.snapshot().kind).toBe('pptx')
      release(); detach(); off()
      view.update({ status: 'late' }); viewer.update({ dirty: true })
      expect(ctl.snapshot()).toMatchObject({ kind: null, loading: false, editing: false, dirty: false, saving: false, canUndo: false, canRedo: false, zoom: null, page: null, outline: null, status: '' })
    }
  })

  test('an unmounted older command registration cannot clear newer commands', () => {
    const ctl = new InternalViewerController(), { view } = mount(ctl)
    let calls = 0
    const old = view.attachView({ undo() { throw Error('stale') } })
    view.attachView({ undo() { calls++ } })
    old(); ctl.undo()
    expect(calls).toBe(1)
  })

  test('observation follows first/last subscriptions, duplicate callbacks and resubscription', () => {
    const ctl = new InternalViewerController(), { view } = mount(ctl)
    const observed: boolean[] = []
    const stop = view.watchObserved(() => observed.push(view.observed))
    const fn = () => {}
    const a = ctl.subscribe(fn), b = ctl.subscribe(fn)
    a(); a(); expect(view.observed).toBe(true)
    b(); const c = ctl.subscribe(fn); c(); stop()
    expect(observed).toEqual([true, false, true, false])
  })

  test('format publication and cleanup require the current command owner, including late callbacks', () => {
    const ctl = new InternalViewerController()
    const { view: old } = mount(ctl)
    const { view: next, detach: detachNext } = mount(ctl)
    const oldCommands = {}, nextCommands = {}
    const format = { bold: true, italic: false, underline: false, strike: false, size: 25 }
    const detachOld = old.attachView(oldCommands)
    next.attachView(nextCommands)
    next.publishFormat(nextCommands, format)
    old.publishFormat(oldCommands, null)
    old.publishFormat(oldCommands, { ...format, size: 10 })
    detachOld()
    expect(ctl.snapshot().format).toEqual(format)
    detachNext()
    next.publishFormat(nextCommands, format)
    expect(ctl.snapshot().format).toBeNull()
  })

  test('invalid sizes never reach a view command; positive finite sizes keep their value', () => {
    const ctl = new InternalViewerController()
    const sizes: number[] = []
    mount(ctl).view.attachView({ format: (c) => { if (c.op === 'size') sizes.push(c.pt) } })
    for (const pt of [NaN, Infinity, -Infinity, 0, -1, 0.25, 12.5, 1000]) ctl.format({ op: 'size', pt })
    expect(sizes).toEqual([0.25, 12.5, 1000])
  })

  test('structural mocks and wrappers fail clearly before mounting', () => {
    const real = createViewerController()
    const fake: Pick<ViewerController, Extract<keyof ViewerController, string>> = { ...real }
    // @ts-expect-error host API is intentionally branded
    expect(() => internalController(fake)).toThrow('must be created by createViewerController()')
    expect(internalController(real)).toBeInstanceOf(InternalViewerController)
  })
})
