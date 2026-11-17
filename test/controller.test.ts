import { describe, expect, test } from 'bun:test'
import { InternalViewerController } from '../src/react/controller'

describe('viewer controller', () => {
  test('observed once a host subscribes; watchers hear the first subscriber only', () => {
    const ctl = new InternalViewerController()
    let heard = 0
    const stop = ctl.watchObserved(() => heard++)
    expect(ctl.observed).toBe(false)
    const off = ctl.subscribe(() => {})
    ctl.subscribe(() => {})
    expect(ctl.observed).toBe(true)
    expect(heard).toBe(1)
    off()
    stop()
  })

  test('state changes count once each; an equal update is no change', () => {
    const ctl = new InternalViewerController()
    let calls = 0
    ctl.subscribe(() => calls++)
    ctl.update({ kind: 'docx', zoom: { value: 'fit', effective: 0.8 } })
    expect(ctl.snapshot()).toMatchObject({ rev: 1, kind: 'docx', zoom: { value: 'fit', effective: 0.8 } })
    ctl.update({ kind: 'docx', zoom: { value: 'fit', effective: 0.8 } })
    ctl.update({ outline: { kind: 'headings', items: [{ id: '3', title: 'Intro', level: 0, current: true }] } })
    ctl.update({ outline: { kind: 'headings', items: [{ id: '3', title: 'Intro', level: 0, current: true }] } })
    expect(calls).toBe(2)
    expect(ctl.snapshot().rev).toBe(2)
  })

  test('commands go to the view and the viewer that registered them; without them they do nothing', async () => {
    const ctl = new InternalViewerController()
    ctl.goto('1')
    ctl.setEditing(true)
    await ctl.save()
    const log: string[] = []
    ctl.attachView({ setZoom: (z) => log.push(`zoom ${z}`), goto: (id) => log.push(`goto ${id}`), gotoPage: (n) => log.push(`page ${n}`), undo: () => log.push('undo') })
    ctl.attachViewer({ setEditing: (on) => log.push(`edit ${on}`), save: async () => void log.push('save'), fileAction: (id) => log.push(`file ${id}`), toggleComments: () => log.push('comments') })
    ctl.goto('7')
    ctl.gotoPage(2)
    ctl.undo()
    ctl.redo()
    ctl.setEditing(true)
    await ctl.save()
    ctl.print()
    expect(log).toEqual(['goto 7', 'page 2', 'undo', 'edit true', 'save', 'file print'])
  })

  test('zoom steps from the effective zoom: in from a fitted 0.8 is 0.9, out from 0.25 is nothing', () => {
    const ctl = new InternalViewerController()
    const set: unknown[] = []
    ctl.attachView({ setZoom: (z) => set.push(z) })
    ctl.update({ zoom: { value: 'fit', effective: 0.8 } })
    ctl.zoomIn()
    ctl.update({ zoom: { value: 0.25, effective: 0.25 } })
    ctl.zoomOut()
    expect(set).toEqual([0.9])
    // a view without zoom (a spreadsheet)
    ctl.update({ zoom: null })
    ctl.zoomIn()
    expect(set).toEqual([0.9])
  })

  test('the view going clears its part of the state', () => {
    const ctl = new InternalViewerController()
    const detach = ctl.attachView({})
    ctl.update({ kind: 'pptx', page: { current: 2, count: 9 }, canUndo: true, status: 'Slide 2 of 9' })
    detach()
    expect(ctl.snapshot()).toMatchObject({ kind: 'pptx', page: null, canUndo: false, status: '', outline: null, zoom: null })
  })

  test('the viewer going clears its part of the state', () => {
    const ctl = new InternalViewerController()
    const detach = ctl.attachViewer({ setEditing: () => {}, save: async () => {}, fileAction: () => {}, toggleComments: () => {} })
    ctl.update({ kind: 'docx', loading: true, editing: true, canEdit: true, dirty: true, saving: true })
    detach()
    expect(ctl.snapshot()).toMatchObject({ kind: null, loading: false, editing: false, canEdit: false, dirty: false, saving: false })
  })
  test('an exiting viewer cannot detach the newer viewer or view', () => {
    const ctl = new InternalViewerController()
    const log: string[] = []
    const commands = (id: string) => ({ setEditing: () => log.push(id), save: async () => {}, fileAction: () => {}, toggleComments: () => {} })
    const oldViewer = ctl.attachViewer(commands('old'))
    const oldView = ctl.attachView({ undo: () => log.push('old undo') })
    const newViewer = ctl.attachViewer(commands('new'))
    const newView = ctl.attachView({ undo: () => log.push('new undo') })
    ctl.update({ kind: 'xlsx', canUndo: true, status: 'new' })
    oldView(); oldViewer()
    ctl.setEditing(true); ctl.undo()
    expect(log).toEqual(['new', 'new undo'])
    expect(ctl.snapshot()).toMatchObject({ kind: 'xlsx', canUndo: true, status: 'new' })
    newView(); newViewer()
    expect(ctl.snapshot()).toMatchObject({ kind: null, canUndo: false, status: '' })
  })

  test('observation follows the last unsubscribe, resubscribe and duplicate callbacks', () => {
    const ctl = new InternalViewerController()
    const observed: boolean[] = []
    const stop = ctl.watchObserved(() => observed.push(ctl.observed))
    const fn = () => {}
    const a = ctl.subscribe(fn), b = ctl.subscribe(fn)
    a(); a()
    expect(ctl.observed).toBe(true)
    b()
    const c = ctl.subscribe(fn)
    c(); stop()
    expect(observed).toEqual([true, false, true, false])
  })

})
