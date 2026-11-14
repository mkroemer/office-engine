import { describe, expect, test } from 'bun:test'
import { createViewerController } from '../src/react/controller'

describe('viewer controller', () => {
  test('state changes count once each; an equal update is no change', () => {
    const ctl = createViewerController()
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
    const ctl = createViewerController()
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
    const ctl = createViewerController()
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
    const ctl = createViewerController()
    ctl.attachView({})
    ctl.update({ kind: 'pptx', page: { current: 2, count: 9 }, canUndo: true, status: 'Slide 2 of 9' })
    ctl.attachView(null)
    expect(ctl.snapshot()).toMatchObject({ kind: 'pptx', page: null, canUndo: false, status: '', outline: null, zoom: null })
  })
})
