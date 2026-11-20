import { describe, expect, test } from 'bun:test'
import type { ViewerController, ViewerState } from '../src'
import { docxBlocks, docxParas, paraText, runsOf, fromB64, useBrowser, viewer, xlsxSheet, type Viewer } from './harness'

useBrowser()

const state = (v: Viewer) => v.page.evaluate(() => window.__ctl.snapshot()) as Promise<ViewerState>
/** waits until the controller's state passes `check` (run in the page) */
const until = (v: Viewer, check: string) => v.page.waitForFunction(`(() => { const s = window.__ctl.snapshot(); return ${check} })()`, null, { timeout: 15_000 })

type HistoryMethod = 'xlsx.set' | 'xlsx.undo' | 'xlsx.redo'
declare global {
  interface Window {
    __xlsxHistory: { hold: HistoryMethod | null; release: (() => void) | null; publications: ViewerState[] }
    __oldHistoryController: ViewerController
  }
}

/** Hold a real worker response after the mutation, before the grid can apply it. */
async function delayedXlsx() {
  const v = await viewer({ controller: true, editable: true, headless: true, onSave: true }, { wait: false })
  await v.page.evaluate(() => window.__setProps({
    src: 'samples/sample.xlsx',
    onLoad: (client) => {
      window.__xlsxHistory = { hold: null, release: null, publications: [] }
      const call = client.call.bind(client)
      client.call = async (method, ...args) => {
        const result = await call(method, ...args)
        if (method === window.__xlsxHistory.hold) {
          window.__xlsxHistory.hold = null
          await new Promise<void>((resolve) => { window.__xlsxHistory.release = resolve })
        }
        return result
      }
    },
  }))
  await v.page.locator('.de-grid').waitFor()
  await until(v, "s.kind === 'xlsx' && s.editing")
  return v
}

async function holdHistory(v: Viewer, method: HistoryMethod) {
  await v.page.evaluate((method) => {
    window.__xlsxHistory.hold = method
    window.__xlsxHistory.release = null
  }, method)
}

async function pendingHistory(v: Viewer) {
  await v.page.waitForFunction(() => !!window.__xlsxHistory.release)
}

async function saveXlsx(v: Viewer) {
  const count = (await v.events('save')).length
  await v.page.evaluate(() => window.__ctl.save())
  const saved = await v.waitEvent('save', count + 1)
  return xlsxSheet(fromB64(saved.b64 as string))
}

describe('controller and headless mode', () => {
  test('docx: zoom, pages and headings through the controller; no chrome of its own', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.docx', headless: true })
    const { page } = v
    expect(await page.locator('.de-bar').count()).toBe(0)
    expect(await page.locator('.de-zoom').count()).toBe(0)
    expect(await page.locator('.de-outline-toggle').count()).toBe(0)

    await until(v, "s.kind === 'docx' && s.page && s.page.count > 0 && s.outline && s.outline.items.length > 0")
    const s = await state(v)
    expect(s.zoom?.value).toBe('fit')
    expect(s.status).toMatch(/Page 1 of \d+/)
    expect(s.outline?.kind).toBe('headings')
    expect(s.outline?.items.map((i) => i.title)).toContain('1. Highlights')

    // a heading further down: the view scrolls there and marks it current
    const table = s.outline!.items.find((i) => /Table/.test(i.title))!
    await page.evaluate((id) => window.__ctl.goto(id), table.id)
    await until(v, `s.outline.items.find((i) => i.id === ${JSON.stringify(table.id)}).current`)

    // zoom in widens the page
    const w0 = (await page.locator('.de-page').first().boundingBox())!.width
    await page.evaluate(() => window.__ctl.zoomIn())
    await until(v, "typeof s.zoom.value === 'number'")
    await page.waitForTimeout(200)
    expect((await page.locator('.de-page').first().boundingBox())!.width).toBeGreaterThan(w0)
    expect(v.errors).toEqual([])
  })

  test('an overlapping viewer survives the old viewer unmounting', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.docx', headless: true })
    await v.page.evaluate(() => window.__overlap(1))
    await v.waitEvent('load', 2)
    await until(v, "s.kind === 'docx' && s.zoom && s.page && s.page.count > 0")
    await v.page.evaluate(() => window.__overlap(2))
    await v.page.evaluate(() => window.__ctl.setZoom(1.25))
    await until(v, 's.kind === "docx" && s.zoom.value === 1.25')
    await v.page.evaluate(() => window.__ctl.setEditing(true))
    await until(v, 's.editing')
    expect(v.errors).toEqual([])
  })

  for (const kind of ['docx', 'xlsx', 'pptx'] as const) {
    test(`StrictMode loading overlap: late docx cannot steal the newer ${kind} controller`, async () => {
      const v = await viewer({ controller: true, headless: true, editable: true, strict: true, onSave: true }, { wait: false })
      let oldRequest: import('playwright-core').Route | undefined
      await v.page.route('**/slow-old.docx', (route) => { oldRequest = route })
      await v.page.waitForFunction(() => !!window.__setProps)
      await v.page.evaluate(() => window.__setProps({ src: 'samples/slow-old.docx' }))
      await v.page.waitForFunction(() => window.__ctl.snapshot().loading)
      for (let i = 0; !oldRequest && i < 50; i++) await v.page.waitForTimeout(10)
      expect(oldRequest).toBeDefined()
      await v.page.evaluate((kind) => window.__overlap(1, { src: `samples/sample.${kind}` }), kind)
      await v.waitEvent('load')
      await until(v, `s.kind === '${kind}' && !s.loading && s.outline && s.outline.items.length > 0`)
      await oldRequest!.fulfill({ path: process.cwd() + '/playground/public/samples/sample.docx' })
      await v.waitEvent('load', 2)
      await v.page.locator('.de-pages').first().waitFor()
      await v.page.evaluate(() => window.__overlap(2))
      await v.page.waitForTimeout(200)
      expect((await state(v)).kind).toBe(kind)
      if (kind === 'xlsx') {
        expect(await state(v)).toMatchObject({ zoom: null, page: null })
        await v.page.evaluate(() => window.__ctl.goto(window.__ctl.snapshot().outline!.items[1].id))
        await until(v, 's.outline.items[1].current')
        await v.typeCell('J12', 'new viewer')
        await until(v, 's.canUndo')
        await v.page.evaluate(() => window.__ctl.undo())
        await until(v, 's.canRedo')
        expect(await v.shown('J12')).toBe('')
      } else {
        await v.page.evaluate(() => { const { setZoom, gotoPage } = window.__ctl; setZoom(1.25); gotoPage(2) })
        await until(v, 's.zoom.value === 1.25 && s.page.current === 2')
      }
      expect(v.errors).toEqual([])
    })
  }

  test('bound React callbacks and destructured commands mutate and save the document', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.xlsx', headless: true, editable: true, onSave: true, hostControls: true })
    await v.typeCell('J12', 'bound callback')
    await until(v, 's.canUndo')
    await v.page.getByRole('button', { name: 'Host undo', exact: true }).click()
    await until(v, 's.canRedo && !s.canUndo')
    expect((await saveXlsx(v)).cell('J12')).toBe('')
    await v.page.getByRole('button', { name: 'Host redo', exact: true }).click()
    await until(v, 's.canUndo')
    expect((await saveXlsx(v)).cell('J12')).toBe('bound callback')
    expect(v.errors).toEqual([])
  })

  for (const editable of [false, true]) for (const allowEdit of [false, true]) {
    test(`editable=${editable}, allowEdit=${allowEdit}: toggle visibility and actual edits remain independent`, async () => {
      const v = await viewer({ controller: true, src: 'samples/sample.xlsx', editable, allowEdit, onSave: true })
      await until(v, 's.kind === "xlsx" && !s.loading')
      expect((await state(v)).editing).toBe(editable)
      expect((await state(v)).canEdit).toBe(true)
      expect(await v.barButton(editable ? 'View' : 'Edit').count()).toBe(allowEdit ? 1 : 0)
      await v.page.evaluate(() => window.__ctl.setEditing(true))
      await until(v, 's.editing')
      await v.typeCell('J12', 'editable without toggle')
      await until(v, 's.canUndo && s.dirty')
      expect((await saveXlsx(v)).cell('J12')).toBe('editable without toggle')
      await v.page.evaluate(() => window.__ctl.setEditing(false))
      await until(v, '!s.editing && !s.canUndo')
      await v.page.evaluate(() => window.__ctl.undo())
      expect((await saveXlsx(v)).cell('J12')).toBe('editable without toggle')
      expect(v.errors).toEqual([])
    })
  }

  test('DOCX editable with a hidden toggle: edit, clamp zoom and save/reopen', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.docx', editable: true, allowEdit: false, onSave: true })
    await until(v, 's.editing && s.canEdit')
    expect(await v.barButton('View').count()).toBe(0)
    await v.selectText('Quarterly Report', 'end')
    await v.page.keyboard.type(' hidden-toggle-proof')
    await until(v, 's.dirty && s.canUndo')
    await v.page.evaluate(() => window.__ctl.save())
    const saved = await v.waitEvent('save')
    expect(JSON.stringify(await docxBlocks(fromB64(saved.b64 as string)))).toContain('hidden-toggle-proof')
    await v.page.evaluate(() => { const { setZoom } = window.__ctl; setZoom(0) })
    await until(v, 's.zoom.value === 0.25')
    await v.page.evaluate(() => { const { setZoom, zoomIn } = window.__ctl; setZoom(NaN); zoomIn() })
    await until(v, 's.zoom.value > 0.25')
    await v.page.evaluate(() => window.__ctl.setZoom(150))
    await until(v, 's.zoom.value === 3')
    expect(v.errors).toEqual([])
  })

  test('host outline mouse-down preserves an open cross-sheet formula', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.xlsx', headless: true, editable: true, onSave: true, hostControls: true })
    await v.clickCell('J12')
    await v.page.keyboard.type('=SUM(')
    await v.page.getByRole('button', { name: 'Host second sheet', exact: true }).click()
    await until(v, 's.outline.items[1].current')
    await v.clickCell('A1')
    await v.page.keyboard.press('Enter')
    await until(v, 's.canUndo && s.dirty')
    const saved = await saveXlsx(v)
    const detail = await saved.engine.call('xlsx.cell', [0, 11, 9]) as { formula?: string }
    expect(detail.formula).toMatch(/^=SUM\(Notes!A1\)$/)
    expect(v.errors).toEqual([])
  })

  test('CSV import remains available when Open is disabled; no-document dialogs do not latch', async () => {
    const v = await viewer({ controller: true, allowOpen: false, onSave: true }, { wait: false })
    await v.page.waitForFunction(() => !!window.__setProps)
    await v.page.evaluate(() => { window.__ctl.fileAction('rename'); window.__ctl.fileAction('details') })
    await v.page.evaluate(() => window.__setProps({ src: 'samples/sample.xlsx' }))
    await v.ready('xlsx')
    expect(await v.page.getByRole('dialog').count()).toBe(0)
    expect(await v.menuLabels()).toContain('Import CSV into new sheet…')
    const chooser = v.page.waitForEvent('filechooser')
    await v.page.evaluate(() => window.__ctl.fileAction('import'))
    await (await chooser).setFiles({ name: 'imported.csv', mimeType: 'text/csv', buffer: Buffer.from('imported,value\nproof,42') })
    await until(v, 's.outline.items.some(i => i.title === "imported")')
    const saved = await v.events('save')
    await saveXlsx(v)
    const imported = await v.waitEvent('save', saved.length + 1)
    expect((await xlsxSheet(fromB64(imported.b64 as string), 1)).cell('B2')).toBe('42')
    expect(v.errors).toEqual([])
  })

  test('docx: editing and saving through the controller', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.docx', headless: true, onSave: true })
    const { page } = v
    await until(v, 's.canEdit')
    await page.evaluate(() => window.__ctl.setEditing(true))
    await until(v, 's.editing')
    // the formatting toolbar keeps its row in a headless viewer, without the outline button
    await page.locator('.de-ebar').waitFor()
    expect(await page.locator('.de-ebar [title="Document outline"]').count()).toBe(0)
    await page.evaluate(() => window.__ctl.save())
    const saved = await v.waitEvent('save')
    expect(saved.name).toBe('sample.docx')
    await page.evaluate(() => window.__ctl.setEditing(false))
    await until(v, '!s.editing')
    expect(v.errors).toEqual([])
  })

  test('docx: undo is offered only while editing', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.docx', headless: true })
    await until(v, 's.canEdit')
    await v.page.evaluate(() => window.__ctl.setEditing(true))
    await until(v, 's.editing')
    await v.selectText('Quarterly Report', 'end')
    await v.page.keyboard.type('!!')
    await until(v, 's.canUndo')
    await v.page.evaluate(() => window.__ctl.setEditing(false))
    await until(v, '!s.editing && !s.canUndo && !s.canRedo')
    await v.page.evaluate(() => window.__ctl.setEditing(true))
    await until(v, 's.editing && s.canUndo')
    expect(v.errors).toEqual([])
  })

  test('docx: formatting state and commands through the controller', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.docx', editable: true, onSave: true, headless: true })
    const { page } = v
    await until(v, 's.editing')
    await v.selectText('document')
    await until(v, 's.format && !s.format.bold')
    await page.evaluate(() => window.__ctl.format({ op: 'toggle', mark: 'bold' }))
    await until(v, 's.format.bold')
    await page.evaluate(() => window.__ctl.format({ op: 'size', pt: 20 }))
    await until(v, 's.format.size === 20')
    // headless: no Save button, the host saves through the controller
    await page.evaluate(() => window.__ctl.save())
    const saved = await v.waitEvent('save')
    const p = (await docxParas(fromB64(saved.b64 as string))).find((x) => paraText(x).startsWith('This document'))!
    const run = runsOf(p).find((r) => r.text.includes('document'))!
    expect(run.s.b).toBe(true)
    expect(run.s.sz).toBe(20)
    // viewing: no formatting state, and commands do nothing
    await page.evaluate(() => window.__ctl.setEditing(false))
    await until(v, '!s.editing && s.format === null')
    expect(v.errors).toEqual([])
  })

  test('docx: select all, then a host control: formatting reaches the whole document, not only the pages shown', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.docx', headless: true, editable: true, onSave: true, hostControls: true })
    const { page } = v
    await until(v, 's.editing')
    await v.selectText('Quarterly Report', 'end')
    await page.keyboard.press('ControlOrMeta+a')
    await page.waitForTimeout(150)
    // the host's control takes the focus, as a click on it would
    await page.getByRole('button', { name: 'Host undo', exact: true }).focus()
    await page.waitForTimeout(150)
    await page.evaluate(() => window.__ctl.format({ op: 'toggle', mark: 'bold' }))
    await until(v, 's.format.bold')
    await page.evaluate(() => window.__ctl.save())
    const saved = await v.waitEvent('save')
    const paras = (await docxParas(fromB64(saved.b64 as string))).filter((p) => paraText(p).trim())
    // the last paragraph is pages away from the first: it is bold too
    expect(runsOf(paras.at(-1)!).every((r) => r.s.b)).toBe(true)
    expect(v.errors).toEqual([])
  })

  test("xlsx: the active cell's format through the controller", async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.xlsx', editable: true, onSave: true, headless: true })
    const { page } = v
    await page.locator('.de-grid').waitFor()
    await v.typeCell('J10', '12.5')
    await v.clickCell('J10')
    await until(v, 's.format && !s.format.bold && s.format.list === undefined')
    await page.evaluate(() => window.__ctl.format({ op: 'toggle', mark: 'bold' }))
    await page.evaluate(() => window.__ctl.format({ op: 'highlight', color: '#ffff00' }))
    await until(v, "s.format.bold && /255, 255, 0|ffff00/i.test(s.format.highlight ?? '')")
    // lists are not a cell format: nothing happens
    await page.evaluate(() => window.__ctl.format({ op: 'list', list: 'bullet' }))

    // viewing: undo is not offered and does not change the workbook
    await until(v, 's.canUndo')
    await page.evaluate(() => window.__ctl.setEditing(false))
    await until(v, '!s.editing && !s.canUndo && s.format === null')
    await page.evaluate(() => window.__ctl.undo())
    await page.waitForTimeout(300)
    await page.evaluate(() => window.__ctl.setEditing(true))
    await v.clickCell('J10')
    // the fill was the last edit: it is still there
    await until(v, "s.format && /255, 255, 0|ffff00/i.test(s.format.highlight ?? '') && s.canUndo")
    expect(v.errors).toEqual([])
  })

  test('xlsx: two toggles sent before the worker answers cancel out', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.xlsx', editable: true, onSave: true, headless: true })
    const { page } = v
    await page.locator('.de-grid').waitFor()
    await v.typeCell('J10', '12.5')
    await v.clickCell('J10')
    await until(v, 's.format && !s.format.bold')
    await page.evaluate(() => {
      window.__ctl.format({ op: 'toggle', mark: 'bold' })
      window.__ctl.format({ op: 'toggle', mark: 'bold' })
    })
    await page.waitForTimeout(600)
    await until(v, 's.format && !s.format.bold')
    expect(v.errors).toEqual([])
  })

  test('xlsx: a column selected by its header is formatted whole through the controller, as by the toolbar', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.xlsx', editable: true, headless: true })
    const { page } = v
    await page.locator('.de-grid').waitFor()
    await page.locator('.de-head-c[data-col="12"]').click()
    await until(v, 's.format')
    await page.evaluate(() => window.__ctl.format({ op: 'highlight', color: '#ffff00' }))
    await page.waitForTimeout(400)
    // a cell of the column past the cells in use has the column's format; the next column has not
    await v.clickCell('M20')
    await until(v, "s.format && /255, 255, 0|ffff00/i.test(s.format.highlight ?? '')")
    await v.clickCell('L20')
    await until(v, "s.format && !/255, 255, 0|ffff00/i.test(s.format.highlight ?? '')")
    expect(v.errors).toEqual([])
  })

  test('xlsx: the sheets as the outline, no tabs of its own', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.xlsx', headless: true })
    const { page } = v
    expect(await page.locator('.de-tabs').isVisible()).toBe(false)
    await until(v, "s.kind === 'xlsx' && s.outline && s.outline.kind === 'sheets'")
    const s = await state(v)
    expect(s.zoom).toBeNull()
    expect(s.outline!.items.map((i) => i.title)).toEqual(['Sales', 'Notes', 'Data'])
    expect(s.outline!.items[0].current).toBe(true)
    await page.evaluate((id) => window.__ctl.goto(id), s.outline!.items[1].id)
    await until(v, 's.outline.items[1].current')
    expect(v.errors).toEqual([])
  })

  test('xlsx: history cannot mutate a workbook outside editing', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.xlsx', editable: true, headless: true })
    await v.typeCell('J12', 'history')
    await until(v, 's.canUndo')
    await v.page.evaluate(() => window.__ctl.setEditing(false))
    await until(v, '!s.editing && !s.canUndo && !s.canRedo')
    await v.page.evaluate(() => { window.__ctl.undo(); window.__ctl.redo() })
    expect(await v.shown('J12')).toBe('history')
    await v.page.evaluate(() => window.__ctl.setEditing(true))
    await until(v, 's.canUndo')
    await v.page.evaluate(() => window.__ctl.undo())
    await until(v, 's.canRedo')
    expect(await v.shown('J12')).toBe('')
  })

  test('xlsx: a delayed edit publishes history to the replacement controller only', async () => {
    const v = await delayedXlsx()
    await holdHistory(v, 'xlsx.set')
    await v.clickCell('J12')
    await v.page.keyboard.type('pending edit')
    await v.page.keyboard.press('Enter')
    await pendingHistory(v)
    await v.page.evaluate(async (path) => {
      const { createViewerController } = await import(path)
      window.__oldHistoryController = window.__ctl
      window.__ctl = createViewerController()
      window.__ctl.subscribe(() => {})
      window.__setProps({ controller: window.__ctl })
    }, '/@fs/' + process.cwd() + '/src/index.ts')
    await until(v, "s.kind === 'xlsx' && s.editing")
    const detached = await v.page.evaluate(() => window.__oldHistoryController.snapshot())
    expect(detached).toMatchObject({ kind: null, editing: false, canUndo: false, canRedo: false })
    await v.page.evaluate(() => window.__xlsxHistory.release!())
    await until(v, 's.dirty && s.canUndo')
    expect(await v.page.evaluate(() => window.__oldHistoryController.snapshot())).toEqual(detached)
    expect((await saveXlsx(v)).cell('J12')).toBe('pending edit')
    await v.page.evaluate(() => window.__ctl.undo())
    await until(v, 's.canRedo && !s.canUndo')
    expect((await saveXlsx(v)).cell('J12')).toBe('')
    expect(v.errors).toEqual([])
  })

  for (const method of ['xlsx.set', 'xlsx.undo', 'xlsx.redo'] as const) {
    test(`xlsx: delayed ${method} completion respects the current edit gate`, async () => {
      const v = await delayedXlsx()
      if (method !== 'xlsx.set') {
        await v.typeCell('J12', 'pending history')
        await until(v, 's.canUndo')
        if (method === 'xlsx.redo') {
          await v.page.evaluate(() => window.__ctl.undo())
          await until(v, 's.canRedo && !s.canUndo')
        }
        await saveXlsx(v)
        await until(v, '!s.dirty')
      }
      await holdHistory(v, method)
      if (method === 'xlsx.set') {
        await v.clickCell('J12')
        await v.page.keyboard.type('pending history')
        await v.page.keyboard.press('Enter')
      } else await v.page.evaluate((method) => method === 'xlsx.undo' ? window.__ctl.undo() : window.__ctl.redo(), method)
      await pendingHistory(v)
      await v.page.evaluate(() => window.__ctl.setEditing(false))
      await until(v, '!s.editing && !s.canUndo && !s.canRedo')
      await v.page.evaluate(() => {
        window.__ctl.subscribe(() => window.__xlsxHistory.publications.push(window.__ctl.snapshot()))
        window.__xlsxHistory.release!()
      })
      // Dirty is published by the mounted workbook after the held response resumes; this waits for completion.
      await until(v, 's.dirty')
      expect(await state(v)).toMatchObject({ editing: false, canUndo: false, canRedo: false })
      expect(await v.page.evaluate(() => window.__xlsxHistory.publications.some((s) => s.canUndo || s.canRedo))).toBe(false)
      expect((await saveXlsx(v)).cell('J12')).toBe(method === 'xlsx.undo' ? '' : 'pending history')
      await v.page.evaluate(() => window.__ctl.setEditing(true))
      await until(v, method === 'xlsx.undo' ? 's.canRedo && !s.canUndo' : 's.canUndo && !s.canRedo')
      expect(v.errors).toEqual([])
    })
  }

  for (const replacement of ['docx', 'xlsx'] as const) test(`xlsx: an unmounted view cannot publish delayed dirty/history over replacement ${replacement}`, async () => {
    const v = await delayedXlsx()
    await holdHistory(v, 'xlsx.set')
    await v.clickCell('J12')
    await v.page.keyboard.type('old workbook')
    await v.page.keyboard.press('Enter')
    await pendingHistory(v)
    await v.page.evaluate((kind) => window.__setProps({ src: `samples/sample.${kind}?replacement=1`, onLoad: () => window.__events.push({ type: 'replacement-load' }) }), replacement)
    await v.waitEvent('replacement-load')
    await until(v, `s.kind === '${replacement}' && !s.loading && !s.dirty && !s.canUndo && !s.canRedo`)
    await v.page.evaluate(async () => {
      window.__ctl.subscribe(() => window.__xlsxHistory.publications.push(window.__ctl.snapshot()))
      window.__xlsxHistory.release!()
      // The response continuation runs in microtasks before the next frame, even though its grid is gone.
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
    })
    expect(await v.page.evaluate(() => window.__xlsxHistory.publications.some((s) => s.canUndo || s.canRedo))).toBe(false)
    expect(await state(v)).toMatchObject({ kind: replacement, dirty: false, canUndo: false, canRedo: false })
    if (replacement === 'xlsx') expect((await saveXlsx(v)).cell('J12')).toBe('')
    expect(v.errors).toEqual([])
  })

  test('fully unmounted viewer rejects delayed completion and keeps empty defaults', async () => {
    const v = await delayedXlsx()
    await holdHistory(v, 'xlsx.set')
    await v.clickCell('J12')
    await v.page.keyboard.type('unmounted')
    await v.page.keyboard.press('Enter')
    await pendingHistory(v)
    await v.page.evaluate(() => window.__overlap(3))
    await until(v, 's.kind === null && !s.dirty && !s.canUndo')
    const detached = await state(v)
    await v.page.evaluate(async () => {
      window.__xlsxHistory.release!()
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
    })
    expect(await state(v)).toEqual(detached)
    expect(v.errors).toEqual([])
  })

  test('controller file actions enforce permissions even with the menu hidden', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.xlsx', fileMenu: false, allowOpen: false, allowEdit: false, editable: true, allowSave: false, allowDownload: false, onSave: true })
    const pickers: string[] = [], downloads: string[] = []
    v.page.on('filechooser', () => pickers.push('picker'))
    v.page.on('download', () => downloads.push('download'))
    await v.page.evaluate(() => {
      for (const id of ['new:docx', 'new:invalid', 'open', 'import', 'save', 'original', 'pdf', 'print']) window.__ctl.fileAction(id)
      window.__ctl.setEditing(true)
    })
    await v.page.waitForTimeout(300)
    expect(pickers).toEqual([])
    expect(downloads).toEqual([])
    expect(await v.events('save')).toEqual([])
    expect((await state(v)).editing).toBe(true)
    expect((await state(v)).kind).toBe('xlsx')
    await v.page.evaluate(() => window.__setProps({ allowDownload: true }))
    const dl = await v.download(() => v.page.evaluate(() => window.__ctl.fileAction('text:csv')))
    expect(dl.name).toBe('sample.csv')
    expect(v.errors).toEqual([])
  })

  test('pptx: slides, zoom and the slide in view; no thumbnails, zoom or Present button', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.pptx', headless: true })
    const { page } = v
    expect(await page.locator('.de-thumbs').count()).toBe(0)
    expect(await page.locator('.de-zoom').count()).toBe(0)
    expect(await page.locator('.de-present-btn').count()).toBe(0)
    await until(v, "s.kind === 'pptx' && s.page && s.outline && s.outline.kind === 'slides'")
    const s = await state(v)
    expect(s.page!.current).toBe(1)
    expect(s.outline!.items.length).toBe(s.page!.count)
    await page.evaluate(() => window.__ctl.gotoPage(3))
    await until(v, 's.page.current === 3')
    // slide titles arrive from the slides' title placeholders
    await until(v, "s.outline.items.some((i) => !/^Slide \\d+$/.test(i.title))")
    expect(v.errors).toEqual([])
  })

  test('with chrome: the controller still follows the view', async () => {
    const v = await viewer({ controller: true, src: 'samples/sample.docx' })
    const { page } = v
    expect(await page.locator('.de-bar').count()).toBe(1)
    await until(v, "s.kind === 'docx' && s.zoom")
    await page.locator('button[aria-label="Zoom in"]').click()
    await until(v, "typeof s.zoom.value === 'number'")
    expect(v.errors).toEqual([])
  })
})
