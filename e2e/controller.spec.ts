import { describe, expect, test } from 'bun:test'
import type { ViewerController, ViewerState } from '../src'
import { fromB64, useBrowser, viewer, xlsxSheet, type Viewer } from './harness'

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
      // Dirty is published by run after the held response resumes, so the assertion cannot pass before completion.
      await until(v, 's.dirty')
      expect(await state(v)).toMatchObject({ editing: false, canUndo: false, canRedo: false })
      expect(await v.page.evaluate(() => window.__xlsxHistory.publications.some((s) => s.canUndo || s.canRedo))).toBe(false)
      expect((await saveXlsx(v)).cell('J12')).toBe(method === 'xlsx.undo' ? '' : 'pending history')
      await v.page.evaluate(() => window.__ctl.setEditing(true))
      await until(v, method === 'xlsx.undo' ? 's.canRedo && !s.canUndo' : 's.canUndo && !s.canRedo')
      expect(v.errors).toEqual([])
    })
  }

  test('xlsx: an unmounted view cannot publish its delayed history over the current view', async () => {
    const v = await delayedXlsx()
    await holdHistory(v, 'xlsx.set')
    await v.clickCell('J12')
    await v.page.keyboard.type('old workbook')
    await v.page.keyboard.press('Enter')
    await pendingHistory(v)
    await v.page.evaluate(() => window.__setProps({ src: 'samples/sample.docx', onLoad: undefined }))
    await v.page.locator('.de-pages').waitFor()
    await until(v, "s.kind === 'docx' && !s.loading && !s.canUndo && !s.canRedo")
    await v.page.evaluate(async () => {
      window.__ctl.subscribe(() => window.__xlsxHistory.publications.push(window.__ctl.snapshot()))
      window.__xlsxHistory.release!()
      // The response continuation runs in microtasks before the next frame, even though its grid is gone.
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
    })
    expect(await v.page.evaluate(() => window.__xlsxHistory.publications.some((s) => s.canUndo || s.canRedo))).toBe(false)
    expect(await state(v)).toMatchObject({ kind: 'docx', canUndo: false, canRedo: false })
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
    expect((await state(v)).editing).toBe(false)
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
