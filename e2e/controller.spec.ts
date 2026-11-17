import { describe, expect, test } from 'bun:test'
import type { ViewerState } from '../src'
import { useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const state = (v: Viewer) => v.page.evaluate(() => window.__ctl.snapshot()) as Promise<ViewerState>
/** waits until the controller's state passes `check` (run in the page) */
const until = (v: Viewer, check: string) => v.page.waitForFunction(`(() => { const s = window.__ctl.snapshot(); return ${check} })()`, null, { timeout: 15_000 })

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
