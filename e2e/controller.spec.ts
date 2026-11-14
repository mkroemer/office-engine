import { describe, expect, test } from 'bun:test'
import type { ViewerState } from '../src'
import { useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const state = (v: Viewer) => v.page.evaluate(() => window.__ctl.snapshot()) as Promise<ViewerState>
/** waits until the controller's state passes `check` (run in the page) */
const until = (v: Viewer, check: string) => v.page.waitForFunction(`(() => { const s = window.__ctl.snapshot(); return ${check} })()`, null, { timeout: 15_000 })

describe('controller and headless mode', () => {
  test('docx: zoom, pages and headings through the controller; no chrome of its own', async () => {
    const v = await viewer({ src: 'samples/sample.docx', headless: true })
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

  test('docx: editing and saving through the controller', async () => {
    const v = await viewer({ src: 'samples/sample.docx', headless: true, onSave: true })
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

  test('xlsx: the sheets as the outline, no tabs of its own', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', headless: true })
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

  test('pptx: slides, zoom and the slide in view; no thumbnails, zoom or Present button', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', headless: true })
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
    const v = await viewer({ src: 'samples/sample.docx' })
    const { page } = v
    expect(await page.locator('.de-bar').count()).toBe(1)
    await until(v, "s.kind === 'docx' && s.zoom")
    await page.locator('button[aria-label="Zoom in"]').click()
    await until(v, "typeof s.zoom.value === 'number'")
    expect(v.errors).toEqual([])
  })
})
