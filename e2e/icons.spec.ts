import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer, type Viewer } from './harness'

useBrowser()

/** host icons (`.de-icon`) and built-in icon svgs left in the chrome matching `sel` */
async function icons(v: Viewer, sel: string) {
  return v.page.evaluate((s) => {
    const roots = [...document.querySelectorAll(s)]
    const svgs = roots.flatMap((r) => [...r.querySelectorAll('svg')])
    return {
      host: roots.reduce((n, r) => n + r.querySelectorAll('.de-icon > svg[data-host-icon]').length, 0),
      builtIn: svgs.filter((x) => !x.closest('.de-icon')).length,
    }
  }, sel)
}

const hostIcon = (v: Viewer, button: string, name: string) => v.page.locator(`${button} .de-icon > svg[data-host-icon="${name}"]`)

describe('host icons', () => {
  test('docx: zoom, outline toggle and close, toolbar, find & replace', async () => {
    const v = await viewer({ src: 'samples/sample.docx', icons: true })
    expect(await hostIcon(v, 'button[aria-label="Zoom in"]', 'zoomIn').count()).toBe(1)
    expect(await hostIcon(v, 'button[aria-label="Zoom out"]', 'zoomOut').count()).toBe(1)
    expect(await hostIcon(v, '.de-outline-toggle', 'outline').count()).toBe(1)
    await v.page.locator('.de-outline-toggle').click()
    expect(await hostIcon(v, 'button[aria-label="Close outline"]', 'close').count()).toBe(1)
    await v.page.locator('button[aria-label="Close outline"]').click()

    await v.edit()
    const bar = await icons(v, '.de-ebar')
    expect(bar.host).toBeGreaterThan(20)
    expect(bar.builtIn).toBe(0)
    await v.btn('Find & replace').click()
    expect(await hostIcon(v, '.de-find button[aria-label="Replace"]', 'replace').count()).toBe(1)
    expect((await icons(v, '.de-find')).builtIn).toBe(0)
    expect(v.errors).toEqual([])
  })

  test('xlsx: toolbar and the add-sheet tab', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', editable: true, icons: true })
    await v.page.locator('.de-ebar').first().waitFor()
    const bar = await icons(v, '.de-ebar')
    expect(bar.host).toBeGreaterThan(20)
    expect(bar.builtIn).toBe(0)
    expect(await hostIcon(v, '.de-tab-add', 'sheetAdd').count()).toBe(1)
    expect(await v.page.locator('.de-tab-add').getAttribute('title')).toBe('Add sheet')
    expect(v.errors).toEqual([])
  })

  test('pptx: Present, zoom, toolbar', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', icons: true })
    expect(await hostIcon(v, '.de-present-btn', 'present').count()).toBe(1)
    expect(await v.page.locator('.de-present-btn').innerText()).toContain('Present')
    expect(await hostIcon(v, 'button[aria-label="Zoom in"]', 'zoomIn').count()).toBe(1)

    // with nothing selected the PPTX toolbar shows only its slide and insert controls
    await v.edit()
    const bar = await icons(v, '.de-ebar')
    expect(bar.host).toBeGreaterThan(5)
    expect(bar.builtIn).toBe(0)
    expect(v.errors).toEqual([])
  })

  test('without icons the built-in svgs stay', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true })
    await v.page.locator('.de-ebar').first().waitFor()
    const bar = await icons(v, '.de-ebar')
    expect(bar.host).toBe(0)
    expect(bar.builtIn).toBeGreaterThan(20)
    expect(await v.page.locator('button[aria-label="Zoom in"] > svg').count()).toBe(1)
  })
})
