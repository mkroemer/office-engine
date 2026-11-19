import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('embedding without React', () => {
  test('<office-viewer>: shadow DOM, attributes, load event', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    const wc = v.page.locator('#wc')
    await wc.locator('.de-page-inner').first().waitFor()
    // rendered inside the element's shadow root, styles included there
    expect(await v.page.evaluate(() => !!document.getElementById('wc')!.shadowRoot?.querySelector('style#office-engine-styles'))).toBe(true)
    expect(await v.page.evaluate(() => (window as any).__wcEvents)).toEqual([{ kind: 'docx', name: 'sample.docx' }])
    // save="false" file-menu="false"
    expect(await wc.locator('.de-file').count()).toBe(0)
    await wc.locator('.de-bar button.de-mode', { hasText: 'Edit' }).click()
    expect(await wc.locator('.de-bar button.de-mode', { hasText: /^(Save|Download)$/ }).count()).toBe(0)
    // attributes are live
    await v.page.evaluate(() => document.getElementById('wc')!.setAttribute('toolbar', 'false'))
    await wc.locator('.de-bar').waitFor({ state: 'detached' })
    await v.page.evaluate(() => document.getElementById('wc')!.setAttribute('src', 'samples/sample.pptx'))
    await wc.locator('.de-slide-inner').first().waitFor()
    await v.page.evaluate(() => document.getElementById('wc')!.setAttribute('theme', 'dark'))
    await wc.locator('.de-root.de-dark').waitFor()
    // N3-010: one `load` per document (an attribute change is not a load)
    await v.page.waitForTimeout(300)
    expect(await v.page.evaluate(() => (window as any).__wcEvents)).toEqual([
      { kind: 'docx', name: 'sample.docx' },
      { kind: 'pptx', name: 'sample.pptx' },
    ])
    expect(v.errors).toEqual([])
  })

  test('<office-viewer>: file property', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    await v.page.evaluate(async () => {
      const wc = document.getElementById('wc') as HTMLElement & { file?: Blob }
      wc.file = await (await fetch('samples/sample.xlsx')).blob()
    })
    await v.page.locator('#wc .de-grid').waitFor()
  })

  test('mount(): options, update(), destroy()', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    const host = v.page.locator('#mounted')
    await host.locator('.de-grid').waitFor()
    expect(await host.locator('.de-bar').count()).toBe(0)
    await v.page.evaluate(() => (window as any).__mounted.update({ toolbar: true }))
    await host.locator('.de-bar').waitFor()
    expect(await host.locator('.de-title').innerText()).toBe('sample.xlsx')
    await v.page.evaluate(() => (window as any).__mounted.destroy())
    expect(await host.innerHTML()).toBe('')
  })

  test('locale and strings update live through mount() and the web component', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    const wc = v.page.locator('#wc')
    const host = v.page.locator('#mounted')
    await wc.locator('.de-page-inner').first().waitFor()
    await host.locator('.de-grid').waitFor()
    await v.page.evaluate(() => {
      const el = document.getElementById('wc') as HTMLElement & { strings?: { file?: { menu?: string } } }
      el.setAttribute('file-menu', 'true')
      const de = window.__german
      el.strings = { ...de, file: { ...de.file, menu: 'Ablage' } }
      ;(window as any).__mounted.update({ toolbar: true, strings: { ...de, file: { ...de.file, menu: 'Dateien' } } })
    })
    await wc.getByRole('button', { name: 'Ablage', exact: true }).waitFor()
    await host.getByRole('button', { name: 'Dateien', exact: true }).waitFor()
    await wc.getByRole('button', { name: 'Ablage', exact: true }).click()
    expect(await wc.locator('.de-dd-pop').innerText()).toContain('Öffnen…')
    await v.page.keyboard.press('Escape')
    await v.page.evaluate(() => {
      const el = document.getElementById('wc') as HTMLElement & { strings?: object }
      el.strings = undefined
      el.setAttribute('locale', 'fr-FR')
      ;(window as any).__mounted.update({ locale: 'en', strings: undefined })
    })
    await wc.getByRole('button', { name: 'File', exact: true }).waitFor()
    await host.getByRole('button', { name: 'File', exact: true }).waitFor()
    expect((await v.page.evaluate(() => (window as any).__wcEvents)).length).toBe(1)
    expect(v.errors).toEqual([])
  })
})
