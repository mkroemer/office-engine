import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('locale', () => {
  test('bar, File menu and details dialog in German', async () => {
    const v = await viewer({ src: 'samples/sample.docx', locale: 'de' })
    expect(await v.page.locator('.de-bar .de-file').innerText()).toContain('Datei')
    expect(await v.barButton('Bearbeiten').count()).toBe(1)

    const labels = await v.menuLabels()
    expect(labels).toContain('Öffnen…')
    expect(labels).toContain('PDF-Dokument')
    expect(labels).toContain('Dokumenteigenschaften')

    await v.menu('Dokumenteigenschaften')
    const dlg = v.page.locator('[role=dialog][aria-label="Dokumenteigenschaften"]')
    await dlg.waitFor()
    expect(await dlg.innerText()).toContain('Dateigrösse')
    await dlg.locator('button', { hasText: 'Schliessen' }).click()
    await dlg.waitFor({ state: 'detached' })

    await v.barButton('Bearbeiten').click()
    expect(await v.barButton('Ansicht').count()).toBe(1)
    expect(v.errors).toEqual([])
  })

  test('PPTX in German: status, zoom, present', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', locale: 'de' })
    expect(await v.status()).toMatch(/Folie 1 von \d+/)
    expect(await v.page.locator('.de-zoom[aria-label="Zoom"] button[aria-label="Vergrössern"]').count()).toBe(1)
    expect(await v.page.locator('.de-present-btn').innerText()).toContain('Bildschirmpräsentation')
    expect(v.errors).toEqual([])
  })

  test('a partial override keeps the other texts', async () => {
    const v = await viewer({ src: 'samples/sample.docx', locale: 'de', fileLabel: 'Ablage' })
    expect(await v.page.locator('.de-bar .de-file').innerText()).toContain('Ablage')
    expect(await v.barButton('Bearbeiten').count()).toBe(1)
  })

  test('English stays the default', async () => {
    const v = await viewer({ src: 'samples/sample.pptx' })
    expect(await v.page.locator('.de-bar .de-file').innerText()).toContain('File')
    expect(await v.status()).toMatch(/Slide 1 of \d+/)
  })
})
