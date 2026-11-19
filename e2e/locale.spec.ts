import { describe, expect, test } from 'bun:test'
import { en } from '../src/react/strings'
import { de } from '../src/locales/de'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('locale', () => {
  test('bar, File menu and details dialog in German', async () => {
    const v = await viewer({ src: 'samples/sample.docx', locale: 'de-DE' })
    expect(await v.page.locator('.de-bar .de-file').innerText()).toContain('Datei')
    expect(await v.barButton('Bearbeiten').count()).toBe(1)

    const labels = await v.menuLabels()
    expect(labels).toContain('Öffnen…')
    expect(labels).toContain('PDF-Dokument')
    expect(labels).toContain('Dokumenteigenschaften')

    await v.menu('Dokumenteigenschaften')
    const dlg = v.page.locator('[role=dialog][aria-label="Dokumenteigenschaften"]')
    await dlg.waitFor()
    expect(await dlg.innerText()).toContain('Dateigröße')
    await dlg.locator('button', { hasText: 'Schließen' }).click()
    await dlg.waitFor({ state: 'detached' })

    await v.barButton('Bearbeiten').click()
    expect(await v.barButton('Ansicht').count()).toBe(1)
    expect(v.errors).toEqual([])
  })

  test('PPTX in German: status, zoom, present', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', locale: 'de-DE' })
    expect(await v.status()).toMatch(/Folie 1 von \d+/)
    expect(await v.page.locator('.de-zoom[aria-label="Zoom"] button[aria-label="Vergrößern"]').count()).toBe(1)
    expect(await v.page.locator('.de-present-btn').innerText()).toContain('Bildschirmpräsentation')
    expect(v.errors).toEqual([])
  })

  test('a partial override keeps the other texts', async () => {
    const v = await viewer({ src: 'samples/sample.docx', locale: 'de-DE', fileLabel: 'Ablage' })
    expect(await v.page.locator('.de-bar .de-file').innerText()).toContain('Ablage')
    expect(await v.barButton('Bearbeiten').count()).toBe(1)
  })

  test('English stays the default', async () => {
    const v = await viewer({ src: 'samples/sample.pptx' })
    expect(await v.page.locator('.de-bar .de-file').innerText()).toContain('File')
    expect(await v.status()).toMatch(/Slide 1 of \d+/)
  })

  for (const kind of ['docx', 'xlsx', 'pptx'] as const) {
    test(`${kind}: switching locale updates comment hints without reopening the document`, async () => {
      const v = await viewer({ src: `samples/sample.${kind}`, editable: true, cmOpen: true })
      const empty = v.page.locator('.de-cm-empty')
      await empty.filter({ hasText: en[kind].commentHint }).waitFor()
      await v.page.evaluate(() => window.__setLanguage('de-DE'))
      await empty.filter({ hasText: de[kind].commentHint }).waitFor()
      expect(await v.page.locator('.de-cm-panel').getAttribute('aria-label')).toBe('Kommentare')
      expect((await v.events('load')).length).toBe(1)
      expect(v.errors).toEqual([])
    })
  }

  test('PPTX comment locations follow the UI language', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', locale: 'de-DE', editable: true, cmOpen: true })
    await v.page.locator('.de-cm-new').click()
    const draft = v.page.locator('.de-cm-draft')
    await draft.waitFor()
    expect(await draft.locator('.de-cm-quote').innerText()).toBe('Folie 1')
    await draft.locator('textarea').fill('Kommentar')
    await draft.locator('button', { hasText: 'Kommentieren' }).click()
    const quote = v.page.locator('.de-cm-card .de-cm-quote')
    await quote.filter({ hasText: 'Folie 1' }).waitFor()
    await v.page.evaluate(() => window.__setLanguage('en'))
    await quote.filter({ hasText: 'Slide 1' }).waitFor()
    expect(v.errors).toEqual([])
  })

  test('unknown locales fall back to English, including prototype property names', async () => {
    const v = await viewer({ src: 'samples/sample.docx', locale: 'constructor' })
    expect(await v.page.locator('.de-file').innerText()).toContain('File')
    expect(await v.status()).toMatch(/Page 1 of /)
    expect(v.errors).toEqual([])
  })
})
