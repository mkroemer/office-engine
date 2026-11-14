import { describe, expect, test } from 'bun:test'
import { de } from '../src/react/strings'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('locale', () => {
  test('DOCX in German: bar, toolbar, File menu, dialog', async () => {
    const v = await viewer({ src: 'samples/sample.docx', locale: 'de' })
    expect(await v.page.locator('.de-bar .de-file').innerText()).toContain('Datei')
    expect(await v.barButton('Bearbeiten').count()).toBe(1)
    expect(await v.status()).toMatch(/Seite 1 von \d+/)

    await v.barButton('Bearbeiten').click()
    await v.page.locator('.de-ebar').first().waitFor()
    expect(await v.barButton('Ansicht').count()).toBe(1)
    await expect(v.btn('Fett').count()).resolves.toBe(1)
    await expect(v.btn('Rückgängig').count()).resolves.toBe(1)

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

    await v.btn(de.docx.pageSetup).click()
    const ps = v.page.locator('[role=dialog][aria-label="Seite einrichten"]')
    await ps.waitFor()
    expect(await ps.innerText()).toContain('Querformat')
    await ps.locator('button', { hasText: 'Abbrechen' }).click()
    expect(v.errors).toEqual([])
  })

  test('XLSX in German: status, toolbar, sheet menu', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', locale: 'de', editable: true })
    expect(await v.status()).toMatch(/Zeilen × \d+ Spalten/)
    await expect(v.btn('Fett').count()).resolves.toBe(1)
    await expect(v.btn('Rahmenlinien').count()).resolves.toBe(1)
    await v.page.locator('.de-tab[role=tab]').first().click({ button: 'right' })
    const menu = v.page.locator('.de-ctx')
    await menu.waitFor()
    expect(await menu.innerText()).toContain('Blatt löschen')
    await v.page.keyboard.press('Escape')
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
    const v = await viewer({ src: 'samples/sample.docx' })
    expect(await v.page.locator('.de-bar .de-file').innerText()).toContain('File')
    expect(await v.status()).toMatch(/Page 1 of \d+/)
  })
})
