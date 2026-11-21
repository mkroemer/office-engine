import { describe, expect, test } from 'bun:test'
import { de } from '../src/locales/de'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('locale: DOCX', () => {
  test('status, toolbar and page setup in German', async () => {
    const v = await viewer({ src: 'samples/sample.docx', texts: 'de', editable: true })
    await v.page.locator('.de-ebar').first().waitFor()
    expect(await v.status()).toMatch(/Seite 1 von \d+/)
    expect(await v.btn('Fett').count()).toBe(1)
    expect(await v.btn('Rückgängig').count()).toBe(1)

    await v.btn(de.docx.pageSetup).click()
    const ps = v.page.locator('[role=dialog][aria-label="Seite einrichten"]')
    await ps.waitFor()
    expect(await ps.innerText()).toContain('Querformat')
    await ps.locator('button', { hasText: 'Abbrechen' }).click()
    await ps.waitFor({ state: 'detached' })
    expect(v.errors).toEqual([])
  })
})
