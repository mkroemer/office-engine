import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('locale: XLSX', () => {
  test('status, toolbar and sheet tab menu in German', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', locale: 'de-DE', editable: true })
    await v.page.locator('.de-ebar').first().waitFor()
    expect(await v.status()).toMatch(/Zeilen × \d+ Spalten/)
    expect(await v.btn('Fett').count()).toBe(1)
    expect(await v.btn('Rahmenlinien').count()).toBe(1)

    await v.page.locator('.de-tab[role=tab]').first().click({ button: 'right' })
    const menu = v.page.locator('.de-ctx')
    await menu.waitFor()
    expect(await menu.innerText()).toContain('Blatt löschen')
    await v.page.keyboard.press('Escape')
    expect(v.errors).toEqual([])
  })
})
