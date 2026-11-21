import { describe, expect, test } from 'bun:test'
import { docxParas, paraText, runsOf, useBrowser, viewer } from './harness'

useBrowser()

const open = async (viewport?: { width: number; height: number }) => {
  const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true, components: true }, { viewport })
  await v.page.locator('.de-ebar').waitFor()
  return v
}

describe('host UI components', () => {
  test('docx: every toolbar button is the host Button, and it keeps the selection for formatting', async () => {
    const v = await open()
    const counts = await v.page.evaluate(() => {
      const bar = document.querySelector('.de-ebar')!
      return {
        host: bar.querySelectorAll('button[data-host-button]').length,
        engine: bar.querySelectorAll('button.de-tb').length,
        seps: bar.querySelectorAll('[data-host-sep]').length,
        engineSeps: bar.querySelectorAll('.de-tb-sep').length,
      }
    })
    expect(counts.host).toBeGreaterThan(10)
    expect(counts.engine).toBe(0)
    expect(counts.seps).toBeGreaterThan(0)
    expect(counts.engineSeps).toBe(0)

    await v.selectText('document')
    await v.btn('Bold').click()
    await v.page.waitForFunction(() => document.querySelector('button[data-host-button][aria-label^="Bold"]')?.hasAttribute('data-active'))
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(runsOf(p).find((r) => r.text.includes('document'))!.s.b).toBe(true)
    expect(v.errors).toEqual([])
  })

  test('docx: a narrow toolbar still moves what does not fit under the host More button', async () => {
    const v = await open({ width: 640, height: 900 })
    const more = v.page.locator('.de-ebar button[data-host-button][aria-label="More tools"]')
    await more.waitFor()
    // Hosts can use a wide text button: reserve its actual width, not the engine's icon width.
    await v.page.waitForFunction(() => {
      const bar = document.querySelector('.de-ebar')!
      const right = bar.getBoundingClientRect().right
      return [...bar.querySelectorAll(':scope > span')].every((el) => el.getBoundingClientRect().right <= right)
    }, null, { timeout: 5000 })
    await more.click()
    expect(await v.page.locator('.de-ebar-more button[data-host-button]').count()).toBeGreaterThan(0)
    expect(v.errors).toEqual([])
  })

  test('docx: the link prompt is the host Dialog and submits', async () => {
    const v = await open()
    await v.selectText('document')
    await v.page.keyboard.press('Control+k')
    const dialog = v.page.locator('[data-host-dialog][role="dialog"]')
    await dialog.waitFor()
    expect(await v.page.locator('.de-dialog').count()).toBe(0)
    await dialog.locator('input').fill('https://example.org/docs')
    await v.page.keyboard.press('Enter')
    await dialog.waitFor({ state: 'detached' })
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(runsOf(p).find((r) => r.text.includes('document'))!.href).toBe('https://example.org/docs')
    expect(v.errors).toEqual([])
  })

  test('docx: the built-in edit-link prompt keeps Remove link on the left and removes the link', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.selectText('document')
    await v.page.keyboard.press('Control+k')
    const dialog = v.page.locator('.de-dialog')
    await dialog.locator('input').fill('https://example.org/docs')
    await v.page.keyboard.press('Enter')
    await dialog.waitFor({ state: 'detached' })
    await v.selectText('document')
    await v.page.keyboard.press('Control+k')
    const extra = dialog.locator('button.de-dialog-extra')
    await extra.waitFor()
    expect(await extra.innerText()).toBe('Remove')
    const bounds = await dialog.locator('.de-dialog-btns').evaluate((row) => {
      const [extra, cancel] = row.querySelectorAll('button')
      return { left: row.getBoundingClientRect().left, extra: extra.getBoundingClientRect().left, gap: cancel.getBoundingClientRect().left - extra.getBoundingClientRect().right }
    })
    expect(bounds.extra).toBe(bounds.left)
    expect(bounds.gap).toBeGreaterThan(20)
    await extra.click()
    await dialog.waitFor({ state: 'detached' })
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    expect(runsOf(p).find((r) => r.text.includes('document'))!.href).toBeUndefined()
    expect(v.errors).toEqual([])
  })
})
