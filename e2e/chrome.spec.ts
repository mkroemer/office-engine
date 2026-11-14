import { describe, expect, test } from 'bun:test'
import { useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const chrome = (v: Viewer, sel: string) => v.page.locator(`#host-chrome ${sel}`)
const until = (v: Viewer, check: string) => v.page.waitForFunction(`(() => { const s = window.__ctl.snapshot(); return ${check} })()`, null, { timeout: 15_000 })

describe('host-placed controls around a headless viewer', () => {
  test('docx: status, zoom and the headings outline drive and follow the view', async () => {
    const v = await viewer({ src: 'samples/sample.docx', headless: true, chrome: true })
    await until(v, 's.outline && s.outline.items.length > 0 && s.page && s.page.count > 0')
    expect(await chrome(v, '.de-ctl-status').innerText()).toMatch(/Page 1 of \d+/)

    const items = chrome(v, 'nav.de-ctl-outline button')
    expect(await items.count()).toBeGreaterThan(3)
    const table = items.filter({ hasText: 'Table' }).first()
    await table.click()
    await v.page.waitForFunction(() => document.querySelector('#host-chrome nav button.on')?.textContent?.includes('Table'))

    expect(await chrome(v, '.de-ctl-zoom').innerText()).toContain('Fit')
    await chrome(v, 'button[aria-label="Zoom in"]').click()
    await until(v, "typeof s.zoom.value === 'number'")
    expect(await chrome(v, '.de-ctl-zoom').innerText()).toMatch(/\d+\s?%/)
    // no formatting bar while viewing
    expect(await chrome(v, '.de-ebar').count()).toBe(0)
    expect(v.errors).toEqual([])
  })

  test('docx: the format bar formats the selection and follows it', async () => {
    const v = await viewer({ src: 'samples/sample.docx', headless: true, chrome: true, editable: true })
    await until(v, 's.editing')
    const bold = chrome(v, '.de-ebar button[aria-label="Bold"]')
    await bold.waitFor()
    await v.selectText('document')
    await until(v, 's.format && !s.format.bold')
    await bold.click()
    await until(v, 's.format.bold')
    expect(await bold.getAttribute('aria-pressed')).toBe('true')
    // lists exist in documents
    expect(await chrome(v, '.de-ebar button[aria-label="Bulleted list"]').count()).toBe(1)

    // the size list opens beyond the host's container, over the viewer, and its last size can be picked
    await chrome(v, '.de-size-input').click()
    const pop = chrome(v, '.de-size-pop')
    await pop.waitFor()
    const visible = await pop.evaluate((el) => {
      const r = el.getBoundingClientRect()
      const host = document.getElementById('host-chrome')!.getBoundingClientRect()
      const top = document.elementFromPoint(r.left + r.width / 2, r.bottom - 4)
      return { beyond: r.bottom > host.bottom, onTop: !!top && el.contains(top) }
    })
    expect(visible).toEqual({ beyond: true, onTop: true })
    expect(v.errors).toEqual([])
  })

  test('xlsx: sheet tabs switch sheets; no zoom, no outline list of headings', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', headless: true, chrome: true })
    const tabs = chrome(v, '.de-tabs [role="tab"]')
    await tabs.first().waitFor()
    expect(await tabs.allInnerTexts()).toEqual(['Sales', 'Notes', 'Data'])
    expect(await chrome(v, '.de-ctl-zoom').count()).toBe(0)
    await tabs.nth(1).click()
    await until(v, 's.outline.items[1].current')
    expect(await tabs.nth(1).getAttribute('aria-selected')).toBe('true')
    expect(v.errors).toEqual([])
  })
})
