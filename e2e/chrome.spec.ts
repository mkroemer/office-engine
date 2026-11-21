import { describe, expect, test } from 'bun:test'
import { docxParas, fromB64, runsOf, useBrowser, viewer, xlsxSheet, type Viewer } from './harness'

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

  test('docx: a size typed into the format bar and Enter sets it on the selection', async () => {
    const v = await viewer({ src: 'samples/sample.docx', headless: true, chrome: true, editable: true })
    await until(v, 's.editing')
    await chrome(v, '.de-ebar button[aria-label="Bold"]').waitFor()
    await v.selectText('document')
    await until(v, 's.format && s.format.size !== 20')
    const size = chrome(v, '.de-size-input')
    await size.fill('20')
    await size.press('Enter')
    await until(v, 's.format && s.format.size === 20')
    expect(v.errors).toEqual([])
  })

  test('docx: the host format bar uses the shared German locale for text and typed sizes', async () => {
    const v = await viewer({ src: 'samples/sample.docx', headless: true, chrome: true, editable: true, onSave: true, texts: 'de' })
    await until(v, 's.editing')
    await v.selectText('document')
    await until(v, 's.format && s.format.size !== 12.5')
    const size = v.page.locator('#host-chrome').getByLabel('Schriftgrad', { exact: true })
    await size.fill('12,5')
    await size.press('Enter')
    await until(v, 's.format?.size === 12.5')
    await size.fill('17pt')
    await size.press('Enter')
    expect(await size.inputValue()).toBe('12,5')
    await v.page.evaluate(() => window.__ctl.save())
    const saved = await v.waitEvent('save')
    const runs = (await docxParas(fromB64(saved.b64 as string))).flatMap(runsOf)
    expect(runs.find((r) => r.text === 'document')?.s.sz).toBe(12.5)
    expect(v.errors).toEqual([])
  })

  test('docx: a list button that shows as pressed removes the list', async () => {
    const v = await viewer({ src: 'samples/sample.docx', headless: true, chrome: true, editable: true })
    await until(v, 's.editing')
    const bullets = chrome(v, '.de-ebar button[aria-label="Bulleted list"]')
    await bullets.waitFor()
    await v.selectText('document')
    await until(v, 's.format && s.format.list === null')
    await bullets.click()
    await until(v, "s.format.list === 'bullet'")
    expect(await bullets.getAttribute('aria-pressed')).toBe('true')
    await bullets.click()
    await until(v, 's.format.list === null')
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

  for (const [what, sel] of [['sheet tab', '.de-tabs [role="tab"]'], ['sheet in the outline list', 'nav.de-ctl-outline button']]) test(`xlsx: a formula being typed stays open across a ${what}, for a cross-sheet reference`, async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', headless: true, chrome: true, editable: true })
    const tabs = chrome(v, sel)
    await tabs.first().waitFor()
    await v.clickCell('J12')
    await v.page.keyboard.type('=SUM(')
    await until(v, 's.pointing')
    await tabs.nth(1).click()
    await until(v, 's.outline.items[1].current')
    await v.clickCell('A1')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(150)
    await tabs.first().click()
    await until(v, 's.outline.items[0].current')
    expect(await v.formula('J12')).toContain('Notes!A1')
    expect(v.errors).toEqual([])
  })

  test('xlsx: text being typed is committed on its own sheet when a sheet tab is clicked', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', headless: true, chrome: true, editable: true, onSave: true })
    const tabs = chrome(v, '.de-tabs [role="tab"]')
    await tabs.first().waitFor()
    await v.clickCell('J12')
    await v.page.keyboard.type('hello')
    // text, not a formula: nothing to point at, so the tab may take the focus
    expect(await v.page.evaluate(() => window.__ctl.snapshot().pointing)).toBe(false)
    await tabs.nth(1).click()
    await until(v, 's.outline.items[1].current')
    // a click on the other sheet does not carry the text over to it
    await v.clickCell('A1')
    expect(await tabs.nth(1).getAttribute('aria-selected')).toBe('true')
    await tabs.first().click()
    await until(v, 's.outline.items[0].current')
    await v.page.locator('.de-cell').getByText('hello', { exact: true }).waitFor()
    expect(await v.shown('J12')).toBe('hello')
    await v.page.evaluate(() => window.__ctl.save())
    const bytes = fromB64((await v.waitEvent('save')).b64 as string)
    expect((await xlsxSheet(bytes, 0)).cell('J12')).toBe('hello')
    expect((await xlsxSheet(bytes, 1)).cell('J12')).toBe('')
    expect(v.errors).toEqual([])
  })
})
