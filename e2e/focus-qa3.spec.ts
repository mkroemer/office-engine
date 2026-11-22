/** QA round 3: the keyboard after a dialog, a thumbnail, a sheet tab */
import { describe, expect, test } from 'bun:test'
import { openFile, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const active = (v: Viewer) =>
  v.page.evaluate(() => {
    const a = document.activeElement as HTMLElement | null
    return { tag: a?.tagName ?? '', inDialog: !!a?.closest('.de-dialog'), inViewer: !!a?.closest('.de-root'), cls: a?.className ?? '' }
  })
const saves = async (v: Viewer) => (await v.events('save')).length

describe('QA 3: a dialog keeps the keyboard and gives it back', () => {
  test('Tab stays inside the dialog; Escape closes it; the document has the focus again and Ctrl+S saves', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    await v.selectText('Quarterly Report')
    await v.page.keyboard.press('Control+k')
    await v.page.locator('.de-dialog').waitFor()
    for (let i = 0; i < 12; i++) {
      await v.page.keyboard.press(i % 3 === 2 ? 'Shift+Tab' : 'Tab')
      expect((await active(v)).inDialog).toBe(true)
    }
    await v.page.keyboard.press('Escape')
    await v.page.locator('.de-dialog').waitFor({ state: 'detached' })
    await v.page.waitForTimeout(100)
    const a = await active(v)
    expect(a.inViewer).toBe(true)
    expect(a.tag).not.toBe('BODY')
    const n = await saves(v)
    await v.page.keyboard.press('Control+s')
    await v.waitEvent('save', n + 1)
    // and typing goes on in the document
    await v.selectText('Quarterly Report', 'end')
    await v.page.keyboard.press('Control+k')
    await v.page.locator('.de-dialog').waitFor()
    await v.page.keyboard.press('Escape')
    await v.page.locator('.de-dialog').waitFor({ state: 'detached' })
    await v.page.waitForTimeout(100)
    await v.page.keyboard.type('!')
    await v.page.waitForTimeout(200)
    expect(await v.page.locator('.de-page-inner').first().innerText()).toContain('Quarterly Report!')
    expect(v.errors).toEqual([])
  })
})

describe('QA 3: presentations', () => {
  test('after a click on a thumbnail the slides have the keyboard: Ctrl+S saves', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    await v.page.locator('.de-thumb').nth(2).click()
    await v.page.waitForTimeout(200)
    expect((await active(v)).cls).toContain('de-scroll')
    const n = await saves(v)
    await v.page.keyboard.press('Control+s')
    await v.waitEvent('save', n + 1)
    expect(v.errors).toEqual([])
  })

  test('a new text box takes what is typed next', async () => {
    const v = await viewer({ src: 'samples/sample.pptx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    await v.pick('Insert', 'Text box')
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.waitForTimeout(200)
    await v.page.keyboard.type('Hello box')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(300)
    const { engine } = await openFile((await v.save()).bytes)
    const texts = ((await engine.call('pptx.slide', [0])) as any).els.map((e: any) => (e.text?.paras ?? []).map((p: any) => p.runs.map((r: any) => r.text).join('')).join('|'))
    expect(texts).toContain('Hello box')
    expect(v.errors).toEqual([])
  })
})

describe('QA 3: spreadsheets', () => {
  test('Ctrl+S saves with the focus on a sheet tab', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    await v.typeCell('J10', 'x')
    await v.page.locator('.de-tab[role=tab]').first().focus()
    const n = await saves(v)
    await v.page.keyboard.press('Control+s')
    await v.waitEvent('save', n + 1)
    expect(v.errors).toEqual([])
  })
})
