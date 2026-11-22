/** QA round 3, presentations: undo of text typed in a shape, the state of Undo / Redo, typing and the slide list */
import { describe, expect, test } from 'bun:test'
import { openFile, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const open = async () => {
  const v = await viewer({ src: 'samples/sample.pptx', editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  return v
}
const slide1 = (v: Viewer) => v.page.locator('.de-slide-frame').first()
const btn = (v: Viewer, name: RegExp) => v.page.locator('.de-ebar').getByRole('button', { name }).first()
const savedText = async (v: Viewer, i = 0) => JSON.stringify(((await (await openFile((await v.save()).bytes)).engine.call('pptx.slide', [i])) as any).els)

describe('QA 3: undo in a presentation', () => {
  test('text typed in a shape is taken back by Ctrl+Z after leaving it, and brought back by Redo; the buttons say what they can do', async () => {
    const v = await open()
    expect(await btn(v, /^Undo/).isDisabled()).toBe(true)
    expect(await btn(v, /^Redo/).isDisabled()).toBe(true)
    await slide1(v).getByText('office-engine').first().click()
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.keyboard.type('Quarterly deck')
    await v.page.keyboard.press('Escape')
    await v.page.waitForFunction(() => !(document.querySelector('.de-ebar button[title^="Undo"]') as HTMLButtonElement | null)?.disabled)
    expect(await btn(v, /^Redo/).isDisabled()).toBe(true)
    expect(await slide1(v).innerText()).toContain('Quarterly deck')
    await v.page.keyboard.press('Escape')
    await v.page.keyboard.press('ControlOrMeta+z')
    await v.page.waitForFunction(() => !document.querySelector('.de-slide-frame')?.textContent?.includes('Quarterly deck'))
    expect(await slide1(v).innerText()).toContain('office-engine')
    expect(await savedText(v)).not.toContain('Quarterly deck')
    await v.page.waitForFunction(() => !(document.querySelector('.de-ebar button[title^="Redo"]') as HTMLButtonElement | null)?.disabled)
    expect(await btn(v, /^Undo/).isDisabled()).toBe(true)
    await btn(v, /^Redo/).click()
    await v.page.waitForFunction(() => !!document.querySelector('.de-slide-frame')?.textContent?.includes('Quarterly deck'))
    expect(await savedText(v)).toContain('Quarterly deck')
    expect(await btn(v, /^Redo/).isDisabled()).toBe(true)
    expect(v.errors).toEqual([])
  })

  test('text being typed is kept when another slide is picked in the list; the keys after go to the slides, not into the text', async () => {
    const v = await open()
    await slide1(v).getByText('office-engine').first().click()
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.keyboard.type('Kept title')
    await v.page.locator('.de-thumb').nth(3).click()
    await v.page.waitForTimeout(300)
    await v.page.keyboard.type('zzz')
    await v.page.waitForTimeout(600)
    const first = await savedText(v)
    expect(first).toContain('Kept title')
    expect(first).not.toContain('zzz')
    for (let i = 1; i < 5; i++) expect(await savedText(v, i)).not.toContain('zzz')
    expect(v.errors).toEqual([])
  })
})
