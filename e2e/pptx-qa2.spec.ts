/** QA round 2, presentations: Tab levels, undo of slide list edits, table cell text, speaker notes export. */
import { describe, expect, test } from 'bun:test'
import { openFile, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const open = async (src = 'samples/sample.pptx') => {
  const v = await viewer({ src, editable: true, onSave: true })
  await v.page.locator('.de-ebar').waitFor()
  return v
}

async function saved(v: Viewer) {
  const { engine, info } = await openFile((await v.save()).bytes)
  return { meta: info.pptx!, slide: (i: number): Promise<any> => engine.call('pptx.slide', [i]) }
}

const frame = (v: Viewer, n: number) => v.page.locator(`.de-slide-frame[data-slide="${n}"]`)
const thumbs = (v: Viewer) => v.page.locator('.de-thumb')

/** shows slide `n` (1-based) the way a user does: its thumbnail */
async function goTo(v: Viewer, n: number) {
  await thumbs(v).nth(n - 1).click()
  await frame(v, n).waitFor()
}

const paraOf = (model: any, text: string) =>
  model.els
    .flatMap((e: any) => e.text?.paras ?? [])
    .find((p: any) => p.runs.map((r: any) => r.text).join('').includes(text))

describe('N2-022 / N2-023 Tab and Shift+Tab in a bullet', () => {
  test('Tab raises the level of the bullet (no tab character), Shift+Tab lowers it', async () => {
    const v = await open()
    await goTo(v, 2)
    await frame(v, 2).getByText('Streaming XML').first().click()
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.keyboard.press('Tab')
    await v.page.waitForFunction(() => document.querySelector('.de-txt-edit')?.textContent?.includes('Streaming XML'))
    expect(await v.page.locator('.de-txt-edit').evaluate((e) => e.textContent)).not.toContain('\t')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(300)
    let p = paraOf(await (await saved(v)).slide(1), 'Streaming XML')
    expect(p.lvl).toBe(1)
    expect(p.runs.map((r: any) => r.text).join('')).not.toContain('\t')

    await frame(v, 2).getByText('Streaming XML').first().click()
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.keyboard.press('Shift+Tab')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(300)
    p = paraOf(await (await saved(v)).slide(1), 'Streaming XML')
    expect(p.lvl ?? 0).toBe(0)
    expect(v.errors).toEqual([])
  })

  test('Shift+Tab at the first level and Tab / Shift+Tab undo as one step each', async () => {
    const v = await open()
    await goTo(v, 2)
    await frame(v, 2).getByText('Streaming XML').first().click()
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.keyboard.press('Shift+Tab')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.press('ControlOrMeta+z')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(300)
    expect(paraOf(await (await saved(v)).slide(1), 'Streaming XML').lvl).toBe(1)
  })

  test('in a paragraph that is not a list item Tab still types a tab', async () => {
    const v = await open()
    await frame(v, 1).getByText('office-engine').first().click()
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.keyboard.press('ControlOrMeta+Home')
    await v.page.keyboard.press('Tab')
    await v.page.keyboard.press('Shift+Tab')
    expect(await v.page.locator('.de-txt-edit').evaluate((e) => e.textContent)).toContain('\toffice-engine')
  })
})

describe('N2-024 slide list undo', () => {
  test('delete a slide, Undo (button and keyboard) brings it back, Redo removes it', async () => {
    const v = await open()
    expect(await thumbs(v).count()).toBe(5)
    await v.pick('Slide', 'Delete slide')
    await v.page.waitForFunction(() => document.querySelectorAll('.de-thumb').length === 4)
    // the undo button is on, and brings the slide back
    expect(await v.btn('Undo').isDisabled()).toBe(false)
    await v.btn('Undo').click()
    await v.page.waitForFunction(() => document.querySelectorAll('.de-thumb').length === 5)
    await v.btn('Redo').click()
    await v.page.waitForFunction(() => document.querySelectorAll('.de-thumb').length === 4)
    // keyboard: the slides have the focus
    await v.page.locator('.de-scroll').focus()
    await v.page.keyboard.press('ControlOrMeta+z')
    await v.page.waitForFunction(() => document.querySelectorAll('.de-thumb').length === 5)
    const s = await saved(v)
    expect(s.meta.count).toBe(5)
    expect(JSON.stringify((await s.slide(0)).els)).toContain('office-engine')
    expect(v.errors).toEqual([])
  })

  test('move and duplicate undo too, in reverse order', async () => {
    const v = await open()
    await v.pick('Slide', 'Duplicate slide')
    await v.page.waitForFunction(() => document.querySelectorAll('.de-thumb').length === 6)
    await v.pick('Slide', 'Move slide down')
    await v.btn('Undo').click()
    await v.btn('Undo').click()
    await v.page.waitForFunction(() => document.querySelectorAll('.de-thumb').length === 5)
    const s = await saved(v)
    expect(s.meta.count).toBe(5)
    expect(JSON.stringify((await s.slide(0)).els)).toContain('office-engine')
  })
})

describe('N2-026 / N2-027 table cells', () => {
  const cell = (v: Viewer, text: string) => frame(v, 4).locator('.de-stbl td', { hasText: text }).first()
  const table = (m: any) => m.els.find((e: any) => e.t === 'table')
  const text = (m: any, r: number, c: number) => table(m).rows[r].cells[c].text?.paras.map((p: any) => p.runs.map((x: any) => x.text).join('')).join('|') ?? ''

  test('a click on a cell edits its text, Delete edits instead of deleting the table, the text is saved', async () => {
    const v = await open()
    await goTo(v, 4)
    await cell(v, 'Users').click()
    await v.page.locator('.de-stbl .de-txt-edit').waitFor()
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.keyboard.press('Delete')
    expect(await frame(v, 4).locator('.de-stbl').count()).toBe(1)
    await v.page.keyboard.type('People')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(400)
    expect(await frame(v, 4).locator('.de-stbl').innerText()).toContain('People')
    const m = await (await saved(v)).slide(3)
    expect(text(m, 1, 0)).toBe('People')
    expect(text(m, 1, 1)).toBe('380')
    expect(text(m, 0, 0)).toBe('Metric')
    expect(v.errors).toEqual([])
  })

  test('Tab goes to the next cell; with the table selected as an object, Delete removes it', async () => {
    const v = await open()
    await goTo(v, 4)
    await cell(v, 'Users').click()
    await v.page.locator('.de-stbl .de-txt-edit').waitFor()
    await v.page.keyboard.press('Tab')
    await v.page.waitForFunction(() => document.querySelector('.de-stbl .de-txt-edit')?.textContent === '380')
    await v.page.keyboard.type('X')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(400)
    // Esc left the table selected: Delete now removes the table as an object (undoable)
    await v.page.keyboard.press('Delete')
    await v.page.waitForFunction(() => document.querySelectorAll('.de-slide-frame[data-slide="4"] .de-stbl').length === 0)
    await v.btn('Undo').click()
    await v.page.waitForFunction(() => document.querySelectorAll('.de-slide-frame[data-slide="4"] .de-stbl').length === 1)
    const m = await (await saved(v)).slide(3)
    expect(text(m, 1, 0)).toBe('Users')
    expect(text(m, 1, 1)).toBe('380X')
  })

  test('an empty cell is editable', async () => {
    const v = await open()
    await goTo(v, 4)
    await cell(v, 'Users').click()
    await v.page.locator('.de-stbl .de-txt-edit').waitFor()
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.keyboard.press('Delete')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(400)
    // the cell is empty now: a click in it still starts editing
    const box = (await frame(v, 4).locator('.de-stbl tr').nth(1).locator('td').first().boundingBox())!
    await v.page.mouse.click(box.x + 8, box.y + box.height / 2)
    await v.page.locator('.de-stbl .de-txt-edit').waitFor()
    await v.page.keyboard.type('Again')
    await v.page.keyboard.press('Escape')
    await v.page.waitForTimeout(400)
    expect(text(await (await saved(v)).slide(3), 1, 0)).toBe('Again')
  })
})

describe('P3-013 notes in the text export', () => {
  test('a converted .odp: Download as plain text has the speaker notes', async () => {
    const v = await viewer({ src: 'samples/sample.odp' })
    await v.page.waitForFunction(() => document.querySelectorAll('.de-thumb').length === 6)
    const txt = await v.downloadAs('Plain text')
    expect(txt.text).toContain('Notes:\nWelcome everyone and introduce the project.')
    expect(txt.text).toContain('Mention the lazy ZIP reader first.')
  })

  test('a converted .ppt: Download as plain text has the speaker notes', async () => {
    const v = await viewer({ src: 'samples/sample.ppt' })
    await v.page.waitForFunction(() => document.querySelectorAll('.de-thumb').length === 5)
    const txt = await v.downloadAs('Plain text')
    expect(txt.text).toContain('Notes:\nSpeaker notes: mention the worker and the lazy ZIP reader.')
  })
})
