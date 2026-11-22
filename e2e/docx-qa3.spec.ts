/** QA round 3 regressions of the Word editor: what a file from elsewhere holds that the samples do not */
import { describe, expect, test } from 'bun:test'
import { bareDocx, para } from '../test/docx-builder'
import { docxParas, paraText, useBrowser, viewer } from './harness'

useBrowser()

/** opens a built document in the fixture, editable */
async function open(doc: Blob | Promise<Blob>, name = 'qa3.docx') {
  const v = await viewer({ editable: true, onSave: true }, { wait: false })
  const b64 = Buffer.from(await (await doc).arrayBuffer()).toString('base64')
  await v.page.evaluate(([b, n]) => window.__setProps({ src: new Blob([Uint8Array.from(atob(b as string), (c) => c.charCodeAt(0))]), name: n as string }), [b64, name] as const)
  await v.ready('docx')
  await v.page.locator('.de-ebar').first().waitFor()
  return v
}

/** a text box (VML, as Word writes the fallback of one) holding `text` */
const textBox = (text: string) =>
  `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:180pt;height:40pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>`

describe('QA 3: a text box anchored in a paragraph', () => {
  test('typing in the paragraph goes on after the first character (the caret does not move into the box)', async () => {
    const v = await open(bareDocx(`<w:p>${textBox('Inside the box, longer than the line typed')}<w:r><w:t>Body text</w:t></w:r></w:p>` + para('After')))
    await v.selectText('Body text', 'end')
    await v.page.keyboard.type(' typed', { delay: 30 })
    await v.page.waitForTimeout(300)
    const at = await v.page.evaluate(() => {
      const n = getSelection()!.anchorNode
      const el = n && (n.nodeType === 1 ? (n as Element) : n.parentElement)
      return { inBox: !!el?.closest('.de-tbox'), text: el?.closest('[data-pa]')?.textContent }
    })
    expect(at.inBox).toBe(false)
    const paras = await docxParas((await v.save()).bytes)
    expect(paras.map(paraText)).toEqual(['Body text typed', 'After'])
    // Backspace and typing in the middle keep their place too
    await v.selectText('Body', 'end')
    await v.page.keyboard.type('XY', { delay: 30 })
    await v.page.keyboard.press('Backspace')
    await v.page.waitForTimeout(300)
    expect((await docxParas((await v.save()).bytes)).map(paraText)).toEqual(['BodyX text typed', 'After'])
    expect(v.errors).toEqual([])
  })
})

describe('QA 3: lists of a file that defines one level only', () => {
  test('Tab on an item moves it a level in and keeps its bullet; Shift+Tab brings it back', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    const item = async () => (await docxParas((await v.save()).bytes)).find((p) => paraText(p) === 'Two new regions launched')!
    expect((await item()).marker?.text).toBeTruthy()
    await v.selectText('Two new regions launched', 'end')
    await v.page.keyboard.press('Tab')
    await v.page.waitForTimeout(400)
    const inner = await item()
    expect(inner.ilvl).toBe(1)
    expect(inner.marker?.text).toBeTruthy()
    // the marker is on the page too, and the other items kept theirs
    const shown = await v.page.evaluate(() => [...document.querySelectorAll('[data-pa]')].filter((p) => /Revenue grew|Two new regions|Customer churn/.test(p.textContent ?? '')).map((p) => (p.textContent ?? '').trim().slice(0, 1)))
    expect(shown.length).toBe(3)
    expect(shown.every((c) => c && !/[A-Za-z]/.test(c))).toBe(true)
    await v.selectText('Two new regions launched', 'end')
    await v.page.keyboard.press('Shift+Tab')
    await v.page.waitForTimeout(400)
    const back = await item()
    expect(back.ilvl ?? 0).toBe(0)
    expect(back.marker?.text).toBeTruthy()
    expect(v.errors).toEqual([])
  })
})
