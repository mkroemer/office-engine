/** Tab stops of a Word document: left / center / right / decimal stops, leaders, stops of a style, absolute position tabs */
import { describe, expect, test } from 'bun:test'
import { bareDocx, para } from '../test/docx-builder'
import { docxParas, paraText, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

async function open(doc: Blob | Promise<Blob>, editable = true) {
  const v = await viewer({ editable, onSave: true }, { wait: false })
  const b64 = Buffer.from(await (await doc).arrayBuffer()).toString('base64')
  await v.page.evaluate(([b]) => window.__setProps({ src: new Blob([Uint8Array.from(atob(b as string), (c) => c.charCodeAt(0))]), name: 'tabs.docx' }), [b64] as const)
  await v.ready('docx')
  if (editable) await v.page.locator('.de-ebar').first().waitFor()
  await v.page.waitForTimeout(300)
  return v
}

/** where `text` stands in its paragraph on the page: left, middle and right, in points from the text margin */
const at = (v: Viewer, text: string) =>
  v.page.evaluate((needle) => {
    for (const p of document.querySelectorAll<HTMLElement>('.de-page .de-p')) {
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT)
      for (let t = walker.nextNode(); t; t = walker.nextNode()) {
        const i = (t.textContent ?? '').indexOf(needle)
        if (i < 0) continue
        const r = document.createRange()
        r.setStart(t, i)
        r.setEnd(t, i + needle.length)
        const box = r.getBoundingClientRect()
        const pr = p.getBoundingClientRect()
        const k = (pr.width / p.offsetWidth) * (96 / 72)
        const pt = (x: number) => Math.round(((x - pr.left) / k) * 10) / 10
        return { left: pt(box.left), mid: pt((box.left + box.right) / 2), right: pt(box.right), top: Math.round(box.top) }
      }
    }
    return null
  }, text)

const tabs = (...stops: [number, string, string?][]) => `<w:tabs>${stops.map(([pos, val, leader]) => `<w:tab w:val="${val}" w:pos="${pos * 20}"${leader ? ` w:leader="${leader}"` : ''}/>`).join('')}</w:tabs>`
const tabbed = (pPr: string, ...parts: string[]) => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${parts.map((t, i) => `${i ? '<w:tab/>' : ''}<w:t xml:space="preserve">${t}</w:t>`).join('')}</w:r></w:p>`
const near = (got: number | undefined, want: number, by = 1.5) => expect(Math.abs((got ?? NaN) - want)).toBeLessThanOrEqual(by)

describe('tab stops', () => {
  test('text after a tab goes to the stop set: left of it, centered on it, ending at it; without one, to the next half inch', async () => {
    const body =
      tabbed(tabs([234, 'center'], [468, 'right']), 'Left', 'Middle', 'Ending') +
      tabbed(tabs([100, 'left'], [300, 'left']), 'aa', 'Second', 'Third', 'Past') +
      tabbed('', 'x', 'Plain')
    const v = await open(bareDocx(body))
    near((await at(v, 'Left'))?.left, 0)
    near((await at(v, 'Middle'))?.mid, 234)
    near((await at(v, 'Ending'))?.right, 468)
    near((await at(v, 'Second'))?.left, 100)
    near((await at(v, 'Third'))?.left, 300)
    // past the last stop set: the next half inch
    near((await at(v, 'Past'))?.left, 360)
    near((await at(v, 'Plain'))?.left, 36)
    // one line each
    expect((await at(v, 'Left'))?.top).toBe((await at(v, 'Ending'))?.top)
    expect(v.errors).toEqual([])
  })

  test('the stops of the style and of the paragraph add up; a stop cleared is gone; the decimal points stand on a decimal stop', async () => {
    const styles = `<w:style w:type="paragraph" w:styleId="Hdr"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr>${tabs([234, 'center'], [468, 'right'])}</w:pPr></w:style>`
    const body =
      tabbed(`<w:pStyle w:val="Hdr"/>${tabs([100, 'left'])}`, 'A', 'Own', 'Styled', 'Last') +
      tabbed(`<w:pStyle w:val="Hdr"/>${tabs([234, 'clear'])}`, 'B', 'NoMiddle') +
      tabbed(tabs([200, 'decimal']), '', '12.5') +
      tabbed(tabs([200, 'decimal']), '', '1234.25') +
      tabbed(tabs([200, 'decimal']), '', '77')
    const v = await open(bareDocx(body, { styles }))
    near((await at(v, 'Own'))?.left, 100)
    near((await at(v, 'Styled'))?.mid, 234)
    near((await at(v, 'Last'))?.right, 468)
    near((await at(v, 'NoMiddle'))?.right, 468)
    near((await at(v, '12'))?.right, 200)
    near((await at(v, '1234'))?.right, 200)
    near((await at(v, '77'))?.right, 200)
    expect(v.errors).toEqual([])
  })

  test('a leader fills the gap; an absolute position tab goes to the middle and the right of the margins, whatever the stops', async () => {
    const ptab = (a: string, text: string) => `<w:r><w:ptab w:relativeTo="margin" w:alignment="${a}" w:leader="none"/><w:t>${text}</w:t></w:r>`
    const body =
      tabbed(tabs([468, 'right', 'dot']), 'Chapter one', '12') +
      `<w:p><w:r><w:t>Company</w:t></w:r>${ptab('center', 'Report')}${ptab('right', 'Page')}</w:p>`
    const v = await open(bareDocx(body), false)
    near((await at(v, '12'))?.right, 468)
    const gap = await v.page.evaluate(() => {
      const t = [...document.querySelectorAll<HTMLElement>('.de-page .de-tabchar')].find((x) => x.dataset.ld === 'dot')
      return t ? { width: t.getBoundingClientRect().width, image: getComputedStyle(t).backgroundImage } : null
    })
    expect(gap!.width).toBeGreaterThan(200)
    expect(gap!.image).toContain('gradient')
    near((await at(v, 'Report'))?.mid, 234)
    near((await at(v, 'Page'))?.right, 468)
    // the web page exported has the tabs at their stops too
    const html = (await v.downloadAs('Web page')).text
    expect(html).toMatch(/class="de-tabchar" data-ptab="center"[^>]*data-on=""[^>]*padding-left: \d{2,}/)
    const dotted = html.match(/<span class="de-tabchar"[^>]*data-ld="dot"[^>]*>/)?.[0] ?? ''
    expect(dotted).toMatch(/padding-left: \d{3}/)
    expect(v.errors).toEqual([])
  })

  test('typing before a tab moves the tab, not the text at its stop; a tab typed goes to the next stop; the stops are saved', async () => {
    const v = await open(bareDocx(tabbed(tabs([234, 'center'], [468, 'right']), 'Left', 'Middle', 'Ending') + para('After')))
    await v.selectText('Left', 'end')
    await v.page.keyboard.type(' side grows', { delay: 20 })
    await v.page.waitForTimeout(300)
    near((await at(v, 'Middle'))?.mid, 234)
    near((await at(v, 'Ending'))?.right, 468)
    // the caret after a tab stands at the text that follows it
    await v.selectText('Middle', 'start')
    await v.page.keyboard.type('X')
    await v.page.waitForTimeout(300)
    near((await at(v, 'XMiddle'))?.mid, 234)
    const saved = await docxParas((await v.save()).bytes)
    expect(saved.map(paraText)).toEqual(['Left side grows\tXMiddle\tEnding', 'After'])
    expect(saved[0].ps.tabs).toEqual([{ pos: 234, a: 'center' }, { pos: 468, a: 'right' }])
    expect(v.errors).toEqual([])
  })
})
