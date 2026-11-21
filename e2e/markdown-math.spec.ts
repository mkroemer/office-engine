/** TeX math in Markdown (#33): drawn as MathML, edited around as an object, saved back as written, printed and exported */
import { describe, expect, test } from 'bun:test'
import { docxParas, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const MD = `# Notes

Euler’s identity: $e^{i\\pi} + 1 = 0$ is famous.

$$
\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}
$$

It costs $5 and \\$10, see \`$x$\`.

Broken: $\\frac{a}{$ and $\\foo x$ end.
`

/** opens `md` as notes.md through File → Open (a viewer that can edit and save) */
async function openMd(md = MD): Promise<Viewer> {
  const v = await viewer({ src: 'samples/sample.xlsx', onSave: true, editable: true })
  await v.ctx.grantPermissions(['clipboard-read', 'clipboard-write'])
  await v.chooseFile(() => v.menu('Open…'), { name: 'notes.md', mimeType: 'text/markdown', buffer: Buffer.from(md) })
  await v.page.waitForFunction(() => document.querySelector('.de-bar .de-title')?.textContent === 'notes.md')
  await v.page.locator('.de-page-inner math').first().waitFor()
  await v.page.waitForTimeout(400)
  return v
}

/** the TeX of the formulas on the pages, in order */
const texts = (v: Viewer) => v.page.locator('.de-pages annotation').allTextContents()

describe('Markdown math in the document', () => {
  test('inline and display formulas are MathML; currency, escapes and code stay text; broken ones show their source', async () => {
    const v = await openMd()
    const { page } = v
    if (process.env.MATH_SHOTS) await page.locator('.de-page').first().screenshot({ path: `${process.env.MATH_SHOTS}/markdown-math.png` })
    expect(await page.locator('.de-pages math:not([display])').count()).toBe(3)
    expect(await page.locator('.de-pages math[display="block"]').count()).toBe(1)
    expect(await texts(v)).toEqual(['e^{i\\pi} + 1 = 0', '\n\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}\n', '\\frac{a}{', '\\foo x'])
    // drawn: the fraction with its parts, the display formula a block of its own, taller than a line
    expect(await page.locator('.de-pages math[display="block"] mfrac msqrt msup').count()).toBe(1)
    const box = (await page.locator('.de-pages math[display="block"]').boundingBox())!
    expect(box.height).toBeGreaterThan(30)
    expect(await page.locator('.de-pages .de-math[data-display]').evaluate((e) => getComputedStyle(e).display)).toBe('block')
    // an equation is one object of the text: not editable inside
    expect(await page.locator('.de-pages .de-math').first().getAttribute('data-obj')).toBe('')
    const text = await v.pagesText()
    expect(text).toContain('It costs $5 and $10, see $x$.')
    expect(text).not.toContain('\\frac{-b')
    // the broken ones: their source, marked, with the reason as the tooltip; the page goes on
    expect(await page.locator('.de-pages .de-math[data-error]').count()).toBe(2)
    expect(await page.locator('.de-pages .de-math[data-error]').first().getAttribute('title')).toBe('Missing }')
    expect(await page.locator('.de-pages merror').allTextContents()).toEqual(['\\frac{a}{', '\\foo'])
    expect(text).toContain('end.')
    // opened is not unsaved
    expect(await page.locator('.de-save-dirty').count()).toBe(0)
    expect(v.errors).toEqual([])
  })

  test('typing next to an equation and saving gives back the formulas as written', async () => {
    const v = await openMd()
    await v.edit()
    await v.selectText('Euler’s identity', 'end')
    await v.page.keyboard.type(' (famous)')
    await v.selectText(' is famous', 'end')
    await v.page.keyboard.type('!')
    await v.page.waitForTimeout(300)
    const saved = await v.save()
    expect(saved.name).toBe('notes.md')
    expect(saved.bytes.toString('utf8')).toBe(MD.replace('Euler’s identity: $e^{i\\pi} + 1 = 0$ is famous.', 'Euler’s identity (famous): $e^{i\\pi} + 1 = 0$ is famous!.'))
    // the Word file: real equations holding their TeX
    const word = await v.downloadAs('Word document')
    const ps = await docxParas(word.bytes)
    const eq = ps[1].runs.find((r) => r.t === 'ref')
    expect(eq).toMatchObject({ t: 'ref', math: { tex: 'e^{i\\pi} + 1 = 0', delim: '$' } })
    expect(v.errors).toEqual([])
  })

  test('the equation is one character for the caret: stepped over, deleted, undone, copied and pasted', async () => {
    const v = await openMd()
    const { page } = v
    await v.edit()
    const inline = () => page.locator('.de-pages math:not([display])').count()
    // the caret before it, a step to the right is after it
    await v.selectText('identity: ', 'end')
    await page.keyboard.press('ArrowRight')
    await page.keyboard.type('X')
    await page.waitForTimeout(200)
    expect(await page.locator('.de-pages .de-p', { hasText: 'Euler' }).innerText()).toContain('X is famous')
    await page.keyboard.press('Control+z')
    await page.waitForTimeout(200)
    // deleted as one, and back with Undo
    await v.selectText('identity: ', 'end')
    await page.keyboard.press('Delete')
    await page.waitForTimeout(200)
    expect(await inline()).toBe(2)
    expect(await texts(v)).not.toContain('e^{i\\pi} + 1 = 0')
    await page.keyboard.press('Control+z')
    await page.waitForTimeout(300)
    expect(await inline()).toBe(3)
    expect(await texts(v)).toContain('e^{i\\pi} + 1 = 0')
    // copied with the text around it, pasted in another paragraph: a formula again
    await page.evaluate(() => {
      const root = document.querySelector<HTMLElement>('.de-pages')!
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      let a: Text | null = null
      let b: Text | null = null
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if (!a && n.textContent!.includes('Euler’s identity')) a = n as Text
        if (!b && n.textContent!.includes(' is famous')) b = n as Text
      }
      root.focus()
      getSelection()!.setBaseAndExtent(a!, 0, b!, b!.textContent!.indexOf(' is famous'))
    })
    await page.keyboard.press('Control+c')
    await v.selectText('It costs', 'start')
    await page.keyboard.press('Control+v')
    await page.waitForTimeout(500)
    expect(await inline()).toBe(4)
    expect((await texts(v)).filter((t) => t === 'e^{i\\pi} + 1 = 0')).toHaveLength(2)
    const saved = (await v.save()).bytes.toString('utf8')
    // (the paragraph is written again: its dollar signs are escaped only where they would be read as math)
    expect(saved).toContain('Euler’s identity: $e^{i\\pi} + 1 = 0$It costs $5 and $10, see `$x$`.')
    expect(saved).toContain('Euler’s identity: $e^{i\\pi} + 1 = 0$ is famous.')
    // (Save took the focus)
    await v.selectText('It costs', 'start')
    await page.keyboard.press('Control+z')
    await page.waitForTimeout(300)
    expect(await inline()).toBe(3)
    expect(v.errors).toEqual([])
  })

  test('the source pane: a formula typed in the Markdown is drawn', async () => {
    const v = await openMd()
    const { page } = v
    await page.getByTestId('de-code-toggle').click()
    const input = page.getByTestId('de-source-input')
    await input.waitFor()
    expect(await input.inputValue()).toBe(MD)
    await input.click()
    await page.keyboard.press('Control+End')
    await page.keyboard.type('\nRoot: $\\sqrt{2}$\n')
    await page.waitForFunction(() => [...document.querySelectorAll('.de-pages annotation')].some((a) => a.textContent === '\\sqrt{2}'))
    expect(await page.locator('.de-pages msqrt').count()).toBe(2)
    expect(v.errors).toEqual([])
  })

  test('print and the web page export carry the equations; text and Markdown exports their TeX', async () => {
    const v = await openMd()
    const { page } = v
    // headless print is a no-op: what reaches the print frame
    await page.evaluate(() => {
      ;(window as any).__printed = []
      new MutationObserver((ms) => {
        for (const m of ms)
          for (const n of m.addedNodes) {
            if ((n as HTMLElement).tagName !== 'IFRAME') continue
            const w = (n as HTMLIFrameElement).contentWindow!
            w.print = () => {
              const d = w.document
              ;(window as any).__printed.push({
                inline: d.querySelectorAll('math:not([display])').length,
                display: d.querySelectorAll('math[display="block"]').length,
                tex: [...d.querySelectorAll('annotation')].map((a) => a.textContent),
                height: d.querySelector('math[display="block"]')?.getBoundingClientRect().height ?? 0,
              })
            }
            w.document.execCommand = () => false
          }
      }).observe(document.body, { childList: true })
    })
    await v.menu('Print')
    await page.waitForFunction(() => (window as any).__printed?.length > 0, null, { timeout: 30_000 })
    const printed = await page.evaluate(() => (window as any).__printed.at(-1))
    expect(printed).toMatchObject({ inline: 3, display: 1 })
    expect(printed.tex[0]).toBe('e^{i\\pi} + 1 = 0')
    expect(printed.height).toBeGreaterThan(30)
    const web = await v.downloadAs('Web page')
    expect(web.name).toBe('notes.html')
    expect(web.text).toContain('<annotation encoding="application/x-tex">e^{i\\pi} + 1 = 0</annotation>')
    expect(web.text).toMatch(/<math display="block"[^>]*><semantics><mfrac>/)
    // the export's style sheet has the math rules (Chromium draws table columns from CSS)
    expect(web.text).toContain('.de-math merror')
    const txt = await v.downloadAs('Plain text')
    expect(txt.text).toContain('Euler’s identity: $e^{i\\pi} + 1 = 0$ is famous.')
    const md = await v.downloadAs('Markdown')
    expect(md.text).toContain('$$\n\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}\n$$')
    expect(v.errors).toEqual([])
  })

  test('TeX that tries to be markup stays text', async () => {
    const v = await openMd('A $<img src=x onerror="window.__hit=1">$ and $\\href{javascript:alert(2)}{x}$ and $\\color{red;background:url(https://example.com/x)}{y}$.\n')
    const { page } = v
    expect(await page.evaluate(() => (window as any).__hit)).toBeUndefined()
    expect(await page.locator('.de-pages img, .de-pages a[href], .de-pages [onerror]').count()).toBe(0)
    expect(await page.locator('.de-pages math [mathcolor]').count()).toBe(0)
    // (the first one does not read as TeX: its source, as text)
    expect(await page.locator('.de-pages merror').allTextContents()).toEqual(['<img src=x onerror="window.__hit=1">', '\\href'])
    expect(v.errors).toEqual([])
  })
})
