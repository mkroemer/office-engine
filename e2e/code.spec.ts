import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { docxParas, SAMPLES, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const MD =
  '# Notes\n\n```javascript\nconst greeting = "Hello";\nconsole.log(greeting);\n```\n\n```cobol\nMOVE A TO B.\n```\n\n```mermaid\nflowchart LR\n  A --> B\n```\n\nEnd.\n'

/** opens `text` as a file named `name` (a Blob, as a host passes one) */
async function openText(v: Viewer, text: string, name: string, extra: Record<string, unknown> = {}) {
  await v.page.evaluate(([t, n, x]) => window.__setProps({ src: new Blob([t as string]), name: n as string, ...(x as object) }), [text, name, extra] as const)
  await v.page.waitForFunction((n) => document.querySelector('.de-bar .de-title')?.textContent === n, name)
}

/** WCAG contrast of two CSS rgb() colors */
function contrast(a: string, b: string): number {
  const lum = (c: string) => {
    const [r, g, bl] = (c.match(/\d+(\.\d+)?/g) ?? []).slice(0, 3).map((x) => {
      const v = Number(x) / 255
      return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4
    })
    return 0.2126 * r + 0.7152 * g + 0.0722 * bl
  }
  const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p)
  return (x + 0.05) / (y + 0.05)
}

/** the color of each token class shown in `scope`, with the background behind it */
async function tokenColors(v: Viewer, scope: string, bg: string) {
  return v.page.evaluate(
    ([s, b]) => {
      const out: Record<string, [string, string]> = {}
      const back = getComputedStyle(document.querySelector(b)!).backgroundColor
      for (const el of document.querySelectorAll<HTMLElement>(`${s} [class^="de-tk-"]`)) out[el.className] ??= [getComputedStyle(el).color, back]
      return out
    },
    [scope, bg] as const,
  )
}

describe('syntax highlighting', () => {
  test('a Markdown code block is colored by its language; opening it changes nothing', async () => {
    // (the dark theme: the page stays paper, its code keeps colors readable on it)
    const v = await viewer({ src: 'samples/sample.docx', onSave: true, theme: 'dark' })
    const { page } = v
    await openText(v, MD, 'notes.md')
    const block = page.locator('.de-page-inner .de-p.de-code[data-code="javascript"]')
    await block.waitFor()
    expect(await block.locator('.de-tk-k', { hasText: 'const' }).count()).toBe(1)
    expect(await block.locator('.de-tk-s').first().innerText()).toBe('"Hello"')
    expect(await block.locator('.de-tk-f', { hasText: 'log' }).count()).toBe(1)
    // the text is the code, as it is
    expect(await block.innerText()).toBe('const greeting = "Hello";\nconsole.log(greeting);')
    // a language the viewer does not know: plain code; the diagram is drawn
    const cobol = page.locator('.de-page-inner .de-p.de-code[data-code="cobol"]')
    expect(await cobol.innerText()).toBe('MOVE A TO B.')
    expect(await cobol.locator('[class^="de-tk-"]').count()).toBe(0)
    expect(await page.locator('.de-page-inner img').count()).toBe(1)
    expect(await page.locator('.de-pages').innerText()).not.toContain('flowchart')
    // paper colors on the page: readable on white in the dark theme too
    const colors = await tokenColors(v, '.de-p.de-code', '.de-page')
    expect(Object.keys(colors).sort()).toEqual(['de-tk-f', 'de-tk-k', 'de-tk-s'])
    for (const [fg, bg] of Object.values(colors)) expect(contrast(fg, bg)).toBeGreaterThan(4.5)
    // not changed by being shown: no unsaved changes, nothing to undo
    await v.edit()
    await v.selectText('greeting', 'end')
    await page.keyboard.press('Control+z')
    await page.waitForTimeout(300)
    expect(await block.innerText()).toBe('const greeting = "Hello";\nconsole.log(greeting);')
    expect((await v.events('dirty')).filter((e) => e.dirty)).toEqual([])
    expect(await page.locator('.de-save-dirty').count()).toBe(0)
    // File → Save: the file as it was
    const n = (await v.events('save')).length
    await v.menu('Save')
    const saved = await v.waitEvent('save', n + 1)
    expect(Buffer.from(saved.b64 as string, 'base64').toString('utf8')).toBe(MD)
    // the source beside the document: the block colored in its language there too
    await page.getByTestId('de-code-toggle').click()
    const hl = page.locator('[data-testid=de-source] .de-src-hl')
    await hl.waitFor()
    expect(await hl.locator('.de-tk-k', { hasText: 'const' }).count()).toBe(1)
    expect(await hl.locator('.de-tk-s', { hasText: '"Hello"' }).count()).toBe(1)
    expect(v.errors).toEqual([])
  })

  test('editing inside a code block: the colors follow, Enter is a new line of code, saved as the right Markdown', async () => {
    const v = await viewer({ src: 'samples/sample.docx', onSave: true })
    const { page } = v
    await openText(v, MD, 'notes.md')
    await page.locator('.de-p.de-code[data-code="javascript"] .de-tk-s').waitFor()
    await v.edit()
    await v.selectText('"Hello"', 'end')
    await page.keyboard.type(' + "!"')
    await page.keyboard.press('End')
    await page.keyboard.press('Enter')
    await page.keyboard.type('done()')
    const block = page.locator('.de-p.de-code[data-code="javascript"]')
    await page.waitForFunction(() => document.querySelector('.de-p.de-code[data-code="javascript"]')?.textContent?.includes('done()'))
    // still one block, colored as typed
    expect(await page.locator('.de-pages .de-p.de-code').count()).toBe(2)
    expect(await block.locator('.de-tk-s', { hasText: '"!"' }).count()).toBe(1)
    expect(await block.locator('.de-tk-f', { hasText: 'done' }).count()).toBe(1)
    // Enter on the empty last line: out of the block, into a new paragraph
    await v.selectText('console', 'start')
    await page.keyboard.press('End')
    await page.keyboard.press('Enter')
    await page.keyboard.press('Enter')
    await page.keyboard.type('Next')
    await page.waitForFunction(() => [...document.querySelectorAll('.de-p:not(.de-code)')].some((p) => p.textContent === 'Next'))
    const saved = await v.save()
    expect(saved.name).toBe('notes.md')
    expect(saved.bytes.toString('utf8')).toBe(
      MD.replace('const greeting = "Hello";\n', 'const greeting = "Hello" + "!";\ndone()\n').replace('console.log(greeting);\n```\n\n', 'console.log(greeting);\n```\n\nNext\n\n'),
    )
    // the Word document saved from it: the code, no colors
    const word = await v.downloadAs('Word document')
    const ps = await docxParas(word.bytes)
    const js = ps.find((p) => p.code === 'javascript')!
    expect(js.runs.map((r) => (r.t === 'text' ? r.text : '')).join('')).toBe('const greeting = "Hello" + "!";\ndone()\nconsole.log(greeting);')
    expect(js.runs.every((r) => r.t !== 'text' || !r.d?.color)).toBe(true)
    expect(v.errors).toEqual([])
  })

  test('a .ts file opens in the code view: colored, edited as text, saved as it is', async () => {
    const v = await viewer({ src: 'samples/sample.ts', onSave: true })
    const { page } = v
    const code = page.getByTestId('de-code')
    await code.waitFor()
    expect(await v.title()).toBe('sample.ts')
    expect(await page.locator('.de-badge').innerText()).toBe('TS')
    expect(await code.getAttribute('data-language')).toBe('typescript')
    expect(await page.locator('.de-page-inner').count()).toBe(0)
    expect(await page.getByTestId('de-code-toggle').count()).toBe(0)
    expect(await code.locator('.de-src-title').innerText()).toBe('TypeScript')
    expect(await code.locator('.de-src-hl .de-tk-k', { hasText: 'interface' }).count()).toBe(1)
    expect(await code.locator('.de-src-hl .de-tk-y', { hasText: 'Reading' }).count()).toBeGreaterThan(0)
    expect(await code.locator('.de-src-hl .de-tk-c').count()).toBeGreaterThan(0)
    const input = page.getByTestId('de-source-input')
    const file = readFileSync(`${SAMPLES}/sample.ts`, 'utf8')
    expect(await input.inputValue()).toBe(file)
    // view: read only; not edited, Save gives the file as it is
    expect(await input.getAttribute('readonly')).not.toBeNull()
    const n = (await v.events('save')).length
    await v.menu('Save')
    const same = await v.waitEvent('save', n + 1)
    expect(same.name).toBe('sample.ts')
    expect(Buffer.from(same.b64 as string, 'base64').equals(readFileSync(`${SAMPLES}/sample.ts`))).toBe(true)
    // edit: the text changes, the colors follow, the plain text is saved
    await v.barButton('Edit').click()
    expect(await input.getAttribute('readonly')).toBeNull()
    const edited = file.replace('const SITES', 'const PLACES').replace('SITES.includes', 'PLACES.includes')
    await input.fill(edited)
    await page.waitForFunction(() => document.querySelector('[data-testid=de-code] .de-src-hl')?.textContent?.includes('PLACES'))
    expect(await page.locator('.de-save-dirty').count()).toBe(1)
    const saved = await v.save()
    expect(saved.name).toBe('sample.ts')
    // the plain text as edited, nothing of the colors
    expect(saved.bytes.toString('utf8')).toBe(edited)
    // File → Download as: the code in a Word document (one code paragraph, no colors), as a Markdown code block
    const word = await v.downloadAs('Word document')
    expect(word.name).toBe('sample.docx')
    const ps = await docxParas(word.bytes)
    expect(ps.map((p) => p.code)).toEqual(['typescript'])
    expect(ps[0].runs.map((r) => (r.t === 'text' ? r.text : '')).join('')).toBe(edited)
    expect(ps[0].runs.every((r) => r.t !== 'text' || !r.d)).toBe(true)
    const md = await v.downloadAs('Markdown')
    expect(md.text).toStartWith('```typescript\n/**')
    expect(v.errors).toEqual([])
  })

  test('a .py file, a language hint, an unknown language; readable colors in the dark theme', async () => {
    const v = await viewer({ src: 'samples/sample.docx', theme: 'dark' })
    const { page } = v
    await openText(v, 'import os\n\n@cache\ndef home(path: str) -> str:\n    """Doc"""\n    return f"{path}/x"  # note\n', 'script.py')
    const code = page.getByTestId('de-code')
    await code.waitFor()
    expect(await code.getAttribute('data-language')).toBe('python')
    expect(await code.locator('.de-tk-m', { hasText: '@cache' }).count()).toBe(1)
    expect(await code.locator('.de-tk-f', { hasText: 'home' }).count()).toBe(1)
    const colors = await tokenColors(v, '[data-testid=de-code] .de-src-hl', '[data-testid=de-code] .de-src-scroll')
    expect(Object.keys(colors).length).toBeGreaterThanOrEqual(5)
    for (const [fg, bg] of Object.values(colors)) expect(contrast(fg, bg)).toBeGreaterThan(4.5)
    // the dark theme's colors, not the light ones
    expect(colors['de-tk-k'][0]).toBe('rgb(210, 168, 255)')
    // the host says what the file is
    await openText(v, 'SELECT id FROM users -- all\n', 'query.txt', { language: 'sql' })
    await page.waitForFunction(() => document.querySelector('[data-testid=de-code]')?.getAttribute('data-language') === 'sql')
    expect(await code.locator('.de-tk-k', { hasText: 'SELECT' }).count()).toBe(1)
    // a language with no colors: plain code
    await openText(v, 'MOVE A TO B.\n', 'prog.txt', { language: 'cobol' })
    await page.waitForFunction(() => document.querySelector('.de-src-title')?.textContent === 'Source')
    expect(await code.locator('[class^="de-tk-"]').count()).toBe(0)
    expect(await code.locator('.de-src-hl').innerText()).toContain('MOVE A TO B.')
    expect(v.errors).toEqual([])
  })

  test('a source file keeps its bytes when not edited, its CRLF line breaks when edited', async () => {
    const v = await viewer({ src: 'samples/sample.docx', onSave: true, editable: true })
    const { page } = v
    const text = '\uFEFFx = 1\r\ny = 2\r\n'
    await openText(v, text, 'crlf.py')
    const input = page.getByTestId('de-source-input')
    await input.waitFor()
    expect(await input.inputValue()).toBe('x = 1\ny = 2\n')
    const n = (await v.events('save')).length
    await v.menu('Save')
    const same = await v.waitEvent('save', n + 1)
    expect(Buffer.from(same.b64 as string, 'base64').equals(Buffer.from(text, 'utf8'))).toBe(true)
    await input.fill('x = 1\ny = 3\n')
    const saved = await v.save()
    expect(saved.name).toBe('crlf.py')
    expect(saved.bytes.toString('utf8')).toBe('x = 1\r\ny = 3\r\n')
    expect(v.errors).toEqual([])
  })

  test('<office-viewer language>: the hint from the element', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    await v.page.evaluate(() => {
      const wc = document.getElementById('wc')!
      wc.setAttribute('language', 'python')
      wc.setAttribute('src', 'samples/sample.ts')
    })
    const code = v.page.locator('#wc [data-testid=de-code]')
    await code.waitFor()
    expect(await code.getAttribute('data-language')).toBe('python')
  })
})
