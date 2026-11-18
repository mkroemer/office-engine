/** Host formatting through the emitted library and real worker, including save/reopen. */
import { expect, test } from 'bun:test'
import { strFromU8, unzipSync } from 'fflate'
import { docxParas, fromB64, paraText, runsOf, useBrowser, viewer, xlsxSheet, type Viewer } from './harness'

useBrowser()

async function select(v: Viewer, kind: 'docx' | 'xlsx' | 'pptx') {
  if (kind === 'xlsx') await v.clickCell('B2')
  else if (kind === 'docx') await v.selectText('document')
  else {
    await v.page.locator('.de-slide-frame').first().getByText('office-engine', { exact: true }).click()
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.keyboard.press('ControlOrMeta+a')
  }
  await v.page.waitForTimeout(150)
}

for (const kind of ['docx', 'xlsx', 'pptx'] as const) {
  test(`built ${kind}: inline React host formatting preserves focus, history and saved/reopened values`, async () => {
    const v = await viewer({}, { page: 'built-locale.html', wait: false })
    await v.page.waitForFunction(() => !!(window as any).__builtRuntime)
    await v.page.evaluate(async ({ pkg, kind }) => {
      const { React: R, createRoot } = (window as any).__builtRuntime
      const api = await import(/* @vite-ignore */ pkg)
      const ctl = api.createViewerController()
      window.__ctl = ctl
      const root = createRoot(document.getElementById('built')!)
      function Host() {
        const f = api.useViewerState(ctl).format
        const [props, setProps] = R.useState({ src: `/samples/sample.${kind}`, name: `sample.${kind}` })
        window.__setProps = (p: any) => setProps((old: any) => ({ ...old, ...p }))
        return R.createElement(R.Fragment, null,
          R.createElement('input', { 'aria-label': 'Production size', type: 'number', disabled: !f,
            onChange: (e: any) => ctl.format({ op: 'size', pt: e.currentTarget.valueAsNumber }) }),
          R.createElement('button', { 'aria-label': 'Production undo', onClick: ctl.undo }, 'Undo'),
          R.createElement(api.DocViewer, { ...props, controller: ctl, editable: true, headless: true,
            onLoad: (c: any) => window.__events.push({ type: 'load', kind: c.kind }),
            onError: (e: Error) => window.__events.push({ type: 'error', message: e.message }),
            onSave: async (blob: Blob, name: string) => {
              let raw = ''
              for (const b of new Uint8Array(await blob.arrayBuffer())) raw += String.fromCharCode(b)
              window.__events.push({ type: 'save', name, b64: btoa(raw) })
            },
          }))
      }
      root.render(R.createElement(Host))
    }, { pkg: `/@fs${process.cwd()}/dist/index.js`, kind })
    await v.ready(kind)
    const cellBefore = kind === 'xlsx' ? await v.shown('B2') : undefined
    const inputBefore = kind === 'xlsx' ? await v.formula('B2') : undefined
    await select(v, kind)
    const input = v.page.getByLabel('Production size', { exact: true })
    await input.click()
    await v.page.keyboard.type('14', { delay: 120 })
    await v.page.waitForTimeout(200)
    expect(await input.inputValue()).toBe('14')
    expect(await input.evaluate((el) => el === document.activeElement)).toBe(true)
    await v.page.keyboard.press('Backspace')
    await v.page.waitForTimeout(150)
    expect(await input.inputValue()).toBe('1')
    expect(await input.evaluate((el) => el === document.activeElement)).toBe(true)
    await v.page.keyboard.type('4')
    await v.page.waitForTimeout(150)
    expect(await v.page.evaluate(() => window.__ctl.snapshot().format?.size)).toBe(14)
    await v.page.getByLabel('Production undo', { exact: true }).click()
    await v.page.waitForTimeout(150)
    expect(await v.page.evaluate(() => window.__ctl.snapshot().format?.size)).toBe(1)
    if (kind !== 'xlsx') expect(await v.page.evaluate(() => document.activeElement?.matches('.de-pages, .de-txt-edit'))).toBe(true)
    await v.page.evaluate(() => window.__ctl.redo())
    await v.page.waitForTimeout(150)
    expect(await v.page.evaluate(() => window.__ctl.snapshot().format?.size)).toBe(14)
    // Main's save flush must include the last PPTX command without waiting for the text debounce.
    await v.page.evaluate(() => window.__ctl.save())
    const bytes = fromB64((await v.waitEvent('save')).b64 as string)
    const part = kind === 'docx' ? 'word/document.xml' : kind === 'pptx' ? 'ppt/slides/slide1.xml' : 'xl/styles.xml'
    const xml = strFromU8(unzipSync(bytes)[part])
    expect(xml).toContain(kind === 'docx' ? 'w:sz w:val="28"' : kind === 'pptx' ? 'sz="1400"' : 'val="14"')
    if (kind === 'docx') {
      const p = (await docxParas(bytes)).find((p) => paraText(p).startsWith('This document'))!
      expect(runsOf(p).find((r) => r.text === 'document')?.s.sz).toBe(14)
    }
    if (kind === 'xlsx') expect((await xlsxSheet(bytes)).cell('B2')).toBe(cellBefore!)
    await v.page.evaluate((data) => window.__setProps({ src: new Blob([new Uint8Array(data)]), name: 'reopened.' + window.__ctl.snapshot().kind }), Array.from(bytes))
    await v.waitEvent('load', 2)
    await v.page.waitForTimeout(300)
    await select(v, kind)
    expect(await v.page.evaluate(() => window.__ctl.snapshot().format?.size)).toBe(14)
    if (kind === 'xlsx') expect(await v.formula('B2')).toBe(inputBefore!)
    expect(v.errors).toEqual([])
  })
}
