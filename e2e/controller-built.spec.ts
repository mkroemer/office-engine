/** Controller + locale + host UI integration against emitted production modules. */
import { expect, test } from 'bun:test'
import { fromB64, useBrowser, viewer, xlsxSheet } from './harness'

useBrowser()

test('built controller uses current callbacks across locale/UI updates, replacement and save', async () => {
  const v = await viewer({}, { page: 'built-locale.html', wait: false, locale: 'de-DE' })
  await v.page.waitForFunction(() => !!(window as any).__builtRuntime)
  await v.page.evaluate(async ({ pkg, german }) => {
    const { React: R, createRoot } = (window as any).__builtRuntime
    const api = await import(/* @vite-ignore */ pkg)
    const de = (await import(/* @vite-ignore */ german)).de
    if ('InternalViewerController' in api || 'internalController' in api || 'de' in api) throw new Error('private or optional export leaked')
    const ctl = api.createViewerController()
    const replacement = api.createViewerController()
    ctl.subscribe(() => {})
    replacement.subscribe(() => {})
    ;(window as any).__builtControllers = { ctl, replacement }
    let current: any = { controller: ctl, locale: 'en-GB' }
    let generation = 'first'
    function Button(p: any) {
      return R.createElement('button', { 'data-built-host': '', 'aria-label': p.title, disabled: p.disabled, onMouseDown: p.onMouseDown, onClick: p.onClick }, p.icon, p.label, p.children)
    }
    const root = createRoot(document.getElementById('built')!)
    function render() {
      const callbackGeneration = generation
      root.render(R.createElement(api.OfficeUIProvider, { strings: de, components: { Button } },
        R.createElement(api.DocViewer, { src: '/samples/sample.xlsx', editable: true, headless: true, ...current,
          onLoad: (c: any) => window.__events.push({ type: 'load', kind: c.kind }),
          onError: (e: Error) => window.__events.push({ type: 'error', message: e.message }),
          onSave: async (blob: Blob, name: string) => { let raw = ''; for (const b of new Uint8Array(await blob.arrayBuffer())) raw += String.fromCharCode(b); window.__events.push({ type: 'save', name, b64: btoa(raw), generation: callbackGeneration }) },
        })))
    }
    ;(window as any).__builtUpdate = (p: any) => { current = { ...current, ...p }; generation = 'updated'; render() }
    ;(window as any).__builtUnmount = () => root.unmount()
    render()
  }, { pkg: `/@fs${process.cwd()}/dist/index.js`, german: `/@fs${process.cwd()}/dist/locales/de.js` })
  await v.ready('xlsx')
  await v.page.locator('button[data-built-host][aria-label^="Fett"]').waitFor()
  await v.page.waitForFunction(() => (window as any).__builtControllers.ctl.snapshot().status.includes('Zeilen'))
  await v.typeCell('J12', 'production integration')
  await v.page.waitForFunction(() => (window as any).__builtControllers.ctl.snapshot().canUndo)
  await v.page.evaluate(() => (window as any).__builtUpdate({ locale: 'vi-VN' }))
  await v.page.evaluate(() => { const { undo } = (window as any).__builtControllers.ctl; undo() })
  await v.page.waitForFunction(() => (window as any).__builtControllers.ctl.snapshot().canRedo)
  expect(await v.shown('J12')).toBe('')
  await v.page.evaluate(() => { const { redo } = (window as any).__builtControllers.ctl; redo() })
  await v.page.waitForFunction(() => (window as any).__builtControllers.ctl.snapshot().canUndo)
  await v.page.evaluate(() => (window as any).__builtControllers.ctl.save())
  const saved = await v.waitEvent('save')
  expect(saved.generation).toBe('updated')
  expect((await xlsxSheet(fromB64(saved.b64 as string))).cell('J12')).toBe('production integration')
  await v.page.evaluate(() => (window as any).__builtUpdate({ controller: (window as any).__builtControllers.replacement }))
  await v.page.waitForFunction(() => {
    const { ctl, replacement } = (window as any).__builtControllers
    return ctl.snapshot().kind === null && replacement.snapshot().kind === 'xlsx' && replacement.snapshot().outline?.items.length > 0
  })
  await v.page.evaluate(() => (window as any).__builtControllers.replacement.goto('1'))
  await v.page.waitForFunction(() => (window as any).__builtControllers.replacement.snapshot().outline.items[1].current)
  expect(await v.events('load')).toHaveLength(1)
  await v.page.evaluate(() => (window as any).__builtUnmount())
  expect(await v.page.evaluate(() => (window as any).__builtControllers.replacement.snapshot().kind)).toBeNull()
  expect(v.errors).toEqual([])
})
