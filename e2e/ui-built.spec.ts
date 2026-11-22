/** Exercise host UI/provider composition against emitted ESM, after `bun run build`. */
import { expect, test } from 'bun:test'
import { docxParas, paraText, runsOf, useBrowser, viewer } from './harness'

import { builtModule, packageRoot, useBuiltPackage } from './built-package'

useBuiltPackage()
useBrowser()

test('built UI inherits German through formatting-only locales and defined components, preserving real edits', async () => {
  const v = await viewer({}, { page: 'built-locale.html', wait: false, locale: 'en-GB' })
  await v.page.waitForFunction(() => !!(window as any).__builtRuntime)
  await v.page.evaluate(async ({ pkg, german }) => {
    const { React: R, createRoot } = (window as any).__builtRuntime
    const api = await import(/* @vite-ignore */ pkg)
    if ('de' in api || 'LOCALES' in api) throw new Error('optional German leaked into root exports')
    const de = (await import(/* @vite-ignore */ german)).de
    function Button(p: any) {
      return R.createElement('button', { 'data-built-host': '', 'aria-label': p.title, 'data-active': p.active ? '' : undefined, disabled: p.disabled, onMouseDown: p.onMouseDown, onClick: p.onClick }, p.icon, p.label, p.children)
    }
    const root = createRoot(document.getElementById('built')!)
    let current: any = { src: '/samples/sample.docx', editable: true, locale: 'en-GB' }
    const onLoad = (c: any) => window.__events.push({ type: 'load', kind: c.kind })
    const onError = (e: Error) => window.__events.push({ type: 'error', message: e.message })
    function render() {
      root.render(R.createElement(api.OfficeUIProvider, { strings: de, components: { Button } },
        R.createElement(api.OfficeUIProvider, { locale: 'fr-FR', components: { Button: undefined } },
          R.createElement(api.DocViewer, { ...current, strings: { file: { menu: 'Ablage' } },
            onLoad, onError,
            onSave: async (f: File) => { const bytes = new Uint8Array(await f.arrayBuffer()); let raw = ''; for (const b of bytes) raw += String.fromCharCode(b); window.__events.push({ type: 'save', b64: btoa(raw) }) },
          }))))
    }
    ;(window as any).__builtUpdate = (p: any) => { current = { ...current, ...p }; render() }
    render()
  }, { pkg: builtModule('index.js'), german: builtModule('locales/de.js') })
  await v.ready('docx')
  expect(await v.page.locator('.de-file').innerText()).toBe('Ablage')
  await v.selectText('document')
  await v.page.locator('button[data-built-host][aria-label^="Fett"]').click()
  await v.page.locator('button[data-built-host][aria-label^="Fett"][data-active]').waitFor()
  await v.page.evaluate(() => (window as any).__builtUpdate({ locale: 'vi-VN' }))
  expect(await v.page.locator('.de-file').innerText()).toBe('Ablage')
  await v.page.locator('.de-bar button.de-mode', { hasText: 'Speichern' }).click()
  const saved = await v.waitEvent('save')
  const p = (await docxParas(Buffer.from(saved.b64 as string, 'base64'))).find((x) => paraText(x).startsWith('This document'))!
  expect(runsOf(p).find((r) => r.text.includes('document'))!.s.b).toBe(true)
  expect(await v.events('load')).toHaveLength(1)
  expect(v.errors).toEqual([])
})

test('built entries can be imported where there is no DOM (a server rendering the page): nothing throws, no element is registered', async () => {
  // (this process has no DOM: what a server importing the package sees)
  expect(typeof HTMLElement).toBe('undefined')
  for (const entry of ['index.js', 'core.js', 'collab.js', 'locales/de.js', 'embed.js']) {
    const m = await import(`${packageRoot}/dist/${entry}`)
    expect([entry, Object.keys(m).length > 0]).toEqual([entry, true])
  }
  const embed = await import(`${packageRoot}/dist/embed.js`)
  expect(typeof embed.OfficeViewerElement).toBe('function')
  expect(() => embed.defineElement()).not.toThrow()
})
