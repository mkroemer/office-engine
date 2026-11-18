/** Production-artifact checks: run `bun run build` first, as CI does before browser tests. */
import { beforeAll, describe, expect, test } from 'bun:test'
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs'
import { parasToDocx, rowsToXlsx } from '../src/core/convert'
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate'
import { useBrowser, viewer, type Viewer, SAMPLES } from './harness'

useBrowser()
const packageRoot = process.env.BUILT_PACKAGE_ROOT ?? process.cwd()
const pkg = `/@fs${packageRoot}/dist/index.js`
const german = `/@fs${packageRoot}/dist/locales/de.js`

/** Load the emitted package in a real browser, sharing the fixture's React runtime. */
async function mount(v: Viewer, props: Record<string, unknown> = {}, outerGerman = false) {
  await v.page.waitForFunction(() => !!(window as any).__builtRuntime)
  await v.page.evaluate(async ({ pkg, german, props, outerGerman }) => {
    const { React, createRoot } = (window as any).__builtRuntime
    const api = await import(/* @vite-ignore */ pkg)
    const de = outerGerman ? (await import(/* @vite-ignore */ german)).de : undefined
    const root = createRoot(document.getElementById('built')!)
    let current = props
    const onLoad = (c: any) => { (window as any).__builtClient = c; window.__events.push({ type: 'load', kind: c.kind, name: c.name }) }
    const onError = (e: Error) => window.__events.push({ type: 'error', message: e.message })
    const render = () => {
      const view = React.createElement(api.DocViewer, {
        ...current,
        onLoad,
        onError,
      })
      root.render(outerGerman ? React.createElement(api.StringsProvider, { strings: de }, view) : view)
    }
    ;(window as any).__builtUpdate = (next: any) => { current = { ...current, ...next }; render() }
    ;(window as any).__builtApi = api
    render()
  }, { pkg, german, props, outerGerman })
  if (props.src) await v.ready()
  else await v.page.locator('.de-file').waitFor()
}

async function open(props: Record<string, unknown>, locale = 'en-GB', outerGerman = false) {
  const v = await viewer({}, { page: 'built-locale.html', wait: false, locale })
  await mount(v, props, outerGerman)
  return v
}

async function update(v: Viewer, props: Record<string, unknown>) {
  await v.page.evaluate((props) => (window as any).__builtUpdate(props), props)
}

function emfDeck() {
  const files = unzipSync(readFileSync(`${SAMPLES}/sample.pptx`))
  const slide = 'ppt/slides/slide1.xml'
  files[slide] = strToU8(strFromU8(files[slide]).replace('</p:spTree>', `<p:pic><p:nvPicPr><p:cNvPr id="99" name="EMF regression"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId99"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree>`))
  const rel = 'ppt/slides/_rels/slide1.xml.rels'
  files[rel] = strToU8(strFromU8(files[rel]).replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/regression.emf"/></Relationships>'))
  files['ppt/media/regression.emf'] = new Uint8Array([1, 0, 0, 0])
  files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Default Extension="emf" ContentType="image/x-emf"/></Types>'))
  return Buffer.from(zipSync(files))
}

// Refuse missing or stale local output rather than silently skipping or testing an old bundle.
function newestSource(path: string): number {
  return Math.max(...readdirSync(path, { withFileTypes: true }).map((entry) => {
    const file = `${path}/${entry.name}`
    return entry.isDirectory() ? newestSource(file) : statSync(file).mtimeMs
  }))
}
beforeAll(() => {
  expect(existsSync(`${packageRoot}/dist/index.js`), 'Run bun run build before browser tests').toBe(true)
  if (!process.env.BUILT_PACKAGE_ROOT) {
    expect(statSync(`${packageRoot}/dist/index.js`).mtimeMs, 'dist is stale; run bun run build').toBeGreaterThanOrEqual(newestSource(`${packageRoot}/src`))
  }
})

describe('built package localization', () => {
  for (const locale of ['en-GB', 'vi-VN', 'de-DE', 'ar-EG']) {
    test(`default formatting follows the ${locale} browser; invalid tags fall back`, async () => {
      const v = await open({ src: 'samples/sample.docx' }, locale)
      const read = () => v.page.evaluate(() => {
        const api = (window as any).__builtApi
        const s = api.resolveStrings()
        const date = new Date('2026-03-05T14:00:00Z')
        return { locale: s.locale, date: date.toLocaleString(s.locale), browserDate: date.toLocaleString(), count: s.docx.status(1, 2, 1234567, s.locale), browserCount: (1234567).toLocaleString() }
      })
      const initial = await read()
      const optionalGermanRequests = await v.page.evaluate(() => performance.getEntriesByType('resource').filter((e) => e.name.includes('/locales/de.js')).length)
      expect(optionalGermanRequests).toBe(0)
      const created = await v.page.evaluate(async () => {
        const xml = await (await (window as any).__builtClient.call('part', 'docProps/core.xml')).text()
        const value = new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 'created')[0]!.textContent!
        return new Date(value).toLocaleString()
      })
      await v.menu('Document details')
      await v.page.locator('[role=dialog] dd', { hasText: created }).first().waitFor()
      expect(await v.page.locator('[role=dialog]').innerText()).toContain('8.5 × 11 in')
      expect(await v.status()).toMatch(/^Page 1 of [0-9]+/)
      await v.page.locator('[role=dialog] button', { hasText: 'Close' }).click()
      expect(initial.locale).toBeUndefined()
      expect(initial.date).toBe(initial.browserDate)
      expect(initial.count).toContain(initial.browserCount)
      expect(await v.page.locator('.de-file').innerText()).toBe('File')
      for (const malformed of [{ locale: 'de_DE' }, { locale: '__proto__' }, { locale: undefined, strings: { locale: 'de_DE' } }]) {
        await update(v, malformed)
        await v.page.locator('.de-page-inner').first().waitFor()
        expect(await v.page.locator('.de-file').innerText()).toBe('File')
        await v.menu('Document details')
        await v.page.locator('[role=dialog] dd', { hasText: created }).first().waitFor()
        await v.page.locator('[role=dialog] button', { hasText: 'Close' }).click()
        expect(v.errors).toEqual([])
      }
      // Render date/count consumers inside the actual provider, not only the resolver.
      const rendered = await v.page.evaluate(async ({ pkg }) => {
        const { React, createRoot } = (window as any).__builtRuntime
        const api = await import(/* @vite-ignore */ pkg)
        const host = document.createElement('div'); document.body.append(host)
        const root = createRoot(host)
        function Probe() {
          const s = api.useStrings()
          return React.createElement('span', null, `${new Date('2026-03-05T14:00:00Z').toLocaleString(s.locale)}|${s.docx.status(1, 2, 1234567, s.locale)}`)
        }
        root.render(React.createElement(api.StringsProvider, { locale: 'de_DE' }, React.createElement(Probe)))
        await new Promise((r) => setTimeout(r, 100))
        const text = host.textContent; root.unmount(); host.remove()
        return text
      }, { pkg })
      expect(rendered).toContain(initial.browserDate)
      expect(rendered).toContain(initial.browserCount)
    })
  }


  test('standalone production embed loads English and explicitly imported German', async () => {
    const v = await viewer({}, { page: 'built-locale.html', wait: false, locale: 'vi-VN' })
    await v.page.addScriptTag({ url: `/@fs${packageRoot}/dist/office-engine.embed.js` })
    await v.page.evaluate(() => {
      const el = document.createElement('office-viewer')
      el.setAttribute('src', '/samples/sample.docx')
      el.setAttribute('locale', 'de_DE')
      el.addEventListener('load', (e: any) => window.__events.push({ type: 'load', ...e.detail }))
      el.addEventListener('error', (e: any) => window.__events.push({ type: 'error', message: e.detail?.message }))
      document.getElementById('built')!.append(el)
    })
    await v.ready('docx')
    expect(await v.page.locator('office-viewer .de-file').innerText()).toBe('File')
    await v.page.evaluate(async ({ german }) => {
      const el = document.querySelector('office-viewer')! as any
      el.strings = (await import(/* @vite-ignore */ german)).de
      el.setAttribute('locale', 'en-GB')
    }, { german })
    await v.page.locator('office-viewer .de-file', { hasText: 'Datei' }).waitFor()
    await v.page.locator('office-viewer .de-file').click()
    await v.page.locator('office-viewer .de-dd-text', { hasText: 'Dokumenteigenschaften' }).click()
    await v.page.locator('office-viewer [role=dialog][aria-label="Dokumenteigenschaften"]').waitFor()
    expect((await v.events('load')).length).toBe(1)
    expect(v.errors).toEqual([])
  })

  test('opt-in German, outer provider, partial overrides, pluralization and live switching', async () => {
    const v = await open({ src: 'samples/sample.xlsx', editable: true, comments: { open: true } }, 'en-GB', true)
    expect(await v.page.locator('.de-file').innerText()).toBe('Datei')
    await v.page.locator('.de-cm-empty', { hasText: 'Zelle' }).waitFor()
    await v.page.evaluate(async ({ german }) => {
      const { React, createRoot } = (window as any).__builtRuntime
      const api = (window as any).__builtApi
      const de = (await import(/* @vite-ignore */ german)).de
      const host = document.createElement('div'); host.id = 'plural-probe'; document.body.append(host)
      function Probe() {
        const s = api.useStrings()
        return React.createElement('span', null, [s.docx.status(1, 1, 1, s.locale), s.xlsx.status(1, 1, s.locale), s.xlsx.usedRangeValue(2, 3, s.locale), s.pptx.status(1, 2, s.locale)].join('|'))
      }
      createRoot(host).render(React.createElement(api.StringsProvider, { strings: de }, React.createElement(Probe)))
      return { loaded: performance.getEntriesByType('resource').map((e) => e.name).filter((x) => x.includes('/locales/de.js')) }
    }, { german })
    await v.page.locator('#plural-probe', { hasText: '1 Wort|1 Zeile × 1 Spalte|2 Zeilen × 3 Spalten|Folie 1 von 2' }).waitFor()
    expect(await v.page.evaluate(() => performance.getEntriesByType('resource').some((e) => e.name.includes('/locales/de.js')))).toBe(true)
    await update(v, { strings: { file: { menu: 'Ablage' } } })
    await v.page.locator('.de-file', { hasText: 'Ablage' }).waitFor()
    expect(await v.barButton('Ansicht').count()).toBe(1)
    // Override the inherited German table with the emitted English table and switch back without reopening.
    await v.page.evaluate(() => (window as any).__builtUpdate({ strings: (window as any).__builtApi.en }))
    await v.page.locator('.de-file', { hasText: 'File' }).waitFor()
    await v.page.locator('.de-cm-empty', { hasText: 'cell' }).waitFor()
    await update(v, { strings: undefined })
    await v.page.locator('.de-file', { hasText: 'Datei' }).waitFor()
    await v.page.locator('.de-cm-empty', { hasText: 'Zelle' }).waitFor()
    expect((await v.events('load')).length).toBe(1)
    expect(v.errors).toEqual([])
  })


  test('running DOCX/XLSX status receives numeric counts and renders singular German', async () => {
    const docx = Buffer.from(await (await parasToDocx([{ runs: [{ text: 'Wort' }] }])).arrayBuffer())
    const xlsx = Buffer.from(await (await rowsToXlsx([['Wert']], 'Blatt')).arrayBuffer())
    const statuses: Record<string, string> = {}
    for (const [kind, bytes] of [['docx', docx], ['xlsx', xlsx]] as const) {
      const v = await viewer({}, { page: 'built-locale.html', wait: false })
      await v.page.route(`**/singular.${kind}`, (route) => route.fulfill({ body: bytes }))
      await mount(v, { src: `/singular.${kind}` }, true)
      const expected = kind === 'docx' ? 'Seite 1 von 1 · 1 Wort' : '1 Zeile × 1 Spalte'
      await v.page.locator('.de-bar .de-status', { hasText: expected }).waitFor()
      statuses[kind] = await v.page.locator('.de-bar .de-status').first().innerText()
      expect(statuses[kind]).toBe(expected)
      if (kind === 'docx') {
        await v.page.evaluate(() => (window as any).__builtUpdate({ strings: { docx: { status: (p: number, ps: number, words: number) => `${typeof p}/${typeof ps}/${typeof words}:${words === 1 ? 'one' : 'many'}` } } }))
        await v.page.locator('.de-bar .de-status', { hasText: 'number/number/number:one' }).waitFor()
      }
      expect(v.errors).toEqual([])
    }
  })

  test('comment navigation survives fresh strings identities on the same sheet', async () => {
    const v = await open({ src: 'samples/sample.xlsx', editable: true, comments: { open: true } })
    await v.clickCell('B3')
    await v.page.locator('.de-cm-new').click()
    await v.page.locator('.de-cm-draft textarea').fill('Navigation regression')
    await v.page.locator('.de-cm-draft button', { hasText: 'Comment' }).click()
    const card = v.page.locator('.de-cm-card').first()
    await card.waitFor()
    await v.clickCell('A1')
    await card.click()
    await v.page.waitForFunction(() => document.querySelector('.de-fx-ref')?.textContent === 'B3')
    await v.clickCell('A1')
    await update(v, { strings: { file: { menu: 'Fresh identity' } } })
    await v.page.locator('.de-file', { hasText: 'Fresh identity' }).waitFor()
    await card.click()
    await v.page.waitForFunction(() => document.querySelector('.de-fx-ref')?.textContent === 'B3', null, { timeout: 3000 })
    expect(await v.page.locator('.de-fx-ref').innerText()).toBe('B3')
    expect(v.errors).toEqual([])
  })

  test('German clipboard failures and paste limits use the active strings', async () => {
    const bytes = Buffer.from(await (await rowsToXlsx(Array.from({ length: 6000 }, (_, i) => [String(i)]), 'Data')).arrayBuffer())
    const v = await viewer({}, { page: 'built-locale.html', wait: false, locale: 'de-DE' })
    await v.page.route('**/clipboard.xlsx', (route) => route.fulfill({ body: bytes }))
    await mount(v, { src: '/clipboard.xlsx', editable: true }, true)
    await v.clickCell('A1')
    await v.page.keyboard.press('ControlOrMeta+a')
    await v.page.evaluate(() => {
      Object.defineProperty(navigator.clipboard, 'writeText', {
        configurable: true,
        value: async () => { throw new DOMException('Clipboard denied', 'NotAllowedError') },
      })
      const data = new DataTransfer()
      document.activeElement!.dispatchEvent(new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }))
    })
    await v.page.locator('.de-toast', { hasText: 'Der Browser hat den Zugriff auf die Zwischenablage blockiert. Zum Kopieren weniger Zeilen auswählen.' }).waitFor()
    await v.clickCell('A1')
    await v.page.evaluate(() => {
      const client = (window as any).__builtClient
      const call = client.call.bind(client)
      client.call = (op: string, ...args: any[]) => {
        const result = call(op, ...args)
        if (op === 'xlsx.set') (window as any).__pasteDone = result
        return result
      }
      const data = new DataTransfer()
      data.setData('text/plain', Array.from({ length: 100001 }, (_, i) => String(i)).join('\n'))
      document.activeElement!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }))
    })
    await v.page.locator('.de-toast', { hasText: 'Nur die ersten 100.000 von 100.001 Zellen wurden eingefügt' }).waitFor()
    await v.page.evaluate(() => (window as any).__pasteDone)
    expect(v.errors).toEqual([])
  })

  test('Visio labels use the active table and main keeps the drawing read-only', async () => {
    const v = await open({ src: 'samples/sample.vsdx', editable: true }, 'en-GB', true)
    await v.menu('Dokumenteigenschaften')
    const germanLabel = await v.page.locator('[role=dialog]').innerText()
    expect(germanLabel).toContain('Microsoft-Visio-Zeichnung')
    expect(await v.barButton('Ansicht').count()).toBe(0)
    await v.page.locator('[role=dialog] button', { hasText: 'Schließen' }).click()
    await v.page.evaluate(() => (window as any).__builtUpdate({ strings: (window as any).__builtApi.en }))
    await v.page.locator('.de-file', { hasText: 'File' }).waitFor()
    await v.menu('Document details')
    const englishLabel = await v.page.locator('[role=dialog]').innerText()
    expect(englishLabel).toContain('Microsoft Visio drawing')
    expect(v.errors).toEqual([])
  })

  test('German EMF label survives HTML export/reopen and print/PDF separate roots', async () => {
    const v = await viewer({}, { page: 'built-locale.html', wait: false })
    await v.page.route('**/locale-emf.pptx', (route) => route.fulfill({ body: emfDeck(), contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }))
    // Avoid opening the native dialog, while retaining the actual production print preparation and root.
    await v.page.addInitScript(() => {
      window.print = () => {}
      Document.prototype.execCommand = () => true
    })
    await mount(v, { src: '/locale-emf.pptx' }, true)
    await v.page.locator('.de-slide-inner .de-ph', { hasText: 'EMF-Bild' }).first().waitFor()
    const html = await v.downloadAs('Webseite')
    expect(html.text.includes('EMF-Bild')).toBe(true)
    expect(html.text.includes('EMF image')).toBe(false)
    const reopened = await v.ctx.newPage()
    await reopened.setContent(html.text)
    expect(await reopened.locator('.de-ph', { hasText: 'EMF-Bild' }).count()).toBeGreaterThan(0)
    for (const action of ['Drucken', 'PDF-Dokument']) {
      await v.menu(action)
      const captured = await v.page.waitForFunction(() => {
        const frame = [...document.querySelectorAll('iframe')].find((f) => !f.hasAttribute('data-captured') && f.contentDocument?.querySelector('.de-print')?.textContent?.includes('EMF-Bild'))
        if (!frame) return false
        frame.setAttribute('data-captured', '')
        return frame.contentDocument!.documentElement.outerHTML
      })
      const prepared = await captured.jsonValue() as string
      expect(prepared).not.toContain('EMF image')
      if (action === 'PDF-Dokument') {
        await reopened.setContent(prepared)
        expect((await reopened.pdf({ preferCSSPageSize: true })).subarray(0, 4).toString()).toBe('%PDF')
      }
    }
    expect(v.errors).toEqual([])
  })
})
