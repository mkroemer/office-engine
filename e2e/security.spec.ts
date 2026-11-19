/**
 * Crafted files in the real viewer, with no Content Security Policy to hide a mistake behind: opening one makes the
 * browser request nothing, and a link that would run script is not opened.
 */
import { describe, expect, test } from 'bun:test'
import { PROBE_HOST, probeDocx, probePptx, probeVisio, probeXlsx, scriptLinkDocx } from '../test/probes'
import { plainToOffice } from '../src/core/plain'
import { READY, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

/** requests leaving the page's own origin from now on (the probe host is answered here, never reached) */
async function outbound(v: Viewer): Promise<string[]> {
  const origin = new URL(v.page.url()).origin
  const urls: string[] = []
  await v.ctx.route(`https://${PROBE_HOST}/**`, (r) => r.fulfill({ status: 204 }))
  v.page.on('request', (r) => {
    const u = r.url()
    if (!u.startsWith(origin) && !/^(blob|data):/.test(u)) urls.push(u)
  })
  return urls
}

async function openBytes(v: Viewer, name: string, file: Blob) {
  await v.chooseFile(() => v.menu('Open…'), { name, mimeType: 'application/octet-stream', buffer: Buffer.from(await file.arrayBuffer()) })
  await v.page.waitForFunction((n) => document.querySelector('.de-bar .de-title')?.textContent === n, name)
}

/**
 * Every url() an element's style would load, other than the document's own pictures (blob: / data:) and references
 * inside the page (#id): computed values of the properties that take an image or a reference, and inline styles.
 */
const loadedUrls = (v: Viewer) =>
  v.page.evaluate(() => {
    const props = ['backgroundImage', 'borderImageSource', 'listStyleImage', 'maskImage', 'content', 'cursor', 'filter', 'clipPath', 'fill', 'stroke', 'markerStart', 'markerEnd'] as const
    const here = location.href.split('#')[0]
    const out = new Set<string>()
    const scan = (text: string, where: string) => {
      for (const m of text.matchAll(/url\(\s*["']?([^"')]*)/gi)) {
        const u = m[1].trim()
        if (/^(blob|data):/i.test(u) || u.startsWith('#') || u.startsWith(here + '#')) continue
        out.add(`${where}: ${u}`)
      }
    }
    for (const el of document.querySelectorAll<HTMLElement>('*')) {
      const cs = getComputedStyle(el)
      for (const p of props) scan(String(cs[p] ?? ''), p)
      scan(el.getAttribute('style') ?? '', 'style')
      for (const a of ['fill', 'stroke']) scan(el.getAttribute(a) ?? '', a)
    }
    return [...out]
  })

describe('a crafted file', () => {
  const cases = [
    ['pptx', probePptx, 'samples/sample.docx'],
    ['xlsx', probeXlsx, 'samples/sample.docx'],
    ['docx', probeDocx, 'samples/sample.xlsx'],
    // a Visio drawing is shown as slides
    ['vsdx', probeVisio, 'samples/sample.docx'],
  ] as const

  for (const [kind, build, start] of cases) {
    test(`${kind}: colors that are also a url() make the browser request nothing`, async () => {
      const v = await viewer({ src: start })
      const requests = await outbound(v)
      await openBytes(v, `probe.${kind}`, await build())
      await v.page.locator(READY[kind === 'vsdx' ? 'pptx' : kind]).first().waitFor()
      // the probes are shown: text of the crafted shapes / the text box, the sheet's cells
      if (kind === 'pptx' || kind === 'vsdx') await v.page.getByText('probe 1').first().waitFor()
      if (kind === 'docx') await v.page.getByText('probe box').first().waitFor()
      if (kind === 'xlsx') await v.page.locator('.de-cell').first().waitFor()
      await v.page.waitForTimeout(1200)
      expect(await loadedUrls(v)).toEqual([])

      // editing draws the same content again (and the toolbar shows the selection's colors); a drawing is view only
      if (kind !== 'vsdx') await v.edit()
      // a cell with the crafted fill, font color and border: the toolbar shows them
      if (kind === 'xlsx') await v.clickCell('A3')
      if (kind === 'pptx') await v.page.getByText('probe 1').first().click()
      await v.page.waitForTimeout(800)
      expect(await loadedUrls(v)).toEqual([])
      expect(requests).toEqual([])
      expect(v.errors).toEqual([])
    })
  }

  test('docx: "Open link" does not open a link that would run script', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true })
    const item = v.page.locator('.de-ctx button', { hasText: 'Open link' })
    const menuOnLink = async () => {
      await v.selectText('example.com', 'end')
      const box = (await v.page.getByText('example.com').first().boundingBox())!
      await v.page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' })
      await item.waitFor()
    }
    // a web link: offered
    await menuOnLink()
    expect(await item.isDisabled()).toBe(false)
    await v.page.keyboard.press('Escape')

    await openBytes(v, 'link.docx', await scriptLinkDocx())
    await v.page.getByText('example.com').first().waitFor()
    await v.page.waitForTimeout(500)
    await v.edit()
    await menuOnLink()
    expect(await item.isDisabled()).toBe(true)
    const pages = v.ctx.pages().length
    await item.click({ force: true }).catch(() => {})
    await v.page.waitForTimeout(300)
    expect(v.ctx.pages().length).toBe(pages)
    expect(await v.page.evaluate(() => (window as unknown as { __ran?: number }).__ran)).toBeUndefined()
  })

  // "Open image in new tab" navigates to the <img>'s URL: for an SVG that must not be a same-origin document
  const EVIL_SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" onload="document.title=\'ran:\'+(function(){try{return localStorage.getItem(\'secret\')}catch(e){return \'blocked\'}})()"><script>document.title="ran:"+(function(){try{return localStorage.getItem("secret")}catch(e){return "blocked"}})()</script><rect width="120" height="80" fill="red"/></svg>'

  async function opensAsOpaque(v: Viewer, imgSelector: string) {
    const img = v.page.locator(imgSelector).first()
    await img.waitFor()
    const src = (await img.getAttribute('src'))!
    expect(src.startsWith('blob:')).toBe(false)
    // what the browser does for the context menu entry: a new tab on the image's address
    const origin = new URL(v.page.url()).origin
    const tab = await v.ctx.newPage()
    await tab.goto(origin + '/')
    await tab.evaluate(() => localStorage.setItem('secret', 'host-data'))
    await tab.goto(src).catch(() => {})
    await tab.waitForFunction(() => document.title.startsWith('ran:'), null, { timeout: 5000 }).catch(() => {})
    const title = await tab.title()
    // the script may run, but not as the host: it sees nothing of the host's storage
    expect(title).not.toContain('host-data')
    expect(await tab.evaluate(() => location.origin)).not.toBe(origin)
  }

  test('an SVG file: its picture is not a same-origin document', async () => {
    const v = await viewer({ src: 'samples/sample.docx' })
    await openBytes(v, 'evil.svg', new Blob([EVIL_SVG], { type: 'image/svg+xml' }))
    await opensAsOpaque(v, '[data-testid=de-picture] img')
  })

  test('an SVG picture in a docx: not a same-origin document', async () => {
    const v = await viewer({ src: 'samples/sample.xlsx' })
    await openBytes(v, 'evil.docx', await plainToOffice(EVIL_SVG, 'svg', 'evil.svg'))
    await opensAsOpaque(v, '.de-page-inner img')
  })
})
