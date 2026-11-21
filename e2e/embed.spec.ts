import { describe, expect, test } from 'bun:test'
import { parasToDocx } from '../src/core/convert'
import { testFont } from '../test/font-builder'
import { useBrowser, viewer } from './harness'

useBrowser()

describe('embedding without React', () => {
  test('<office-viewer>: shadow DOM, attributes, load event', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    const wc = v.page.locator('#wc')
    await wc.locator('.de-page-inner').first().waitFor()
    // rendered inside the element's shadow root, styles included there
    expect(await v.page.evaluate(() => !!document.getElementById('wc')!.shadowRoot?.querySelector('style#office-engine-styles'))).toBe(true)
    expect(await v.page.evaluate(() => (window as any).__wcEvents)).toEqual([{ kind: 'docx', name: 'sample.docx' }])
    // save="false" file-menu="false"
    expect(await wc.locator('.de-file').count()).toBe(0)
    await wc.locator('.de-bar button.de-mode', { hasText: 'Edit' }).click()
    expect(await wc.locator('.de-bar button.de-mode', { hasText: /^(Save|Download)$/ }).count()).toBe(0)
    // attributes are live
    await v.page.evaluate(() => document.getElementById('wc')!.setAttribute('toolbar', 'false'))
    await wc.locator('.de-bar').waitFor({ state: 'detached' })
    await v.page.evaluate(() => document.getElementById('wc')!.setAttribute('src', 'samples/sample.pptx'))
    await wc.locator('.de-slide-inner').first().waitFor()
    await v.page.evaluate(() => document.getElementById('wc')!.setAttribute('theme', 'dark'))
    await wc.locator('.de-root.de-dark').waitFor()
    // N3-010: one `load` per document (an attribute change is not a load)
    await v.page.waitForTimeout(300)
    expect(await v.page.evaluate(() => (window as any).__wcEvents)).toEqual([
      { kind: 'docx', name: 'sample.docx' },
      { kind: 'pptx', name: 'sample.pptx' },
    ])
    expect(v.errors).toEqual([])
  })

  test('<office-viewer>: file property', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    await v.page.evaluate(async () => {
      const wc = document.getElementById('wc') as HTMLElement & { file?: Blob }
      wc.file = await (await fetch('samples/sample.xlsx')).blob()
    })
    await v.page.locator('#wc .de-grid').waitFor()
  })

  test('<office-viewer>: fonts attribute (URLs), fonts property (bytes too)', async () => {
    const font = Buffer.from(testFont({ family: 'Wide Test' }))
    const docx = Buffer.from(await (await parasToDocx([{ runs: [{ text: 'Wide', set: { font: 'Wide Test' } }] }])).arrayBuffer())
    const asked: string[] = []
    const v = await viewer(
      {},
      {
        page: 'embed.html',
        setup: (v) =>
          v.ctx.route('**/gen/**', (r) => {
            const name = new URL(r.request().url()).pathname.split('/').pop()!
            asked.push(name)
            return r.fulfill({ body: name.endsWith('.docx') ? docx : font })
          }),
      },
    )
    const wc = v.page.locator('#wc')
    await wc.locator('.de-page-inner').first().waitFor()
    const faces = () => v.page.evaluate(() => [...document.fonts].map((f) => `${f.family.replace(/"/g, '')} ${f.status}`))
    await v.page.evaluate(() => {
      const el = document.getElementById('wc')!
      el.setAttribute('fonts', JSON.stringify([{ family: 'Wide Test', src: '/gen/wide.ttf' }]))
      el.setAttribute('src', 'gen/wide.docx')
    })
    await v.page.waitForFunction(() => [...document.fonts].some((f) => f.status === 'loaded'))
    expect(asked).toEqual(['wide.docx', 'wide.ttf'])
    expect(await faces()).toEqual(['Wide Test loaded'])
    // the property wins over the attribute: the bytes, and the file at the URL is let go
    await v.page.evaluate((bytes) => {
      const el = document.getElementById('wc') as HTMLElement & { fonts?: unknown }
      el.fonts = [{ family: 'Wide Test', src: new Uint8Array(bytes).buffer, weight: 700 }]
    }, [...font])
    await v.page.waitForFunction(() => [...document.fonts].length === 1 && [...document.fonts][0].weight === '700')
    expect(await faces()).toEqual(['Wide Test loaded'])
    expect(asked).toEqual(['wide.docx', 'wide.ttf'])
    expect(v.errors).toEqual([])
  })

  test('mount(): options, update(), destroy()', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    const host = v.page.locator('#mounted')
    await host.locator('.de-grid').waitFor()
    expect(await host.locator('.de-bar').count()).toBe(0)
    await v.page.evaluate(() => (window as any).__mounted.update({ toolbar: true }))
    await host.locator('.de-bar').waitFor()
    expect(await host.locator('.de-title').innerText()).toBe('sample.xlsx')
    await v.page.evaluate(() => (window as any).__mounted.destroy())
    expect(await host.innerHTML()).toBe('')
  })

  test('<office-viewer locale>: numbers typed are read in it; the attribute is live (#17)', async () => {
    const v = await viewer({}, { page: 'embed.html' })
    const wc = v.page.locator('#wc')
    await wc.locator('.de-page-inner').first().waitFor()
    await v.page.evaluate(() => {
      const el = document.getElementById('wc')!
      el.setAttribute('locale', 'de-DE')
      el.setAttribute('src', 'samples/sample.xlsx')
    })
    await wc.locator('.de-grid').waitFor()
    await wc.locator('.de-bar button.de-mode', { hasText: 'Edit' }).click()
    const size = wc.locator('input[aria-label="Font size"]')
    const typeSize = async (text: string) => {
      await size.click()
      await size.fill(text)
      await size.press('Enter')
      await v.page.waitForTimeout(300)
    }
    await typeSize('12,5')
    expect(await size.inputValue()).toBe('12,5')
    await v.page.evaluate(() => document.getElementById('wc')!.removeAttribute('locale'))
    await v.page.waitForFunction(() => (document.getElementById('wc')!.shadowRoot!.querySelector('input[aria-label="Font size"]') as HTMLInputElement).value === '12.5')
    await typeSize('1,234')
    expect(await size.inputValue()).toBe('12.5')
    await typeSize('10,5')
    expect(await size.inputValue()).toBe('10.5')
    expect(v.errors).toEqual([])
  })
})
