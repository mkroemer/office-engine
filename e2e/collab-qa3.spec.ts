/** QA round 3, collaboration in the browser: text being typed in a slide while a collaborator changes the slide list */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { SAMPLES, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

const PORT = 19_000 + Math.floor(Math.random() * 1000)
const HTTP = `http://localhost:${PORT}`
let server: ReturnType<typeof Bun.spawn> | null = null

beforeAll(async () => {
  server = Bun.spawn(['bun', 'scripts/collab-server.ts'], { env: { ...process.env, PORT: String(PORT) }, stdout: 'ignore', stderr: 'inherit' })
  for (let i = 0; i < 100; i++) {
    try {
      if ((await fetch(HTTP)).ok) return
    } catch {
      // not listening yet
    }
    await Bun.sleep(50)
  }
  throw new Error('collab server did not start')
})
afterAll(() => server?.kill())

async function room(sample: string) {
  const id = 'qa3-' + Math.random().toString(36).slice(2, 8)
  await fetch(`${HTTP}/rooms/${id}/file`, { method: 'PUT', body: readFileSync(`${SAMPLES}/${sample}`), headers: { 'X-File-Name': sample } })
  return id
}
const join = (id: string, user: string) =>
  viewer({ src: `${HTTP}/rooms/${id}/file?seq=0`, name: 'shared.pptx', collab: `ws://localhost:${PORT}/collab`, room: id, user, editable: true })
const count = (v: Viewer, n: number) => v.page.waitForFunction((k) => new RegExp(`of ${k}$`).test(document.querySelector('.de-bar .de-status')?.textContent ?? ''), n, { timeout: 20_000 })
const slideText = (v: Viewer, n: number) => v.page.evaluate((k) => document.querySelector(`.de-slide-frame[data-slide="${k}"]`)?.textContent ?? '', n)

describe('QA 3: a presentation shared in a room', () => {
  test('a slide added by a collaborator does not end the text being typed: it goes on in its slide, at its new place', async () => {
    const id = await room('sample.pptx')
    const a = await join(id, 'Ann')
    const b = await join(id, 'Bob')
    for (const v of [a, b]) await v.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    // Ann types in the third slide
    await a.page.locator('.de-thumb').nth(2).click()
    const third = a.page.locator('.de-slide-frame[data-slide="3"]')
    await third.waitFor()
    await third.locator('[data-sid]').first().click()
    await a.page.locator('.de-txt-edit').waitFor()
    await a.page.keyboard.press('ControlOrMeta+a')
    await a.page.keyboard.type('Ann one')
    await b.page.waitForFunction(() => document.body.innerText.includes('Ann one'), null, { timeout: 20_000 }).catch(() => undefined)
    // Bob adds a slide after the first: Ann's slide is the fourth now
    await b.pick('Slide', 'New slide')
    await count(b, 6)
    await count(a, 6)
    await a.page.locator('.de-slide-frame[data-slide="4"] .de-txt-edit').waitFor({ timeout: 20_000 })
    // she types on, without a click
    await a.page.waitForTimeout(300)
    await a.page.keyboard.type(' two')
    await a.page.waitForFunction(() => (document.querySelector('.de-slide-frame[data-slide="4"]')?.textContent ?? '').includes('Ann one two'), null, { timeout: 20_000 })
    await a.page.keyboard.press('Escape')
    // Bob sees her text in the fourth slide; nothing of it went to another slide
    await b.page.locator('.de-thumb').nth(3).click()
    await b.page.waitForFunction(() => (document.querySelector('.de-slide-frame[data-slide="4"]')?.textContent ?? '').includes('Ann one two'), null, { timeout: 20_000 })
    expect(await slideText(b, 3)).not.toContain('two')
    expect(a.errors).toEqual([])
    expect(b.errors).toEqual([])
  })
})
