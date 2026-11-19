import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { rebuildZip } from '../src/core/zipwrite'
import { SAMPLES, useBrowser, viewer, fromB64, xlsxSheet, docxParas, paraText, runsOf } from './harness'

useBrowser()

const PORT = 18_000 + Math.floor(Math.random() * 1000)
const HTTP = `http://localhost:${PORT}`
let server: ReturnType<typeof Bun.spawn> | null = null

beforeAll(async () => {
  // frequent snapshots: live editing must not notice them, newcomers join from them
  server = Bun.spawn(['bun', 'scripts/collab-server.ts'], { env: { ...process.env, PORT: String(PORT), SNAPSHOT_EVERY: '2' }, stdout: 'ignore', stderr: 'inherit' })
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
  const id = 'e2e-' + Math.random().toString(36).slice(2, 8)
  await fetch(`${HTTP}/rooms/${id}/file`, { method: 'PUT', body: readFileSync(`${SAMPLES}/${sample}`), headers: { 'X-File-Name': sample } })
  return id
}

const join = (id: string, user: string, base = 0) =>
  viewer({ src: `${HTTP}/rooms/${id}/file?seq=${base}`, name: 'shared.xlsx', collab: `ws://localhost:${PORT}/collab`, room: id, user, base, editable: true })

describe('real-time collaboration', () => {
  test('edits reach the other editor; presence is shown', async () => {
    const id = await room('sample.xlsx')
    const a = await join(id, 'Ann')
    const b = await join(id, 'Bob')
    for (const v of [a, b]) await v.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await a.page.waitForFunction(() => /2 editing/.test(document.querySelector('.de-collab')?.textContent ?? ''), null, { timeout: 20_000 })
    expect(await a.page.locator('.de-peer').first().getAttribute('title')).toBe('Bob')

    await a.typeCell('J10', 'from Ann')
    await b.page.waitForFunction(() => document.body.innerText.includes('from Ann'), null, { timeout: 20_000 })
    expect(await b.shown('J10')).toBe('from Ann')
    await b.typeCell('J11', '=LEN(J10)')
    await a.page.waitForFunction(() => [...document.querySelectorAll('.de-cell')].some((c) => (c as HTMLElement).innerText === '8'), null, { timeout: 20_000 })
  })

  test('controller history follows collaboration across rename and save/reopen', async () => {
    const id = await room('sample.xlsx')
    const v = await viewer({ src: `${HTTP}/rooms/${id}/file?seq=0`, name: 'shared.xlsx', collab: `ws://localhost:${PORT}/collab`, room: id, user: 'Ann', editable: true, controller: true, onSave: true })
    await v.page.locator('.de-collab-live').waitFor()
    await v.typeCell('J12', 'before rename')
    await v.page.waitForFunction(() => window.__ctl.snapshot().canUndo)
    await v.page.locator('.de-tab[role=tab]').first().dblclick()
    await v.page.locator('.de-tab-input').fill('Renamed')
    await v.page.keyboard.press('Enter')
    await v.page.waitForFunction(() => window.__ctl.snapshot().outline?.items[0].title === 'Renamed')
    expect(await v.page.evaluate(() => window.__ctl.snapshot().canUndo)).toBe(false)
    // Verify Ctrl+Z and the controller agree on the actual document, rather than just flags.
    await v.page.locator('.de-grid').focus()
    await v.page.keyboard.press('Control+z')
    await v.page.evaluate(() => window.__ctl.save())
    let saved = await v.waitEvent('save')
    expect((await xlsxSheet(fromB64(saved.b64 as string))).cell('J12')).toBe('before rename')
    await v.typeCell('J12', 'after rename')
    await v.page.waitForFunction(() => window.__ctl.snapshot().canUndo)
    await v.page.evaluate(() => window.__ctl.undo())
    await v.page.waitForFunction(() => window.__ctl.snapshot().canRedo && !window.__ctl.snapshot().canUndo)
    await v.page.evaluate(() => window.__ctl.save())
    saved = await v.waitEvent('save', 2)
    expect((await xlsxSheet(fromB64(saved.b64 as string))).cell('J12')).toBe('before rename')
    await v.page.evaluate(() => window.__ctl.redo())
    await v.page.waitForFunction(() => window.__ctl.snapshot().canUndo && !window.__ctl.snapshot().canRedo)
    await v.page.evaluate(() => window.__ctl.save())
    saved = await v.waitEvent('save', 3)
    const reopened = await xlsxSheet(fromB64(saved.b64 as string))
    expect(reopened.names[0]).toBe('Renamed')
    expect(reopened.cell('J12')).toBe('after rename')
    expect(v.errors).toEqual([])
  })

  test('a newcomer opens the latest snapshot and only replays later edits', async () => {
    const id = await room('sample.xlsx')
    const a = await join(id, 'Ann')
    await a.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    for (const [ref, text] of [['J10', 'one'], ['J11', 'two'], ['J12', 'three']]) await a.typeCell(ref, text)
    let snapshot = 0
    for (let i = 0; i < 100 && snapshot < 2; i++) {
      snapshot = ((await (await fetch(`${HTTP}/rooms/${id}`)).json()) as { snapshot: number }).snapshot
      await Bun.sleep(100)
    }
    expect(snapshot).toBeGreaterThanOrEqual(2)

    const b = await join(id, 'Bob', snapshot)
    await b.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await b.page.waitForFunction(() => document.body.innerText.includes('three'), null, { timeout: 20_000 })
    expect(await b.shown('J10')).toBe('one')
    await a.typeCell('J13', 'after')
    await b.page.waitForFunction(() => document.body.innerText.includes('after'), null, { timeout: 20_000 })
    await b.typeCell('J14', 'from Bob')
    await a.page.waitForFunction(() => document.body.innerText.includes('from Bob'), null, { timeout: 20_000 })
    expect(b.errors).toEqual([])
  })

  test('New / Open are not offered in a shared room', async () => {
    const id = await room('sample.xlsx')
    const a = await join(id, 'Ann')
    await a.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    const labels = await a.menuLabels()
    expect(labels).not.toContain('Open…')
    expect(labels).not.toContain('Document')
    expect(labels).toContain('Excel workbook')
  })
})

describe('docx: a built-in style the document lacks', () => {
  /** a document without a styles part: every built-in style is added on first use */
  async function blankRoom() {
    const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    const blank = await rebuildZip(
      null,
      new Map([
        ['[Content_Types].xml', `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`],
        ['_rels/.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`],
        ['word/document.xml', `${head}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`],
      ]),
    )
    const id = 'e2e-' + Math.random().toString(36).slice(2, 8)
    await fetch(`${HTTP}/rooms/${id}/file`, { method: 'PUT', body: await blank.arrayBuffer(), headers: { 'X-File-Name': 'blank.docx' } })
    return id
  }
  type V = Awaited<ReturnType<typeof viewer>>
  const text = async (v: V) => (await v.pagesText()).replace(/\s+/g, ' ').trim()
  const colour = (v: V, s: string) =>
    v.page.evaluate((x) => {
      const el = [...document.querySelectorAll('.de-pages *')].find((e) => e.childElementCount === 0 && (e.textContent ?? '').includes(x))
      return el ? getComputedStyle(el).color : 'not found'
    }, s)
  /** Heading 1, then typing at once: the style is being added to the file (in a room: ordered by the hub first) */
  async function headingThenType(v: V) {
    await v.page.locator('.de-page-inner').first().click()
    await v.page.keyboard.type('Hello', { delay: 15 })
    await v.page.waitForTimeout(300)
    await v.pick('Paragraph style', 'Heading 1')
    await v.page.keyboard.type(' world')
    await v.page.waitForTimeout(800)
    await v.page.keyboard.type('!')
    await v.page.waitForTimeout(500)
  }
  const HEADING = 'rgb(47, 84, 150)'

  test('in a room: what is typed right after picking it reaches everyone, the caret stays', async () => {
    const id = await blankRoom()
    const join = (user: string) => viewer({ src: `${HTTP}/rooms/${id}/file?seq=0`, name: 'shared.docx', collab: `ws://localhost:${PORT}/collab`, room: id, user, base: 0, editable: true })
    const a = await join('Ann')
    const b = await join('Bob')
    for (const v of [a, b]) await v.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await headingThenType(a)
    // replays the operation that adds the style, then the edits
    const c = await join('Cid')
    await c.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await c.page.waitForFunction(() => document.querySelector('.de-pages')?.textContent?.includes('!'), null, { timeout: 20_000 })
    await c.page.waitForTimeout(800)
    for (const v of [a, b, c]) {
      expect(await text(v)).toBe('Hello world!')
      expect(await colour(v, 'Hello')).toBe(HEADING)
      expect(v.errors).toEqual([])
    }
  })

  test('alone: the caret stays where the typing took it while the style was added', async () => {
    const id = await blankRoom()
    const v = await viewer({ src: `${HTTP}/rooms/${id}/file?seq=0`, name: 'blank.docx', editable: true })
    await v.page.locator('.de-ebar').waitFor()
    await headingThenType(v)
    expect(await text(v)).toBe('Hello world!')
    expect(await colour(v, 'Hello')).toBe(HEADING)
    expect(v.errors).toEqual([])
  })
})


for (const control of ['input', 'button'] as const) {
  test(`docx: focused host ${control} keeps its formatting target across peer structure and text edits`, async () => {
    const id = await room('sample.docx')
    const joinDoc = (user: string) => viewer({ src: `${HTTP}/rooms/${id}/file?seq=0`, name: 'shared.docx', collab: `ws://localhost:${PORT}/collab`, room: id, user, base: 0, editable: true, controller: true, onSave: true })
    const a = await joinDoc('Ann')
    const b = await joinDoc('Bob')
    for (const v of [a, b]) await v.page.locator('.de-collab-live').waitFor({ timeout: 20_000 })
    await b.page.evaluate(() => {
      const input = document.createElement('input')
      input.id = 'collab-format-input'; input.style.cssText = 'position:fixed;top:30px;left:0;z-index:999999'
      document.body.append(input)
      for (const [i, command] of [
        { op: 'toggle', mark: 'underline' }, { op: 'size', pt: 31 }, { op: 'color', color: '#00aa00' },
      ].entries()) {
        const button = document.createElement('button')
        button.id = `collab-format-${i}`; button.textContent = String(i)
        button.style.cssText = `position:fixed;top:0;left:${i * 60}px;z-index:999999`
        button.onclick = () => window.__ctl.format(command as any)
        document.body.append(button)
      }
    })
    await b.selectText('document')
    const commands = [{ op: 'toggle', mark: 'underline' }, { op: 'size', pt: 31 }, { op: 'color', color: '#00aa00' }]
    const apply = async (i: number) => {
      if (control === 'input') {
        await b.page.locator('#collab-format-input').click()
        await b.page.evaluate((command) => window.__ctl.format(command as any), commands[i])
      } else if (i === 2) {
        await b.page.locator(`#collab-format-${i}`).focus()
        await b.page.keyboard.press('Enter')
      } else await b.page.locator(`#collab-format-${i}`).click()
      await b.page.waitForTimeout(200)
    }
    await apply(0)
    await a.selectText('Quarterly Report', 'start')
    await a.page.keyboard.type('Peer prefix')
    await a.page.keyboard.press('Enter')
    await b.page.waitForFunction(() => document.querySelector('.de-pages')?.textContent?.includes('Peer prefix'))
    await b.page.waitForTimeout(300)
    expect(await b.page.evaluate(() => window.__ctl.snapshot().format?.underline)).toBe(true)
    await apply(1)
    await a.selectText('This', 'start')
    await a.page.keyboard.type('Zq')
    await b.page.waitForFunction(() => document.querySelector('.de-pages')?.textContent?.includes('ZqThis document'))
    await b.page.waitForTimeout(300)
    await apply(2)
    expect(await b.page.evaluate(() => window.__ctl.snapshot().format)).toMatchObject({ underline: true, size: 31, color: '#00aa00' })
    await a.page.waitForTimeout(500)
    for (const v of [a, b]) {
      await v.page.evaluate(() => window.__ctl.save())
      const bytes = fromB64((await v.waitEvent('save')).b64 as string)
      const p = (await docxParas(bytes)).find((p) => paraText(p).startsWith('ZqThis document'))!
      const run = runsOf(p).find((r) => r.text === 'document')!
      expect(run.s.u).toBe('single')
      expect(run.s.sz).toBe(31)
      expect(run.s.color).toMatch(/00aa00/i)
      expect(v.errors).toEqual([])
    }
  })
}
