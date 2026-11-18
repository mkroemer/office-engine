import { strFromU8, unzipSync } from 'fflate'
import { expect, test } from 'bun:test'
import { docxParas, fromB64, openFile, xlsxSheet, paraText, runsOf, useBrowser, viewer, type Viewer } from './harness'

useBrowser()
const state = (v: Viewer) => v.page.evaluate(() => window.__ctl.snapshot())
const settle = (v: Viewer) => v.page.waitForTimeout(150)
const formatting = (s: Awaited<ReturnType<typeof state>>) => ({ format: s.format, dirty: s.dirty, canUndo: s.canUndo, canRedo: s.canRedo })

async function selectTarget(v: Viewer, kind: 'docx' | 'pptx', index = 0) {
  if (kind === 'docx' && index === 0) return v.selectText('document')
  if (kind === 'docx') {
    await v.page.evaluate((i) => {
      const root = document.querySelectorAll<HTMLElement>('.de-pages')[i]
      root.focus()
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const start = n.textContent!.indexOf('document')
        if (start < 0) continue
        const range = document.createRange()
        range.setStart(n, start)
        range.setEnd(n, start + 8)
        getSelection()!.removeAllRanges()
        getSelection()!.addRange(range)
        break
      }
    }, index)
  } else {
    await v.page.locator('.de-main').nth(index).locator('.de-slide-frame').first().getByText('office-engine', { exact: true }).click()
    await v.page.locator('.de-txt-edit').waitFor()
    await v.page.keyboard.press('ControlOrMeta+a')
  }
  await settle(v)
}

for (const kind of ['docx', 'pptx'] as const) {
  test(`${kind}: live host inputs retain the target, focus and text selection; saved size is 25`, async () => {
    const v = await viewer({ controller: true, src: `samples/sample.${kind}`, editable: true, headless: true, onSave: true })
    await selectTarget(v, kind)
    expect((await state(v)).format).not.toBeNull()
    await v.page.evaluate(() => {
      for (const type of ['number', 'text']) {
        const input = document.createElement('input')
        input.id = `host-${type}`
        input.type = type
        input.value = type === 'number' ? '14' : 'Calibri'
        input.style.cssText = `position:fixed;top:0;left:${type === 'text' ? 220 : 0}px;z-index:999999;background:white`
        input.addEventListener('input', () => window.__ctl.format(type === 'number' ? { op: 'size', pt: Number(input.value) } : { op: 'font', name: input.value }))
        const sync = () => { input.disabled = !window.__ctl.snapshot().format }
        window.__ctl.subscribe(sync)
        sync()
        document.body.append(input)
      }
    })
    for (const type of ['number', 'text']) {
      const input = v.page.locator(`#host-${type}`)
      await input.click()
      await input.selectText()
      await settle(v)
      expect(await input.isEnabled()).toBe(true)
      expect((await state(v)).format).not.toBeNull()
      const value = type === 'number' ? '25' : 'Arial'
      await v.page.keyboard.type(value, { delay: 200 })
      await settle(v)
      expect(await input.inputValue()).toBe(value)
      expect(await input.evaluate((el) => el === document.activeElement)).toBe(true)
      expect((await state(v)).format?.size).toBe(25)
      if (type === 'text') {
        expect((await state(v)).format?.font).toBe('Arial')
        await input.evaluate((el) => (el as HTMLInputElement).setSelectionRange(0, 1))
        await v.page.evaluate(() => window.__ctl.format({ op: 'color', color: '#123456' }))
        await settle(v)
        expect(await input.evaluate((el) => [(el as HTMLInputElement).selectionStart, (el as HTMLInputElement).selectionEnd])).toEqual([0, 1])
        expect(await input.evaluate((el) => el === document.activeElement)).toBe(true)
      }
    }
    // PPTX commits are debounced (#18 is handled separately); verify the committed formatting here.
    if (kind === 'pptx') await v.page.waitForTimeout(500)
    await v.page.evaluate(() => window.__ctl.save())
    const saved = fromB64((await v.waitEvent('save')).b64 as string)
    if (kind === 'docx') {
      const p = (await docxParas(saved)).find((p) => paraText(p).startsWith('This document'))!
      expect(runsOf(p).find((r) => r.text === 'document')?.s.sz).toBe(25)
    } else {
      const { engine } = await openFile(saved)
      const slide: any = await engine.call('pptx.slide', [0])
      expect(slide.els.flatMap((e: any) => e.text?.paras ?? []).flatMap((p: any) => p.runs).find((r: any) => r.text === 'office-engine')?.sz).toBe(25)
    }
    // A caret uses the synchronous restore path, rather than the range's pending re-render restore.
    if (kind === 'docx') await v.selectText('document', 'end')
    else {
      await v.page.locator('.de-txt-edit').focus()
      await v.page.keyboard.press('ControlOrMeta+a')
      await v.page.keyboard.press('ArrowRight')
      await settle(v)
    }
    await v.page.locator('#host-number').click()
    await v.page.locator('#host-number').selectText()
    await v.page.keyboard.type('26', { delay: 200 })
    await settle(v)
    expect(await v.page.locator('#host-number').inputValue()).toBe('26')
    expect(await v.page.locator('#host-number').evaluate((el) => el === document.activeElement)).toBe(true)
    expect((await state(v)).format?.size).toBe(26)
    if (kind === 'docx') await v.selectText('document')
    else {
      await v.page.locator('.de-txt-edit').focus()
      await v.page.keyboard.press('ControlOrMeta+a')
      await settle(v)
    }
    await v.page.locator('#host-number').click()
    // A real DOM selection loss must also invalidate a pending range restore.
    await v.page.evaluate(() => {
      // A pending restore from this command must also respect the selection loss.
      window.__ctl.format({ op: 'size', pt: 27 })
      ;(document.activeElement as HTMLElement).blur()
      getSelection()!.removeAllRanges()
      document.dispatchEvent(new Event('selectionchange'))
    })
    await settle(v)
    expect((await state(v)).format).toBeNull()
    expect(await v.page.evaluate(() => document.activeElement?.matches('.de-pages, .de-txt-edit'))).toBe(false)
    await v.page.evaluate(() => window.__ctl.format({ op: 'size', pt: 70 }))
    expect((await state(v)).format).toBeNull()
    await v.page.evaluate(() => window.__ctl.setEditing(false))
    await settle(v)
    expect((await state(v)).format).toBeNull()
    expect(v.errors).toEqual([])
  })

  test(`${kind}: an older editing view cannot publish or clear its successor's format`, async () => {
    const v = await viewer({ controller: true, src: `samples/sample.${kind}`, editable: true, headless: true })
    await v.page.evaluate(() => window.__overlap(1))
    await v.waitEvent('load', 2)
    await settle(v)
    await selectTarget(v, kind, 1)
    const before = (await state(v)).format
    expect(before).not.toBeNull()
    // The old editor also hears this selection change; its null notification must be ignored.
    await v.page.evaluate(() => document.dispatchEvent(new Event('selectionchange')))
    await settle(v)
    expect((await state(v)).format).toEqual(before)
    await v.page.evaluate(() => window.__overlap(2))
    await settle(v)
    expect((await state(v)).format).toEqual(before)
    await v.page.evaluate(() => window.__ctl.format({ op: 'size', pt: 23 }))
    await settle(v)
    expect((await state(v)).format?.size).toBe(23)
    expect(v.errors).toEqual([])
  })

  test(`${kind}: controller replacement, edit exit and source replacement clear format ownership`, async () => {
    const v = await viewer({ controller: true, src: `samples/sample.${kind}`, editable: true, headless: true })
    await selectTarget(v, kind)
    await v.page.evaluate(async (path) => {
      const { createViewerController } = await import(path)
      const old = window.__ctl
      const next = createViewerController()
      ;(window as any).__old = old
      window.__ctl = next
      next.subscribe(() => {})
      window.__setProps({ controller: next })
    }, '/@fs/' + process.cwd() + '/src/index.ts')
    await settle(v)
    expect(await v.page.evaluate(() => (window as any).__old.snapshot().format)).toBeNull()
    expect((await state(v)).format).not.toBeNull()
    await v.page.evaluate(() => window.__ctl.setEditing(false))
    await settle(v)
    expect((await state(v)).format).toBeNull()
    await v.page.evaluate(() => window.__ctl.setEditing(true))
    await selectTarget(v, kind)
    await v.page.evaluate(() => window.__setProps({ src: 'samples/sample.xlsx' }))
    await v.waitEvent('load', 2)
    await v.page.locator('.de-grid').waitFor()
    await settle(v)
    expect((await state(v)).kind).toBe('xlsx')
    expect((await state(v)).format?.list).toBeUndefined()
    expect(await v.page.evaluate(() => (window as any).__old.snapshot().format)).toBeNull()
    expect(v.errors).toEqual([])
  })
}


test('xlsx: live host number input retains focus and saved cell value through typing and Backspace', async () => {
  const v = await viewer({ controller: true, src: 'samples/sample.xlsx', editable: true, headless: true, onSave: true })
  await v.typeCell('B2', 'keep-value')
  await v.clickCell('B2')
  await v.page.evaluate(() => {
    const input = document.createElement('input')
    input.type = 'number'; input.id = 'host-size'; input.value = ''
    input.style.cssText = 'position:fixed;top:0;left:0;z-index:999999'
    input.oninput = () => window.__ctl.format({ op: 'size', pt: input.valueAsNumber })
    document.body.append(input)
  })
  const input = v.page.locator('#host-size')
  await input.click()
  await v.page.keyboard.type('14', { delay: 200 })
  await settle(v)
  expect(await input.inputValue()).toBe('14')
  expect(await input.evaluate((el) => el === document.activeElement)).toBe(true)
  await v.page.keyboard.press('Backspace')
  await settle(v)
  expect(await input.inputValue()).toBe('1')
  await v.page.keyboard.type('4')
  await settle(v)
  await v.page.evaluate(() => window.__ctl.save())
  const bytes = fromB64((await v.waitEvent('save')).b64 as string)
  expect((await xlsxSheet(bytes)).cell('B2')).toBe('keep-value')
  expect((await state(v)).format?.size).toBe(14)
  expect(strFromU8(unzipSync(bytes)['xl/styles.xml'])).toContain('val="14"')
  // Direct calls retain whichever element the host focused; the grid's keyboard still works afterward.
  await v.page.evaluate(() => window.__ctl.format({ op: 'size', pt: 18 }))
  await settle(v)
  expect(await input.evaluate((el) => el === document.activeElement)).toBe(true)
  await v.page.evaluate(() => window.__ctl.undo())
  await settle(v)
  expect((await state(v)).format?.size).toBe(14)
  await v.page.evaluate(() => window.__ctl.redo())
  await settle(v)
  expect((await state(v)).format?.size).toBe(18)
  await v.page.evaluate(() => window.__ctl.undo())
  await settle(v)
  await v.page.evaluate((data) => window.__setProps({ src: new Blob([new Uint8Array(data)]), name: 'reopened.xlsx' }), Array.from(bytes))
  await v.waitEvent('load', 2)
  await v.page.locator('.de-grid').waitFor()
  await settle(v)
  await v.clickCell('B2')
  expect(await v.shown('B2')).toBe('keep-value')
  expect((await state(v)).format?.size).toBe(14)
  await v.page.keyboard.press('ArrowRight')
  expect(await v.page.locator('.de-fx-ref').innerText()).toBe('C2')
  await v.page.evaluate(() => window.__ctl.goto('1'))
  await settle(v)
  expect((await state(v)).outline?.items.find((i) => i.current)?.id).toBe('1')

  expect(v.errors).toEqual([])
})

test('docx: controlled host rerenders with inline callbacks preserve the logical target', async () => {
  const v = await viewer({ controller: true, src: 'samples/sample.docx', editable: true, headless: true, onSave: true, renderHost: true })
  await v.selectText('document')
  const input = v.page.getByLabel('React host size')
  await input.click()
  await settle(v)
  expect(await input.isEnabled()).toBe(true)
  await input.selectText()
  await v.page.keyboard.type('14', { delay: 200 })
  await settle(v)
  expect(await input.inputValue()).toBe('14')
  expect((await state(v)).format?.size).toBe(14)
  await v.page.evaluate(() => window.__ctl.save())
  const bytes = fromB64((await v.waitEvent('save')).b64 as string)
  expect(runsOf((await docxParas(bytes)).find((p) => paraText(p).startsWith('This document'))!).find((r) => r.text === 'document')?.s.sz).toBe(14)
})

test('docx without controller: Ctrl+S then Ctrl+K preserves selection and saves the link', async () => {
  const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
  await v.selectText('document')
  await v.page.keyboard.press('Control+s')
  await v.waitEvent('save')
  await v.page.keyboard.press('Control+k')
  await v.page.locator('.de-dialog input').fill('https://example.org/after-save')
  await v.page.keyboard.press('Enter')
  await settle(v)
  const bytes = (await v.save()).bytes
  const p = (await docxParas(bytes)).find((p) => paraText(p).startsWith('This document'))!
  expect(runsOf(p).find((r) => r.text === 'document')?.href).toBe('https://example.org/after-save')
  expect(v.errors).toEqual([])
})

for (const kind of ['docx', 'pptx'] as const) {
  test(`${kind}: a host format command does not suppress later undo caret restoration`, async () => {
    const v = await viewer({ controller: true, src: `samples/sample.${kind}`, editable: true, headless: true, onSave: true })
    await selectTarget(v, kind)
    const original = (await state(v)).format!.size
    await v.page.evaluate(() => {
      const input = document.createElement('input')
      input.id = 'focus-host'; input.style.cssText = 'position:fixed;top:0;left:0;z-index:999999'
      const undo = document.createElement('button')
      undo.id = 'undo-host'; undo.textContent = 'Undo host'
      undo.style.cssText = 'position:fixed;top:0;left:220px;z-index:999999'
      undo.onclick = window.__ctl.undo
      document.body.append(input, undo)
    })
    await v.page.locator('#focus-host').click()
    await v.page.evaluate(() => window.__ctl.format({ op: 'size', pt: 35 }))
    await settle(v)
    expect(await v.page.locator('#focus-host').evaluate((el) => el === document.activeElement)).toBe(true)
    await v.page.locator('#undo-host').click()
    await settle(v)
    expect(await v.page.evaluate(() => document.activeElement?.matches('.de-pages, .de-txt-edit'))).toBe(true)
    expect((await state(v)).format?.size).toBe(original)
    expect(await v.page.evaluate(() => getSelection()?.toString())).toBe(kind === 'docx' ? 'document' : 'office-engine')
    await v.page.evaluate(() => window.__ctl.redo())
    await settle(v)
    expect((await state(v)).format?.size).toBe(35)
    expect(v.errors).toEqual([])
  })
}

for (const kind of ['docx', 'xlsx', 'pptx'] as const) {
  test(`${kind}: invalid runtime commands preserve dirty/history/XML and opaque colors save and reopen`, async () => {
    const v = await viewer({ controller: true, src: `samples/sample.${kind}`, editable: true, headless: true, onSave: true })
    if (kind === 'xlsx') await v.clickCell('B2')
    else await selectTarget(v, kind)
    const invalid = async () => {
      await v.page.evaluate(() => {
        const commands = [
          { op: 'bold' }, { op: 'font', name: '' }, { op: 'font', name: '  ' },
          ...['red', 'rgb(255,0,0)', '#ff000080'].flatMap((color) => [{ op: 'color', color }, { op: 'highlight', color }]),
          ...[0.5, 1000, NaN, Infinity, Number.MIN_VALUE, Number.MAX_VALUE].map((pt) => ({ op: 'size', pt })),
        ]
        commands.forEach((c) => (window.__ctl.format as (c: unknown) => void)(c))
      })
      await settle(v)
    }
    const before = await state(v)
    await invalid()
    expect(formatting(await state(v))).toEqual(formatting(before))
    const path = kind === 'docx' ? 'word/document.xml' : kind === 'pptx' ? 'ppt/slides/slide1.xml' : 'xl/styles.xml'
    const saveXML = async () => {
      const count = (await v.events('save')).length
      await v.page.evaluate(() => window.__ctl.save())
      const bytes = fromB64((await v.waitEvent('save', count + 1)).b64 as string)
      return { bytes, xml: strFromU8(unzipSync(bytes)[path]) }
    }
    const initial = await saveXML()
    await v.page.evaluate(() => {
      window.__ctl.format({ op: 'color', color: '#123456ff' })
      window.__ctl.format({ op: 'highlight', color: '#abc' })
    })
    await v.page.waitForTimeout(kind === 'pptx' ? 550 : 200)
    const colored = await saveXML()
    expect(colored.xml).not.toBe(initial.xml)
    expect(colored.xml).toMatch(/123456/i)
    expect(colored.xml).toMatch(/AABBCC/i)
    const stable = await state(v)
    await invalid()
    expect(formatting(await state(v))).toEqual(formatting(stable))
    expect((await saveXML()).xml).toBe(colored.xml)
    await v.page.evaluate((data) => window.__setProps({ src: new Blob([new Uint8Array(data.bytes)]), name: `reopened.${data.kind}` }), { bytes: Array.from(colored.bytes), kind })
    await v.waitEvent('load', 2)
    await v.page.waitForTimeout(500)
    if (kind === 'xlsx') await v.clickCell('B2')
    else await selectTarget(v, kind)
    expect((await state(v)).format?.color?.toLowerCase()).toBe('#123456')
    expect((await state(v)).format?.highlight?.toLowerCase()).toBe('#aabbcc')
    expect(v.errors).toEqual([])
  })
}

test('shared SizeBox retains decimal-comma semantics and commits half points once', async () => {
  const v = await viewer({ controller: true, src: 'samples/sample.docx', editable: true, onSave: true })
  await v.selectText('document')
  const size = v.page.getByRole('textbox', { name: 'Font size', exact: true })
  await size.fill('12,5')
  await size.press('Enter')
  await settle(v)
  expect((await state(v)).format?.size).toBe(12.5)
  const p = (await docxParas((await v.save()).bytes)).find((p) => paraText(p).startsWith('This document'))!
  expect(runsOf(p).find((r) => r.text === 'document')?.s.sz).toBe(12.5)
  await v.page.evaluate(() => window.__ctl.undo())
  await settle(v)
  expect((await state(v)).canUndo).toBe(false)
})
