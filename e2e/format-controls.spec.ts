import { expect, test } from 'bun:test'
import { docxParas, fromB64, openFile, paraText, runsOf, useBrowser, viewer, type Viewer } from './harness'

useBrowser()
const state = (v: Viewer) => v.page.evaluate(() => window.__ctl.snapshot())
const settle = (v: Viewer) => v.page.waitForTimeout(150)

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
