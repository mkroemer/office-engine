import { expect, test } from 'bun:test'
import { fromB64, openFile, useBrowser, viewer, type Viewer } from './harness'
import type { SlideModel } from '../src/core/pptx/types'

useBrowser()

declare global {
  interface Window {
    __pptxSave: {
      holdEdit?: boolean
      failEdit?: boolean
      failSave?: boolean
      holdHost?: boolean
      failHost?: boolean
      release?: () => void
      calls: string[]
    }
  }
}

/** Gate actual worker calls, without relying on debounce timing or worker speed. */
async function controlledViewer(mode: 'holdEdit' | 'failEdit' | 'holdHost') {
  const v = await viewer({ editable: true, onSave: true })
  await v.page.evaluate((mode) => {
    window.__pptxSave = { [mode]: true, calls: [] }
    window.__setProps({
      src: 'samples/sample.pptx',
      onLoad(client) {
        window.__events.push({ type: 'load', kind: client.kind })
        const call = client.call.bind(client)
        client.call = (async (method: string, ...args: unknown[]) => {
          const state = window.__pptxSave
          if (method === 'pptx.setText' || method === 'save') state.calls.push(method)
          if (method === 'pptx.setText') {
            if (state.failEdit) throw new Error('Slide commit rejected')
            if (state.holdEdit) {
              state.holdEdit = false
              await new Promise<void>((resolve) => { state.release = resolve })
            }
          }
          if (method === 'save' && state.failSave) throw new Error('Serialization rejected')
          return Reflect.apply(call, client, [method, ...args])
        }) as typeof client.call
      },
      onSave: async (blob, name) => {
        const state = window.__pptxSave
        if (state.failHost) throw new Error('Host write rejected')
        const bytes = new Uint8Array(await blob.arrayBuffer())
        let binary = ''
        for (const byte of bytes) binary += String.fromCharCode(byte)
        window.__events.push({ type: 'save', name, b64: btoa(binary) })
        if (state.holdHost) {
          state.holdHost = false
          await new Promise<void>((resolve) => { state.release = resolve })
        }
      },
    })
  }, mode)
  await v.ready('pptx')
  return v
}

async function editTitle(v: Viewer) {
  await v.page.locator('.de-slide-frame').first().getByText('office-engine', { exact: true }).click()
  await v.page.locator('.de-txt-edit').waitFor()
  await v.page.keyboard.press('ControlOrMeta+a')
}

async function savedRuns(b64: string) {
  const { engine } = await openFile(fromB64(b64))
  const slide = await engine.call('pptx.slide', [0]) as SlideModel
  return slide.els.flatMap((e) => e.t === 'shape' ? e.text?.paras.flatMap((p) => p.runs) ?? [] : [])
}

test('immediate keyboard save includes pending slide text and formatting', async () => {
  const v = await viewer({ src: 'samples/sample.pptx', editable: true, onSave: true })
  await editTitle(v)
  await v.page.keyboard.type('Saved before debounce')
  await v.page.keyboard.press('ControlOrMeta+a')
  await v.page.keyboard.press('ControlOrMeta+b')
  await v.page.keyboard.press('ControlOrMeta+s')
  const file = await v.waitEvent('save')
  const run = (await savedRuns(file.b64)).find((r) => r.text === 'Saved before debounce')
  expect(run).toBeDefined()
  expect(run?.b).toBe(true)
  await v.page.waitForFunction(() => !document.querySelector('.de-save-dirty'))
  expect((await v.events('dirty')).at(-1)?.dirty).toBe(false)
  // Undo after a successful save must become dirty again, and flush just like typing.
  await v.page.keyboard.press('ControlOrMeta+z')
  expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
  await v.page.keyboard.press('ControlOrMeta+s')
  const undone = await v.waitEvent('save', 2)
  expect((await savedRuns(undone.b64)).find((r) => r.text === 'Saved before debounce')?.b).not.toBe(true)
  expect(v.errors).toEqual([])
})

test('save awaits an in-flight commit and newer active edits; repeated shortcuts do not overlap', async () => {
  const v = await controlledViewer('holdEdit')
  await editTitle(v)
  await v.page.keyboard.type('First commit')
  await v.page.keyboard.press('Escape')
  await v.page.waitForFunction(() => !!window.__pptxSave.release)
  // Dispatch in the same turn to exercise the guard before React can render Saving….
  await v.page.evaluate(() => {
    for (let i = 0; i < 2; i++) document.activeElement!.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true }))
  })
  expect(await v.events('save')).toHaveLength(0)
  expect(await v.page.evaluate(() => window.__pptxSave.calls)).toEqual(['pptx.setText'])
  await editTitle(v)
  await v.page.keyboard.type('Newest commit')
  await v.page.evaluate(() => window.__pptxSave.release!())
  const file = await v.waitEvent('save')
  expect((await savedRuns(file.b64)).some((r) => r.text === 'Newest commit')).toBe(true)
  await v.page.waitForFunction(() => !document.querySelector('.de-save-dirty'))
  expect(await v.page.evaluate(() => window.__pptxSave.calls)).toEqual(['pptx.setText', 'pptx.setText', 'save'])
  expect(v.errors).toEqual([])
})

test('failed slide commits survive ending text editing; failed serialization also stays dirty for retry', async () => {
  const v = await controlledViewer('failEdit')
  await editTitle(v)
  await v.page.keyboard.type('Retained after failure')
  await v.page.keyboard.press('Escape')
  await v.page.waitForFunction(() => window.__pptxSave.calls.includes('pptx.setText'))
  await v.page.keyboard.press('ControlOrMeta+s')
  expect((await v.waitEvent('error')).message).toBe('Slide commit rejected')
  expect(await v.events('save')).toHaveLength(0)
  expect(await v.page.locator('.de-save-dirty').count()).toBe(1)
  expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)

  await v.page.evaluate(() => { window.__pptxSave.failEdit = false; window.__pptxSave.failSave = true })
  await v.page.keyboard.press('ControlOrMeta+s')
  expect((await v.waitEvent('error', 2)).message).toBe('Serialization rejected')
  expect(await v.events('save')).toHaveLength(0)
  expect(await v.page.locator('.de-save-dirty').count()).toBe(1)

  await v.page.evaluate(() => { window.__pptxSave.failSave = false })
  await v.page.keyboard.press('ControlOrMeta+s')
  const file = await v.waitEvent('save')
  expect((await savedRuns(file.b64)).some((r) => r.text === 'Retained after failure')).toBe(true)
  await v.page.waitForFunction(() => !document.querySelector('.de-save-dirty'))
  expect(v.errors).toHaveLength(3)
  expect(v.errors.every((e) => /Slide commit rejected|Serialization rejected/.test(e))).toBe(true)
})

test('a slide edit the worker refuses every time stops failing saves once it is given up', async () => {
  const v = await controlledViewer('failEdit')
  await editTitle(v)
  await v.page.keyboard.type('Never accepted')
  await v.page.keyboard.press('Escape')
  await v.page.waitForFunction(() => window.__pptxSave.calls.includes('pptx.setText'))
  // sent again at each save, refused again: after the third refusal it is dropped
  await v.page.keyboard.press('ControlOrMeta+s')
  expect((await v.waitEvent('error')).message).toBe('Slide commit rejected')
  await v.page.keyboard.press('ControlOrMeta+s')
  expect((await v.waitEvent('error', 2)).message).toBe('Slide commit rejected')
  expect(await v.events('save')).toHaveLength(0)
  // the next save goes through, without the edit
  await v.page.keyboard.press('ControlOrMeta+s')
  const file = await v.waitEvent('save')
  expect((await savedRuns(file.b64)).some((r) => r.text === 'Never accepted')).toBe(false)
  expect(await v.page.evaluate(() => window.__pptxSave.calls.filter((c) => c === 'pptx.setText').length)).toBe(3)
})

test('edits during host save and host write failures remain dirty until the next successful save', async () => {
  const v = await controlledViewer('holdHost')
  await editTitle(v)
  await v.page.keyboard.type('First saved version')
  await v.page.keyboard.press('ControlOrMeta+s')
  const first = await v.waitEvent('save')
  await v.page.waitForFunction(() => !!window.__pptxSave.release)
  await v.page.keyboard.press('ControlOrMeta+a')
  await v.page.keyboard.type('Edited during save')
  await v.page.keyboard.press('ControlOrMeta+s')
  expect(await v.events('save')).toHaveLength(1)
  await v.page.evaluate(() => window.__pptxSave.release!())
  await v.page.waitForFunction(() => !document.querySelector('button[title="Save a copy (Ctrl+S)"]:disabled'))
  expect((await savedRuns(first.b64)).some((r) => r.text === 'First saved version')).toBe(true)
  expect(await v.page.locator('.de-save-dirty').count()).toBe(1)
  expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)

  await v.page.evaluate(() => { window.__pptxSave.failHost = true })
  await v.page.keyboard.press('ControlOrMeta+s')
  expect((await v.waitEvent('error')).message).toBe('Host write rejected')
  expect(await v.page.locator('.de-save-dirty').count()).toBe(1)
  await v.page.evaluate(() => { window.__pptxSave.failHost = false })
  await v.page.keyboard.press('ControlOrMeta+s')
  const latest = await v.waitEvent('save', 2)
  expect((await savedRuns(latest.b64)).some((r) => r.text === 'Edited during save')).toBe(true)
  await v.page.waitForFunction(() => !document.querySelector('.de-save-dirty'))
  expect(v.errors).toHaveLength(1)
  expect(v.errors.every((e) => /Host write rejected/.test(e))).toBe(true)
})

test('completion of a save from a replaced document cannot clear the new document dirty state', async () => {
  const v = await controlledViewer('holdHost')
  await editTitle(v)
  await v.page.keyboard.type('Old document')
  await v.page.keyboard.press('ControlOrMeta+s')
  await v.waitEvent('save')
  await v.page.waitForFunction(() => !!window.__pptxSave.release)
  await v.page.evaluate(() => window.__setProps({ src: 'samples/sample.xlsx' }))
  await v.page.locator('.de-grid').waitFor()
  await v.typeCell('J10', 'New document edit')
  await v.page.evaluate(() => window.__pptxSave.release!())
  expect(await v.page.locator('.de-save-dirty').count()).toBe(1)
  expect((await v.events('dirty')).at(-1)?.dirty).toBe(true)
  // The replacement can save independently of the previous save session.
  expect((await v.save()).name).toBe('sample.xlsx')
  expect(v.errors).toEqual([])
})

test('unmounting the slide view during a flush cancels the stale save without an unhandled rejection', async () => {
  const v = await controlledViewer('holdEdit')
  await editTitle(v)
  await v.page.keyboard.type('Closing document')
  await v.page.keyboard.press('Escape')
  await v.page.waitForFunction(() => !!window.__pptxSave.release)
  await v.page.keyboard.press('ControlOrMeta+s')
  await v.page.evaluate(() => window.__setProps({ src: null }))
  await v.page.locator('.de-slide-inner').first().waitFor({ state: 'detached' })
  const closed = v.page.waitForEvent('console', { predicate: (m) => /slide edit failed.*Document closed/.test(m.text()) })
  await v.page.evaluate(() => window.__pptxSave.release!())
  // The background commit reports the closed client; save itself must not report a stale error.
  await closed
  expect(await v.events('save')).toHaveLength(0)
  expect(await v.events('error')).toHaveLength(0)
  expect(v.errors).toHaveLength(1)
  expect(v.errors.every((e) => /slide edit failed.*Document closed/.test(e))).toBe(true)
})

test('Office, text and printable exports await the same pending slide work', async () => {
  for (const format of ['PowerPoint presentation', 'Plain text', 'Web page']) {
    const v = await controlledViewer('holdEdit')
    await editTitle(v)
    await v.page.keyboard.type('Latest exported text')
    await v.page.keyboard.press('Escape')
    await v.page.waitForFunction(() => !!window.__pptxSave.release)
    const exporting = v.downloadAs(format)
    await v.page.waitForFunction(() => !!document.querySelector('.de-file:disabled'))
    expect(await v.page.evaluate(() => window.__pptxSave.calls)).toEqual(['pptx.setText'])
    await v.page.evaluate(() => window.__pptxSave.release!())
    const file = await exporting
    const text = format === 'PowerPoint presentation' ? JSON.stringify(await savedRuns(file.bytes.toString('base64'))) : file.text
    expect(text).toContain('Latest exported text')
    expect(v.errors).toEqual([])
  }
})
