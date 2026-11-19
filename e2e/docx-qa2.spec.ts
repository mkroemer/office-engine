/** QA round 2 regressions of the Word editor (N2-0xx, P3-034): typing, lists, links, clipboard, tables, find, print */
import { describe, expect, test } from 'bun:test'
import type { Block, Paragraph, Table } from '../src/core/docx/types'
import { docxBlocks, docxParas, paraText, PNG, runsOf, useBrowser, viewer, type Viewer } from './harness'

useBrowser()

/** a new blank document in edit mode, caret in its first paragraph */
async function blank() {
  const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
  await v.menu('Document')
  await v.page.locator('.de-ebar').waitFor()
  await v.page.locator('.de-page-inner').first().click()
  return v
}

/** pastes `html` / `text` through the beforeinput the browser sends for Ctrl+V */
async function paste(v: Viewer, d: { html?: string; text?: string }) {
  await v.page.evaluate((x) => {
    const dt = new DataTransfer()
    if (x.html) dt.setData('text/html', x.html)
    dt.setData('text/plain', x.text ?? '')
    document.querySelector('.de-pages')!.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertFromPaste', dataTransfer: dt, bubbles: true, cancelable: true, composed: true }))
  }, d)
  await v.page.waitForTimeout(400)
}

const saved = async (v: Viewer) => docxBlocks((await v.save()).bytes)
const paras = (bs: Block[]) => bs.filter((b): b is Paragraph => b.t === 'p')
const texts = (bs: Block[]) => paras(bs).map(paraText)

describe('QA 2: typing across a page boundary', () => {
  test('N2-011: a paragraph that moves to the next page after an edit keeps the caret and the toolbar', async () => {
    const v = await blank()
    // fill the first page; the caret is in the empty paragraph that starts the second
    for (let k = 0; k < 200 && (await v.page.locator('.de-page').count()) < 2; k++) {
      await v.page.keyboard.insertText(`Line ${k}`)
      await v.page.keyboard.press('Enter')
    }
    await v.page.waitForTimeout(400)
    // the last line of the first page grows by a few lines: it no longer fits and moves to the next page, with the caret in it
    await v.page.keyboard.press('ArrowUp')
    await v.page.keyboard.press('End')
    await v.page.keyboard.insertText('word '.repeat(60))
    // (the page is measured again after the edit: wait for that, then look)
    await v.page.waitForTimeout(600)
    const now = await v.page.evaluate(() => {
      const s = getSelection()!
      const n = s.anchorNode && (s.anchorNode.nodeType === 1 ? (s.anchorNode as Element) : s.anchorNode.parentElement)
      const bold = document.querySelector<HTMLButtonElement>('button[aria-label^="Bold"]')
      return { inParagraph: !!n?.closest('[data-pa]'), toolbarOff: !!bold?.disabled, onPage2: n?.closest('.de-page')?.getAttribute('data-page') }
    })
    expect(now).toEqual({ inParagraph: true, toolbarOff: false, onPage2: '2' })
    // typing goes on where the caret was
    await v.page.keyboard.insertText('END')
    await v.page.waitForTimeout(300)
    const last = paras(await saved(v)).filter((p) => paraText(p).includes('word word'))
    expect(last.length).toBe(1)
    expect(paraText(last[0]).endsWith('word END')).toBe(true)
    expect(v.errors).toEqual([])
  })

  test('N2-012: typing without delays across a page boundary keeps every Enter', async () => {
    const v = await blank()
    const N = 100
    // events dispatched back to back (a macro, a fast typist): no frame between a paragraph break and the next characters
    await v.page.evaluate(async (n) => {
      const root = document.querySelector('.de-pages')!
      const fire = (inputType: string, data?: string) => root.dispatchEvent(new InputEvent('beforeinput', { inputType, data, bubbles: true, cancelable: true, composed: true }))
      for (let k = 0; k < n; k++) {
        for (const ch of `P${k}: text`) fire('insertText', ch)
        fire('insertParagraph')
        if (k % 10 === 9) await new Promise((r) => setTimeout(r, 0))
      }
    }, N)
    await v.page.waitForTimeout(500)
    const got = texts(await saved(v)).filter((t) => t)
    expect(got).toEqual(Array.from({ length: N }, (_, k) => `P${k}: text`))
  })
})

describe('QA 2: lists', () => {
  test('N2-006: Enter on an empty list item leaves the list', async () => {
    const v = await blank()
    await v.page.keyboard.insertText('One')
    await v.btn('Bulleted list').click()
    await v.page.keyboard.press('Enter')
    await v.page.keyboard.insertText('Two')
    await v.page.keyboard.press('Enter')
    await v.page.keyboard.press('Enter')
    await v.page.keyboard.insertText('After')
    await v.page.waitForTimeout(300)
    const ps = paras(await saved(v)).filter((p) => paraText(p))
    expect(ps.map((p) => [paraText(p), !!p.lst])).toEqual([['One', true], ['Two', true], ['After', false]])
    const all = paras(await saved(v))
    // the emptied item is the plain paragraph that now holds "After": no empty bullet is left behind
    expect(all.filter((p) => p.lst).length).toBe(2)
  })

  test('N2-006: so does Enter coming before the worker has said the paragraph is a list (the button, then keys without a frame between)', async () => {
    const v = await blank()
    await v.page.keyboard.insertText('One')
    await v.page.waitForTimeout(200)
    // the click and the keys in one task: nothing the worker answers gets in between
    await v.page.evaluate(() => {
      const root = document.querySelector('.de-pages')!
      const fire = (inputType: string, data?: string) => root.dispatchEvent(new InputEvent('beforeinput', { inputType, data, bubbles: true, cancelable: true, composed: true }))
      document.querySelector<HTMLButtonElement>('button[aria-label^="Bulleted list"]')!.click()
      fire('insertParagraph')
      for (const ch of 'Two') fire('insertText', ch)
      fire('insertParagraph')
      fire('insertParagraph')
      for (const ch of 'After') fire('insertText', ch)
    })
    await v.page.waitForTimeout(500)
    const all = paras(await saved(v))
    expect(all.filter((p) => paraText(p)).map((p) => [paraText(p), !!p.lst])).toEqual([['One', true], ['Two', true], ['After', false]])
    expect(all.filter((p) => p.lst).length).toBe(2)
    expect(v.errors).toEqual([])
  })
})

describe('QA 2: links', () => {
  const opened = async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.page.locator('.de-ebar').waitFor()
    // the caret inside the link "example.com", no selection
    await v.selectText('example.com', 'start')
    for (let k = 0; k < 3; k++) await v.page.keyboard.press('ArrowRight')
    return v
  }
  const linkOf = async (v: Viewer) => {
    const p = (await docxParas((await v.save()).bytes)).find((x) => paraText(x).startsWith('This document'))!
    return { text: paraText(p), runs: runsOf(p).filter((r) => r.text.includes('example')) }
  }

  test('N2-016: Remove with only a caret in a link removes that link', async () => {
    const v = await opened()
    await v.page.keyboard.press('Control+k')
    await v.page.locator('.de-dialog input').waitFor()
    expect(await v.page.locator('.de-dialog input').inputValue()).toContain('example.com')
    await v.page.locator('.de-dialog button', { hasText: 'Remove' }).click()
    await v.page.waitForTimeout(300)
    const l = await linkOf(v)
    expect(l.text).toContain('Visit example.com')
    expect(l.runs.every((r) => !r.href)).toBe(true)
  })

  test('N2-017: a new address with only a caret in a link changes the link, no text is added', async () => {
    const v = await opened()
    const before = (await linkOf(v)).text
    await v.selectText('example.com', 'start')
    for (let k = 0; k < 3; k++) await v.page.keyboard.press('ArrowRight')
    await v.page.keyboard.press('Control+k')
    await v.page.locator('.de-dialog input').fill('https://vi.wikipedia.org/')
    await v.page.keyboard.press('Enter')
    await v.page.waitForTimeout(300)
    const l = await linkOf(v)
    expect(l.text).toBe(before)
    expect(l.runs.map((r) => [r.text, r.href])).toEqual([['example.com', 'https://vi.wikipedia.org/']])
  })
})

describe('QA 2: clipboard', () => {
  /** selects from the start of `a` to the end of `b` (two paragraphs) */
  async function selectBetween(v: Viewer, a: string, b: string) {
    await v.page.evaluate(
      ([x, y]) => {
        const root = document.querySelector<HTMLElement>('.de-pages')!
        const find = (t: string) => {
          const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
          for (let n = w.nextNode(); n; n = w.nextNode()) if (n.textContent!.includes(t) && !n.parentElement?.closest('[contenteditable=false]')) return n
          throw new Error('no ' + t)
        }
        const na = find(x)
        const nb = find(y)
        root.focus()
        getSelection()!.setBaseAndExtent(na, na.textContent!.indexOf(x), nb, nb.textContent!.indexOf(y) + y.length)
      },
      [a, b],
    )
    await v.page.waitForTimeout(150)
  }

  test('N2-009 / N2-010: copying list items inside the document and pasting them: no marker character, no hard black', async () => {
    const v = await viewer({ src: 'samples/sample.docx', editable: true, onSave: true })
    await v.ctx.grantPermissions(['clipboard-read', 'clipboard-write'])
    await v.page.locator('.de-ebar').waitFor()
    await selectBetween(v, 'Revenue grew', 'Customer')
    await v.page.keyboard.press('Control+c')
    await v.selectText('Hire the platform team', 'start')
    await v.page.keyboard.press('Control+v')
    await v.page.waitForTimeout(500)
    const ps = paras(await saved(v))
    const texts = ps.map(paraText)
    expect(texts.join('\n')).not.toContain('•')
    const i = texts.indexOf('Two new regions launched', texts.indexOf('Customer churn fell to 3.1%'))
    expect(i).toBeGreaterThan(0)
    expect(ps[i].lst).toBe('bullet')
    expect(texts.some((t) => t === 'CustomerHire the platform team')).toBe(true)
    const pasted = ps.filter((p) => ['Revenue grew 24% year over year', 'Two new regions launched', 'CustomerHire the platform team'].includes(paraText(p)))
    const fresh = pasted.slice(-2)
    expect(fresh.flatMap((p) => runsOf(p).map((r) => r.s.color)).every((c) => c === undefined)).toBe(true)
  })
})

describe('QA 2: pasted HTML', () => {
  test('N2-002 / N2-003: a paragraph after a list is not numbered; Word list paragraphs become items', async () => {
    const v = await blank()
    await paste(v, { html: '<p>Intro</p><ol><li>Uno</li><li>Dos</li></ol><p>Closing</p>', text: 'x' })
    await v.page.keyboard.press('Enter')
    const item = (t: string) => `<p class=MsoListParagraph style='mso-list:l0 level1 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>·<span>&nbsp;&nbsp;</span></span><![endif]>${t}</p>`
    await paste(v, { html: item('Word one') + item('Word two') + '<p class=MsoNormal>Done</p>', text: 'x' })
    const ps = paras(await saved(v))
    const at = (t: string) => ps.find((p) => paraText(p) === t)!
    expect([at('Uno').lst, at('Dos').lst]).toEqual(['number', 'number'])
    expect(at('Closing').lst).toBeUndefined()
    expect([at('Word one').lst, at('Word two').lst]).toEqual(['bullet', 'bullet'])
    expect(at('Done')?.lst).toBeUndefined()
    expect(ps.map(paraText).join('\n')).not.toContain('·')
  })

  test('N2-004: an HTML table is pasted as a table', async () => {
    const v = await blank()
    await v.page.keyboard.insertText('Before')
    await v.page.keyboard.press('Enter')
    await paste(v, { html: '<table><tr><th>Name</th><th>Qty</th></tr><tr><td>Apple</td><td>12</td></tr><tr><td colspan=2>Total</td></tr></table>', text: 'Name\tQty\nApple\t12' })
    // the caret is after the table: typing goes there
    await v.page.keyboard.insertText('After')
    await v.page.waitForTimeout(300)
    const bs = await saved(v)
    const t = bs.find((b): b is Table => b.t === 'tbl')!
    expect(t).toBeTruthy()
    const cellText = (c: Table['rows'][number]['cells'][number]) => c.blocks.map((b) => (b.t === 'p' ? paraText(b) : '')).join('')
    expect(t.rows.map((r) => r.cells.map(cellText))).toEqual([['Name', 'Qty'], ['Apple', '12'], ['Total']])
    expect(t.rows[2].cells[0].colSpan).toBe(2)
    expect(texts(bs)).toContain('Before')
    expect(texts(bs)).toContain('After')
    expect(bs.findIndex((b) => b.t === 'tbl')).toBeLessThan(bs.findIndex((b) => b.t === 'p' && paraText(b) === 'After'))
  })

  test('N2-005: a data: picture in pasted HTML is kept', async () => {
    const v = await blank()
    const url = `data:image/png;base64,${PNG.toString('base64')}`
    await paste(v, { html: `<p>Look <img src="${url}" width="40" height="40"> here</p>`, text: 'Look here' })
    await v.page.waitForFunction(() => document.querySelectorAll('.de-pages img').length > 0)
    const p = paras(await saved(v)).find((x) => paraText(x).includes('Look'))!
    expect(p.runs.some((r) => r.t === 'img')).toBe(true)
    expect(paraText(p)).toBe('Look  here')
  })
})

describe('QA 2: table, find, print', () => {
  test('N2-013: ArrowRight after the last character of a table leaves it', async () => {
    const v = await blank()
    await v.btn('Insert table').click()
    // 2 × 2: 8 cells per picker row
    await v.page.locator('.de-gridpick button').nth(8 + 1).click()
    await v.page.waitForTimeout(300)
    for (const c of ['a', 'b', 'c', 'd']) {
      await v.page.keyboard.insertText(c)
      if (c !== 'd') await v.page.keyboard.press('Tab')
    }
    await v.page.keyboard.press('ArrowRight')
    await v.page.keyboard.insertText('After')
    await v.page.waitForTimeout(300)
    const bs = await saved(v)
    const t = bs.find((b): b is Table => b.t === 'tbl')!
    expect(t.rows.map((r) => r.cells.map((c) => c.blocks.map((b) => (b.t === 'p' ? paraText(b) : '')).join('')))).toEqual([['a', 'b'], ['c', 'd']])
    expect(texts(bs)).toContain('After')
  })

  test('N2-014: Find matches text typed decomposed with a composed query', async () => {
    const v = await blank()
    await v.page.keyboard.insertText('Vie\u0302\u0323t Nam, Vi\u1ec7t')
    await v.page.waitForTimeout(200)
    await v.btn('Find & replace').click()
    await v.page.locator('.de-find input[aria-label="Find"]').fill('Vi\u1ec7t')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-find-count').innerText()).toBe('1 / 2')
    await v.page.locator('.de-find input[aria-label="Find"]').fill('vie\u0302\u0323t')
    await v.page.waitForTimeout(300)
    expect(await v.page.locator('.de-find-count').innerText()).toBe('1 / 2')
  })

  test('N2-015: a table row taller than a page is not cut off in print', async () => {
    const v = await blank()
    await v.btn('Insert table').click()
    await v.page.locator('.de-gridpick button').first().click()
    await v.page.waitForTimeout(300)
    await paste(v, { text: Array.from({ length: 120 }, (_, k) => `Row line ${k + 1}`).join('\n') })
    await v.page.waitForFunction(() => document.body.innerText.includes('Row line 120'))
    await v.page.evaluate(() => {
      ;(window as any).__printed = []
      new MutationObserver((ms) => {
        for (const m of ms)
          for (const n of m.addedNodes) {
            if ((n as HTMLElement).tagName !== 'IFRAME') continue
            const w = (n as HTMLIFrameElement).contentWindow!
            w.print = () => {
              const d = w.document
              const last = [...d.querySelectorAll('p, div, span, td')].filter((e) => e.children.length === 0 && e.textContent === 'Row line 120').pop()!
              const r = last.getBoundingClientRect()
              // a clip (overflow other than visible) that ends above the line hides it
              let clipped = false
              for (let a: Element | null = last.parentElement; a && a !== d.documentElement; a = a.parentElement) {
                const cs = w.getComputedStyle(a)
                if ((cs.overflowY !== 'visible' || cs.overflowX !== 'visible') && r.bottom > a.getBoundingClientRect().bottom + 1) clipped = true
              }
              ;(window as any).__printed.push({ clipped, pages: d.querySelectorAll('.de-print-page').length })
            }
            w.document.execCommand = () => false
          }
      }).observe(document.body, { childList: true })
    })
    await v.menu('Print')
    await v.page.waitForFunction(() => (window as any).__printed?.length > 0, null, { timeout: 30_000 })
    const out = await v.page.evaluate(() => (window as any).__printed.at(-1))
    expect(out.clipped).toBe(false)
  })
})
