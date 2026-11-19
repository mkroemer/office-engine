/** QA round 2: Word editor findings (N2-0xx, P3-034) that need no browser */
import { afterEach, describe, expect, test } from 'bun:test'
import type { Block, Inline, Paragraph, Table } from '../src/core/docx/types'
import { htmlToParas } from '../src/core/html'
import { DocumentClient } from '../src/core/worker/client'
import { startCollab } from '../src/collab'
import { spliceRuns } from '../src/collab/docx'
import { CollabHub } from '../src/collab/hub'
import { registerHost } from '../src/collab/hosts'
import { MemoryTransport } from '../src/collab/transport'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { DocxEditor, type DocState } from '../src/react/docx/DocxEditor'
import { findAll } from '../src/react/docx/find'
import { opEnter, opInsert } from '../src/react/edit/input'
import { linkSpan, plen, runsText, type RRun, type Sel } from '../src/react/edit/rich'

const clients: DocumentClient[] = []
afterEach(() => {
  for (const c of clients.splice(0)) c.destroy()
})

async function editor() {
  const client = await DocumentClient.open(Bun.file(`${import.meta.dir}/../playground/public/samples/sample.docx`) as unknown as Blob, { worker: false })
  clients.push(client)
  const st: DocState = { blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }
  client.on('docx:blocks', (b) => st.blocks.push(...b))
  await client.call('docx.start')
  const ed = new DocxEditor(client, st, () => {})
  return { client, ed, st }
}

async function until(ed: DocxEditor, cond: () => boolean) {
  for (let i = 0; i < 2000 && (!cond() || ed.busy); i++) await Bun.sleep(2)
  expect(cond()).toBe(true)
}

const text = (b: Block | undefined) => (b?.t === 'p' ? runsText(b.runs as RRun[]) : '')
const caretAt = (pa: string, o = 0): Sel => ({ a: { pa, o }, f: { pa, o } })
/** an empty paragraph after the first one, the caret in it: pasted lines are not joined to text around them */
async function emptyPara(ed: DocxEditor, st: DocState) {
  const pa = st.blocks.findIndex((b) => text(b).startsWith('This document'))
  opEnter(ed, caretAt(String(pa), plen((st.blocks[pa] as Paragraph).runs as RRun[])))
  await until(ed, () => text(st.blocks[pa + 1]) === '')
  return caretAt(String(pa + 1))
}
const clip = (html: string) => ({ files: [] as unknown as FileList, getData: (t: string) => (t === 'text/html' ? html : '') }) as unknown as DataTransfer
const blocksText = (st: DocState) => st.blocks.map(text)

describe('N2-002 / N2-003: lists in pasted HTML', () => {
  test('Word list paragraphs (mso-list) are list items, with their level and kind', () => {
    const html =
      `<p class=MsoListParagraphCxSpFirst style='text-indent:-18.0pt;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>Item one<o:p></o:p></p>` +
      `<p class=MsoListParagraphCxSpMiddle style='margin-left:72pt;mso-list:l0 level2 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>o<span>&nbsp;&nbsp;</span></span><![endif]>Item two</p>` +
      `<p class=MsoListParagraphCxSpLast style='mso-list:l1 level1 lfo2'><![if !supportLists]><span style='mso-list:Ignore'>1.<span>&nbsp;&nbsp;</span></span><![endif]>Numbered</p>` +
      `<p class=MsoNormal>Plain</p>`
    for (const dom of [true, false]) {
      const p = htmlToParas(html, dom)
      expect(p.map((x) => [x.runs.map((r) => (r.t === 'text' ? r.text : '')).join(''), x.list, x.level])).toEqual([
        ['Item one', 'bullet', undefined],
        ['Item two', 'bullet', 1],
        ['Numbered', 'number', undefined],
        ['Plain', undefined, undefined],
      ])
    }
  })

  test('nested HTML lists carry their level', () => {
    const p = htmlToParas('<ul><li>a<ul><li>b<ol><li>c</li></ol></li></ul></li><li>d</li></ul>')
    expect(p.map((x) => [x.list, x.level])).toEqual([['bullet', undefined], ['bullet', 1], ['number', 2], ['bullet', undefined]])
  })

  test('a paragraph pasted after a list is not an item', async () => {
    const { ed, st } = await editor()
    const at0 = await emptyPara(ed, st)
    expect(ed.pasteRich(at0, clip('<p>Intro</p><ol><li>One</li><li>Two</li></ol><p>Closing</p>'))).toBe(true)
    await until(ed, () => blocksText(st).includes('Closing'))
    const at = (t: string) => st.blocks.find((b) => text(b) === t) as Paragraph
    expect(at('One').lst).toBe('number')
    expect(at('Two').lst).toBe('number')
    expect(at('Closing').lst).toBeUndefined()
    expect(at('Closing').marker).toBeUndefined()
  })

  test('a Word bullet is pasted as a list item', async () => {
    const { ed, st } = await editor()
    const at0 = await emptyPara(ed, st)
    const item = (t: string) => `<p class=MsoListParagraph style='mso-list:l0 level1 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>·<span>&nbsp;&nbsp;</span></span><![endif]>${t}</p>`
    ed.pasteRich(at0, clip(item('Bullet text') + item('Second')))
    await until(ed, () => blocksText(st).includes('Second'))
    expect(blocksText(st).join('\n')).not.toContain('·')
    for (const t of ['Bullet text', 'Second']) expect((st.blocks.find((b) => text(b) === t) as Paragraph).lst).toBe('bullet')
  })
})

describe('N2-004: tables in pasted HTML', () => {
  const TABLE = '<table><tr><td>a1</td><td colspan=2>b1</td></tr><tr><td rowspan=2>a2</td><td>b2</td><td>c2</td></tr><tr><td>b3</td></tr></table>'

  test('rich reading keeps the cells with their spans; plain reading flattens', () => {
    const rich = htmlToParas(`<p>before</p>${TABLE}<p>after</p>`, true, { tables: true })
    expect(rich.map((p) => !!p.table)).toEqual([false, true, false])
    expect(rich[1].table!.map((r) => r.map((c) => [c.paras.map((q) => q.runs.map((x) => (x.t === 'text' ? x.text : '')).join('')).join('|'), c.colSpan, c.rowSpan]))).toEqual([
      [['a1', 1, 1], ['b1', 2, 1]],
      [['a2', 1, 2], ['b2', 1, 1], ['c2', 1, 1]],
      [['b3', 1, 1]],
    ])
    expect(htmlToParas(TABLE).map((p) => p.table)).toEqual([undefined, undefined, undefined])
  })

  test('pasted into a paragraph: a table between the two halves, the caret after it', async () => {
    const { ed, st } = await editor()
    const pa = String(st.blocks.findIndex((b) => text(b).startsWith('This document')))
    const before = st.blocks.filter((b) => b.t === 'tbl').length
    ed.pasteRich(caretAt(pa, 4), clip(TABLE))
    await until(ed, () => st.blocks.filter((b) => b.t === 'tbl').length === before + 1)
    const i = st.blocks.findIndex((b, k) => b.t === 'tbl' && text(st.blocks[k - 1]) === 'This')
    expect(i).toBeGreaterThan(0)
    const t = st.blocks[i] as Table
    expect(t.cols.length).toBe(3)
    expect(t.rows.map((r) => r.cells.map((c) => c.blocks.map(text).join('')))).toEqual([['a1', 'b1'], ['a2', 'b2', 'c2'], ['b3', '']])
    expect(t.rows[1].cells[0].rowSpan).toBe(2)
    expect(t.rows[0].cells[1].colSpan).toBe(2)
    // the rest of the paragraph follows the table, which is where the caret is
    expect(text(st.blocks[i + 1]).startsWith(' document shows')).toBe(true)
    expect(ed.pendingSel).toEqual(caretAt(String(i + 1)))
    // and it is in the file
    const again = await DocumentClient.open((await ed['client'].call('save')) as Blob, { worker: false })
    clients.push(again)
    const got: Block[] = []
    again.on('docx:blocks', (b) => got.push(...b))
    await again.call('docx.start')
    expect(got.filter((b) => b.t === 'tbl').length).toBe(before + 1)
  })
})

describe('N2-005: pictures in pasted HTML', () => {
  const PNG1 = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP4z8DAwMDAxMDAwMDAAAANHQEDK+mmyAAAAABJRU5ErkJggg=='
  test('rich reading turns a data: picture into a run; plain reading leaves it out', () => {
    const html = `<p>Pic <img src="data:image/png;base64,${PNG1}" width="40" height="30"> end <img src="https://example.com/x.png"></p>`
    const [p] = htmlToParas(html, true, { tables: false })
    expect(p.runs.map((r) => r.t)).toEqual(['text', 'img', 'text'])
    expect(p.runs[1]).toMatchObject({ w: 30, h: 22.5 })
    expect(htmlToParas(html)[0].runs.every((r) => r.t === 'text')).toBe(true)
  })

  test('pasted into the document: stored as a picture of the file', async () => {
    const { ed, st } = await editor()
    const pa = String(st.blocks.findIndex((b) => text(b).startsWith('This document')))
    ed.pasteRich(caretAt(pa), clip(`<p>Look <img src="data:image/png;base64,${PNG1}" width="40" height="30"> here</p>`))
    await until(ed, () => blocksText(st).some((t) => t.startsWith('Look')))
    const p = st.blocks.find((b) => text(b).startsWith('Look')) as Paragraph
    const img = p.runs.find((r): r is Inline & { t: 'img' } => r.t === 'img')!
    expect(img).toBeTruthy()
    expect(img.rid).toBeTruthy()
    expect(img.target).toMatch(/^word\/media\//)
  })
})

describe('N2-009 / N2-010: copy inside the document', () => {
  test('copied list items carry no marker text and no color the text did not have', async () => {
    const { ed, st } = await editor()
    const a = st.blocks.findIndex((b) => text(b).startsWith('Revenue grew'))
    const c = st.blocks.findIndex((b) => text(b).startsWith('Customer churn'))
    const store: Record<string, string> = {}
    const dt = { setData: (t: string, v: string) => void (store[t] = v) } as unknown as DataTransfer
    expect(ed.copy({ a: { pa: String(a), o: 8 }, f: { pa: String(c), o: 8 } }, dt, false)).toBe(true)
    expect(store['text/html']).not.toContain('•')
    expect(store['text/plain']).toBe('•\tgrew 24% year over year\n•\tTwo new regions launched\n•\tCustomer')
    expect(store['text/html']).not.toMatch(/#000000|rgb\(0, 0, 0\)/)
    // pasted back: items, no marker characters
    const target = String(st.blocks.findIndex((b) => text(b).startsWith('Hire the platform')))
    ed.pasteRich(caretAt(target), clip(store['text/html']))
    await until(ed, () => blocksText(st).includes('CustomerHire the platform team'))
    expect(blocksText(st).join('\n')).not.toContain('•')
    expect((st.blocks.find((b) => text(b) === 'Two new regions launched') as Paragraph).lst).toBe('bullet')
    // pasted runs have no color of their own
    const p = st.blocks.find((b) => text(b) === 'CustomerHire the platform team') as Paragraph
    expect(p.runs.every((r) => r.t !== 'text' || r.s.color === undefined)).toBe(true)
  })

  test('cut puts the text on the clipboard and deletes it', async () => {
    const { ed, st } = await editor()
    const a = st.blocks.findIndex((b) => text(b).startsWith('This document'))
    const store: Record<string, string> = {}
    const dt = { setData: (t: string, v: string) => void (store[t] = v) } as unknown as DataTransfer
    expect(ed.copy({ a: { pa: String(a), o: 0 }, f: { pa: String(a), o: 4 } }, dt, true)).toBe(true)
    expect(store['text/plain']).toBe('This')
    expect(text(st.blocks[a]).startsWith(' document')).toBe(true)
  })
})

describe('N2-006: Enter in an empty list item', () => {
  test('leaves the list; one level out first in a nested item', async () => {
    const { ed, st } = await editor()
    const i = st.blocks.findIndex((b) => text(b).startsWith('Customer churn'))
    // a new empty item after the last bullet
    const end = plen((st.blocks[i] as Paragraph).runs as RRun[])
    opEnter(ed, caretAt(String(i), end))
    await until(ed, () => !!st.blocks[i + 1] && text(st.blocks[i + 1]) === '' && (st.blocks[i + 1] as Paragraph).lst === 'bullet')
    opEnter(ed, caretAt(String(i + 1)))
    await until(ed, () => (st.blocks[i + 1] as Paragraph).lst === undefined)
    // no third paragraph was made: the empty item itself became the plain paragraph
    expect(text(st.blocks[i + 1])).toBe('')
    expect(text(st.blocks[i + 2])).toBe('2. Next steps')
    expect((st.blocks[i + 1] as Paragraph).marker).toBeUndefined()
  })
})

describe('N2-016 / N2-017: a caret inside a link', () => {
  const link = (t: string, href?: string): RRun => ({ t: 'text', text: t, href, s: {} })
  test('linkSpan', () => {
    const runs = [link('Visit '), link('example', 'https://example.com'), link('.com', 'https://example.com'), link(' now')]
    expect(linkSpan(runs, 8)).toEqual([6, 17])
    expect(linkSpan(runs, 6)).toBeNull()
    expect(linkSpan(runs, 17)).toEqual([6, 17])
    expect(linkSpan(runs, 18)).toBeNull()
    expect(linkSpan(runs, 3)).toBeNull()
    expect(linkSpan([link('only', 'https://a.b')], 0)).toEqual([0, 4])
  })

  async function withLink() {
    const { ed, st } = await editor()
    const pa = String(st.blocks.findIndex((b) => text(b).startsWith('This document')))
    const sel: Sel = { a: { pa, o: 5 }, f: { pa, o: 13 } }
    ed.input.lastSel = sel
    await ed.setLink('https://example.com')
    await until(ed, () => ((st.blocks[Number(pa)] as Paragraph).runs as Inline[]).some((r) => r.t === 'text' && r.href === 'https://example.com'))
    return { ed, st, pa }
  }
  // (the paragraph has a link of its own, "example.com", further on: it is not the one under test)
  const linkRuns = (st: DocState, pa: string) => ((st.blocks[Number(pa)] as Paragraph).runs as (Inline & { t: 'text' })[]).filter((r) => r.href && r.text !== 'example.com')

  test('Remove with only a caret removes the link it is in', async () => {
    const { ed, st, pa } = await withLink()
    ed.input.lastSel = caretAt(pa, 8)
    await ed.setLink(null)
    await until(ed, () => true)
    expect(linkRuns(st, pa)).toEqual([])
    expect(text(st.blocks[Number(pa)]).startsWith('This document shows')).toBe(true)
  })

  test('a new address with only a caret changes the link, no text is inserted', async () => {
    const { ed, st, pa } = await withLink()
    const before = text(st.blocks[Number(pa)])
    ed.input.lastSel = caretAt(pa, 8)
    await ed.setLink('https://vi.wikipedia.org/')
    await until(ed, () => true)
    expect(text(st.blocks[Number(pa)])).toBe(before)
    expect(linkRuns(st, pa).map((r) => [r.text, r.href])).toEqual([[before.slice(5, 13), 'https://vi.wikipedia.org/']])
    expect(ed.pendingSel).toEqual(caretAt(pa, 8))
  })
})

describe('N2-014: Find treats NFC and NFD alike', () => {
  const nfd = (s: string) => s.normalize('NFD')
  test('composed query, decomposed text and the other way round, with positions of the original', () => {
    const text = `Quốc ${nfd('Việt')} Nam, ${'Việt'.normalize('NFC')}`
    const want = nfd('Việt')
    const hits = findAll(text, 'Việt'.normalize('NFC'))
    expect(hits.length).toBe(2)
    expect(text.slice(hits[0][0], hits[0][1])).toBe(want)
    expect(text.slice(hits[1][0], hits[1][1])).toBe('Việt'.normalize('NFC'))
    expect(findAll(text, nfd('việt')).length).toBe(2)
  })
  test('a base letter does not match inside a composed one; case is ignored', () => {
    expect(findAll(`cafe ${nfd('café')} CAFÉ`, 'cafe')).toEqual([[0, 4]])
    expect(findAll('Hello hello', 'HELLO')).toEqual([[0, 5], [6, 11]])
    expect(findAll('abc', '')).toEqual([])
  })
})

describe('P3-034: a collaborator\'s inserted word is one revision', () => {
  test('characters inserted one by one with the same tracked change join into one run', () => {
    const tc = { k: 'ins' as const, author: 'Ann', date: '2026-11-19T10:00:00Z' }
    let runs: Inline[] = [{ t: 'text', text: 'The fox', s: {} }]
    let at = 3
    for (const ch of ' quickly') {
      runs = spliceRuns(runs, at, 0, [{ t: 'text', text: ch, s: {}, tc }])
      at++
    }
    expect(runs.map((r) => [r.t === 'text' ? r.text : '', r.t === 'text' ? !!r.tc : false])).toEqual([['The', false], [' quickly', true], [' fox', false]])
  })
})

describe('P3-034: tracked typing seen by the other editor', () => {
  test('a word typed one character at a time is one w:ins in both files', async () => {
    const hub = new CollabHub()
    const sample = () => Bun.file(`${import.meta.dir}/../playground/public/samples/sample.docx`) as unknown as Blob
    const open = async (name: string) => {
      const client = await DocumentClient.open(sample(), { worker: false })
      clients.push(client)
      const st: DocState = { blocks: [], heights: new WeakMap(), dirty: new Set(), frontier: 0 }
      client.on('docx:blocks', (b) => st.blocks.push(...b))
      const done = new Promise<void>((r) => client.on('docx:done', () => r()))
      void client.call('docx.start')
      await done
      const editor = new DocxEditor(client, st, () => {})
      registerHost(client, 'docx', { editor, st })
      const collab = startCollab(client, { transport: new MemoryTransport(hub, () => 1), room: 'd', user: { name } })
      return { client, st, editor, collab }
    }
    const settle = async () => {
      for (let i = 0; i < 300; i++) {
        await Bun.sleep(10)
        await Promise.all([a, b].map((x) => x.collab.session.idle()))
        if ([a, b].every((x) => x.collab.pending === 0 && x.collab.status === 'live') && a.collab.session.rev === b.collab.session.rev) return
      }
      throw new Error('did not settle')
    }
    const a = await open('Ann')
    const b = await open('Bob')
    await settle()
    a.editor.author = () => 'Ann'
    a.editor.setTracking(true)
    const pi = a.st.blocks.findIndex((x) => text(x).startsWith('Hire the platform team'))
    const pa = String(pi)
    let o = 'Hire the platform'.length
    a.editor.input.lastSel = caretAt(pa, o)
    for (const ch of ' quickly') {
      opInsert(a.editor, caretAt(pa, o), ch)
      o++
    }
    await settle()
    const ins = async (c: typeof a) => {
      const zip = await ZipArchive.open(new BlobSource((await c.client.call('save')) as Blob))
      const xml = await zip.text('word/document.xml')
      return [...xml.matchAll(/<w:ins [^>]*>([\s\S]*?)<\/w:ins>/g)].map((m) => [...m[1].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''))
    }
    expect(await ins(a)).toEqual([' quickly'])
    expect(await ins(b)).toEqual([' quickly'])
    for (const x of [a, b]) x.collab.close()
  })
})
