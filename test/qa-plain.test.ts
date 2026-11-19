/**
 * QA 2026-11-18 (Markdown / HTML): a file edited as a document and saved keeps what the document model cannot hold;
 * escapes in Markdown are stable.
 */
import { describe, expect, test } from 'bun:test'
import type { Block, Paragraph } from '../src/core/docx/types'
import { mergeSource } from '../src/core/plainmerge'
import { Engine } from '../src/core/worker/engine'
import { docxBodyHtml, docxToMarkdown } from '../src/react/file/text'

const enc = (s: string) => new TextEncoder().encode(s)

async function blocksOf(text: string, name: string): Promise<Block[]> {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(enc(text), { name })
  await engine.call('docx.start', [])
  return events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
}

/** types one character at the end of the paragraph reading `text` */
function typeInto(blocks: Block[], text: string, ch = 'X'): void {
  const p = blocks.find((b): b is Paragraph => b.t === 'p' && b.runs.map((r) => (r.t === 'text' ? r.text : '')).join('').includes(text))!
  const run = [...p.runs].reverse().find((r) => r.t === 'text')!
  ;(run as { text: string }).text += ch
}

const saveMd = (src: string, blocks: Block[]) => mergeSource(src, 'markdown', blocks, (bs) => docxToMarkdown(bs)) ?? docxToMarkdown(blocks)
const saveHtml = (src: string, blocks: Block[]) => mergeSource(src, 'html', blocks, (bs) => docxBodyHtml(bs)) ?? docxBodyHtml(blocks)

const MD = `---
title: Notes
tags: [a, b]

---

# Title

A plain paragraph with **bold** and a [link](https://example.com).

![logo](https://example.com/logo.png)

| a | b |
| --- | --- |
| 1 | 2 |
| 3 | 4 |

- one
  - nested one
    - nested two
- two

1. first
2. second

\`\`\`ts

const x = 1

console.log(x)
\`\`\`

> a quote
> over two lines

Last paragraph.
`

const HTML = `<!doctype html>
<html><head><meta charset="utf-8"><title>Page</title><style>p{color:red}</style></head>
<body>
<div class="wrap">
<h1>Title</h1>
<p>A plain paragraph with <b>bold</b>.</p>
<table><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2</td></tr></table>
<ul><li>one<ul><li>nested</li></ul></li><li>two</li></ul>
<pre><code class="language-ts">const x = 1
console.log(x)</code></pre>
<blockquote>a quote</blockquote>
<p><img src="https://example.com/a.png" alt="pic"></p>
<p>Last paragraph.</p>
</div>
</body></html>
`

describe('Markdown / HTML saved after an edit as a document', () => {
  test('markdown: one typed character keeps tables, nested lists, pictures, code, quotes and front matter', async () => {
    const blocks = await blocksOf(MD, 'a.md')
    // not edited: the file as it is
    expect(saveMd(MD, blocks)).toBe(MD)
    typeInto(blocks, 'A plain paragraph')
    const out = saveMd(MD, blocks)
    expect(out).toBe(MD.replace('[link](https://example.com).', '[link](https://example.com).X'))
    for (const keep of ['---\ntitle: Notes\ntags: [a, b]\n\n---', '![logo](https://example.com/logo.png)', '| a | b |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |', '- one\n  - nested one\n    - nested two\n- two', '```ts\n\nconst x = 1\n\nconsole.log(x)\n```', '> a quote\n> over two lines']) expect(out).toContain(keep)
  })

  test('markdown: an edit in the last paragraph, a deleted paragraph', async () => {
    const blocks = await blocksOf(MD, 'a.md')
    typeInto(blocks, 'Last paragraph', '!')
    expect(saveMd(MD, blocks)).toBe(MD.replace('Last paragraph.', 'Last paragraph.!'))
    const b2 = await blocksOf(MD, 'a.md')
    const k = b2.findIndex((b) => b.t === 'p' && b.runs.some((r) => r.t === 'text' && r.text.startsWith('A plain')))
    b2.splice(k, 1)
    const out = saveMd(MD, b2)
    expect(out).not.toContain('A plain paragraph')
    expect(out).toContain('| 3 | 4 |')
    expect(out).toContain('---\ntitle: Notes')
  })

  test('markdown: a CRLF file keeps its line breaks', async () => {
    const src = MD.replace(/\n/g, '\r\n')
    const blocks = await blocksOf(src, 'a.md')
    typeInto(blocks, 'A plain paragraph')
    const out = saveMd(src, blocks)
    expect(out.replace(/\r\n/g, '')).not.toContain('\n')
    expect(out).toContain('| 1 | 2 |\r\n')
  })

  test('html: one typed character keeps the markup around', async () => {
    const blocks = await blocksOf(HTML, 'a.html')
    expect(saveHtml(HTML, blocks)).toBe(HTML)
    typeInto(blocks, 'A plain paragraph')
    const out = saveHtml(HTML, blocks)
    for (const keep of ['<!doctype html>', '<style>p{color:red}</style>', '<div class="wrap">', '<table><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2</td></tr></table>', '<ul><li>one<ul><li>nested</li></ul></li><li>two</li></ul>', '<pre><code class="language-ts">const x = 1\nconsole.log(x)</code></pre>', '<blockquote>a quote</blockquote>', '<p><img src="https://example.com/a.png" alt="pic"></p>', '</div>\n</body></html>']) expect(out).toContain(keep)
    expect(out).toContain('bold')
    expect(out).toContain('X</p>')
    expect(out.length).toBeLessThan(HTML.length + 60)
  })
})

describe('Markdown escapes', () => {
  const SRC = 'Not \\*emphasis\\* here, a back\\\\slash, \\_under\\_, \\# hash, \\[x\\](y), \\`tick\\`, \\~\\~no\\~\\~ and \\<tag\\>.\n\n\\# not a heading\n\n1\\. not a list\n'

  async function pass(md: string): Promise<string> {
    return docxToMarkdown(await blocksOf(md, 'a.md'))
  }

  test('md → doc → md → doc … reaches a fixed point after the first pass', async () => {
    const one = await pass(SRC)
    const two = await pass(one)
    const three = await pass(two)
    const four = await pass(three)
    expect(two).toBe(one)
    expect(three).toBe(one)
    expect(four).toBe(one)
    // the text is what the file says, once
    expect(one.length).toBeLessThan(SRC.length + 10)
  })

  test('an escape reads as the character it escapes', async () => {
    const blocks = await blocksOf('\\*not emphasis\\* \\\\ \\_ \\#', 'a.md')
    const text = blocks.flatMap((b) => (b.t === 'p' ? b.runs.map((r) => (r.t === 'text' ? r.text : '')) : [])).join('')
    expect(text).toBe('*not emphasis* \\ _ #')
  })
})
