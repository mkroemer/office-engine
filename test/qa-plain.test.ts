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

  test('html: a page of a site (links in a nav, text colored by its container) keeps everything but the paragraph typed in', async () => {
    const page = `<!DOCTYPE html>
<html lang="en">
<head><meta charset="utf-8"><title>Site</title><script>window.tracked = true</script></head>
<body>
<header><nav><a href="/">Home</a> | <a href="/about">About</a></nav></header>
<main>
<div style="color:#999999"> 1 / 5</div>
<p>Read the <a href="/guide">vacation guide </a> before you <em>go</em>.</p>
<a href="/logo" title="a link around a picture that is not shown">
  <img src="/logo.png" alt="">
</a>
<table border="1"><caption>Sales</caption><tr><td>North</td><td>1,280</td></tr></table>
<blockquote><p>Quoted.</p></blockquote>
<pre><code>def hello():
    pass
</code></pre>
<dl><dt>Term</dt><dd>Definition</dd></dl>
<hr>
<p>Edit me</p>
</main>
<footer><p>&copy; 2024 Example</p></footer>
</body>
</html>
`
    const blocks = await blocksOf(page, 'site.html')
    // nothing edited: the file as it was
    expect(saveHtml(page, blocks)).toBe(page)
    typeInto(blocks, 'Edit me')
    const out = saveHtml(page, blocks)
    const lines = new Set(out.split('\n'))
    expect(page.split('\n').filter((l) => !lines.has(l))).toEqual(['<p>Edit me</p>'])
    expect(out).toContain('Edit meX</p>')
    // typing in the nav rewrites its line of links only
    const again = await blocksOf(page, 'site.html')
    typeInto(again, 'Home | About')
    const nav = saveHtml(page, again)
    expect(nav).toContain('<header><nav>')
    expect(nav).toContain('AboutX</a>')
    for (const keep of ['<head><meta charset="utf-8"><title>Site</title><script>window.tracked = true</script></head>', '<table border="1"><caption>Sales</caption><tr><td>North</td><td>1,280</td></tr></table>', '<blockquote><p>Quoted.</p></blockquote>', '<dl><dt>Term</dt><dd>Definition</dd></dl>', '<hr>', '<footer><p>&copy; 2024 Example</p></footer>']) expect(nav).toContain(keep)
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

describe('QA 2026-11-22 (Markdown): rules, underscores, bold links, alerts', () => {
  const textOf = (blocks: Block[]) => blocks.flatMap((b) => (b.t === 'p' ? [b.runs.map((r) => (r.t === 'text' ? r.text : '')).join('')] : []))

  test('a thematic break survives an edit of the paragraph before it or after it, whatever its spelling', async () => {
    for (const rule of ['---', '***', '___', '- - -', ' * * *']) {
      const src = `First para\n\n${rule}\n\nSecond para\n\n${rule}\n\nThird para\n`
      expect(saveMd(src, await blocksOf(src, 'a.md'))).toBe(src)
      for (const edited of ['First', 'Second', 'Third']) {
        const blocks = await blocksOf(src, 'a.md')
        typeInto(blocks, edited)
        expect(saveMd(src, blocks)).toBe(src.replace(`${edited} para`, `${edited} paraX`))
      }
      // a paragraph deleted next to a rule: the rule stays
      const blocks = await blocksOf(src, 'a.md')
      blocks.splice(1, 1)
      expect(saveMd(src, blocks).split('\n').filter((l) => l === rule).length).toBe(2)
    }
  })

  test('a rule with no blank line around it survives an edit next to it', async () => {
    const src = 'First para\n***\nSecond para\n'
    const blocks = await blocksOf(src, 'a.md')
    typeInto(blocks, 'Second')
    expect(saveMd(src, blocks)).toContain('***')
  })

  test('underscores at the edges of a word are emphasis, a lone one inside a word is text', async () => {
    const blocks = await blocksOf('snake_case_name, __bold__, a **b** c and _it_ here', 'a.md')
    const p = blocks[0] as Paragraph
    const look = (k: 'b' | 'i') => p.runs.filter((r) => r.t === 'text' && r.s[k]).map((r) => (r as { text: string }).text)
    expect(textOf(blocks)[0]).toBe('snake_case_name, bold, a b c and it here')
    expect(look('b')).toEqual(['bold', 'b'])
    expect(look('i')).toEqual(['it'])
  })

  // The importer is `mdInline` in src/core/convert.ts (outside the files this fix was allowed to touch): its `__` pass
  // has no word-boundary rule. Once `spans(s, ['**', '__'] …)` there tells `__` from `**` (`__` opens only after a
  // non-word character and closes only before one), this test passes and `.failing` must go.
  test('double underscores inside a word are text', async () => {
    const blocks = await blocksOf('snake__case__name, a__b and __bold__', 'a.md')
    const p = blocks[0] as Paragraph
    expect(textOf(blocks)[0]).toBe('snake__case__name, a__b and bold')
    expect(p.runs.filter((r) => r.t === 'text' && r.s.b).map((r) => (r as { text: string }).text)).toEqual(['bold'])
    // letters of any script are part of the word; a mark at the edge of one still marks
    const vi = (await blocksOf('số_lượng_ớn và tên__biến__ấy, _nghiêng_ và __đậm__', 'a.md'))[0] as Paragraph
    expect(textOf([vi])[0]).toBe('số_lượng_ớn và tên__biến__ấy, nghiêng và đậm')
    expect(vi.runs.filter((r) => r.t === 'text' && (r.s.b || r.s.i)).map((r) => (r as { text: string }).text)).toEqual(['nghiêng', 'đậm'])
  })

  test('snake_case written back stays as typed', async () => {
    const src = 'The snake_case_name and a plain line\n\nTail\n'
    const blocks = await blocksOf(src, 'a.md')
    typeInto(blocks, 'snake_case_name')
    expect(saveMd(src, blocks)).toBe('The snake_case_name and a plain lineX\n\nTail\n')
  })

  test('bold around a link: written with underscores or stars it reads and saves the same', async () => {
    const stars = await blocksOf('**bold [link](https://example.com)** end\n', 'a.md')
    const under = await blocksOf('__bold [link](https://example.com)__ end\n', 'a.md')
    expect(JSON.stringify(under)).toBe(JSON.stringify(stars))
    typeInto(under, 'end')
    const out = saveMd('__bold [link](https://example.com)__ end\n', under)
    // the link keeps its bold
    expect(out).toBe('**bold** [**link**](https://example.com) endX\n')
    const again = await blocksOf(out, 'a.md')
    const link = (again[0] as Paragraph).runs.find((r) => r.t === 'text' && r.href)!
    expect((link as { s: { b?: boolean } }).s.b).toBe(true)
  })

  const ALERT = 'Before\n\n> [!IMPORTANT]\n> Crucial info\n> second line\n\nAfter\n'

  test('a GitHub alert: unedited it is byte-identical, an edit elsewhere leaves it as it is', async () => {
    expect(saveMd(ALERT, await blocksOf(ALERT, 'a.md'))).toBe(ALERT)
    for (const edited of ['Before', 'After']) {
      const blocks = await blocksOf(ALERT, 'a.md')
      typeInto(blocks, edited)
      expect(saveMd(ALERT, blocks)).toBe(ALERT.replace(edited, `${edited}X`))
    }
  })

  test('a GitHub alert edited inside keeps its marker line and its > prefixes', async () => {
    for (const kind of ['NOTE', 'TIP', 'IMPORTANT', 'WARNING', 'CAUTION']) {
      const src = ALERT.replace('IMPORTANT', kind)
      for (const [edited, want] of [['Crucial', '> Crucial infoX\n> second line'], ['second line', '> Crucial info\n> second lineX']]) {
        const blocks = await blocksOf(src, 'a.md')
        typeInto(blocks, edited)
        expect(saveMd(src, blocks)).toBe(`Before\n\n> [!${kind}]\n${want}\n\nAfter\n`)
      }
    }
    // a quote with a blank `>` line and a CRLF file
    const src = 'Intro\r\n\r\n> [!NOTE]\r\n> one\r\n>\r\n> two\r\n\r\nOutro\r\n'
    const blocks = await blocksOf(src, 'a.md')
    typeInto(blocks, 'two')
    expect(saveMd(src, blocks)).toBe(src.replace('> two', '> twoX'))
  })
})
