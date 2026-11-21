/**
 * Code blocks of Markdown documents and source files: the language and the exact code go through the document and back;
 * the colors are only drawn (never in the document).
 */
import { describe, expect, test } from 'bun:test'
import { paraSpec } from '../src/core/docx/spec'
import type { Block, EditResult, Paragraph } from '../src/core/docx/types'
import { markdownParas, plainFormat } from '../src/core/plain'
import { mergeSource } from '../src/core/plainmerge'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { docxBodyHtml, docxToMarkdown } from '../src/react/file/text'
import { render } from '../packages/markdown/src'

const enc = (s: string) => new TextEncoder().encode(s)
const textOf = (p: Paragraph) => p.runs.map((r) => (r.t === 'text' ? r.text : '')).join('')

async function open(src: Blob | Uint8Array, name?: string, language?: string) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, { ...(name ? { name } : {}), ...(language ? { language } : {}) })
  return { engine, info, events }
}

async function blocksOf(engine: Engine, events: { event: string; data: any }[]): Promise<Block[]> {
  await engine.call('docx.start', [])
  return events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
}

const paras = (bs: Block[]) => bs.filter((b): b is Paragraph => b.t === 'p')

async function part(blob: Blob, path: string) {
  return (await ZipArchive.open(new BlobSource(blob))).text(path)
}

const saveMd = (src: string, blocks: Block[]) => mergeSource(src, 'markdown', blocks, (bs) => docxToMarkdown(bs)) ?? docxToMarkdown(blocks)

const CODE = '  const greeting = "Hello";\n\n\tconsole.log(greeting)   \n\n\n// end '
const MD = `# Notes\n\nBefore the code.\n\n\`\`\`javascript\n${CODE}\n\`\`\`\n\n\`\`\`\nno language\n\`\`\`\n\n\`\`\`mermaid\nflowchart LR\n  A --> B\n\`\`\`\n\nAfter.\n`

describe('Markdown code blocks: language and code through the document', () => {
  test('imported as one code paragraph each: its language, its exact text (spaces, tabs, blank lines)', () => {
    const ps = markdownParas(MD, false)
    const code = ps.filter((p) => p.code !== undefined)
    expect(code.map((p) => p.code)).toEqual(['javascript', ''])
    expect(code[0].runs.map((r) => r.text).join('')).toBe(CODE)
    expect(code[1].runs.map((r) => r.text).join('')).toBe('no language')
    // the diagram is drawn
    expect(ps.filter((p) => p.image)).toHaveLength(1)
  })

  test('fences: longer fences hold shorter ones, a fence with a language does not close, the fence indentation goes, an open block runs to the end', () => {
    const md = '````markdown\n```js\nx\n```\n````\n\n- item\n\n  ```py\n  def f():\n      pass\n  ```\n\n~~~sh\necho a\n```\n~~~\n\n```rust\nfn open() {}\n```rust\nstill code'
    const code = markdownParas(md, false).filter((p) => p.code !== undefined)
    expect(code.map((p) => [p.code, p.runs.map((r) => r.text).join('')])).toEqual([
      ['markdown', '```js\nx\n```'],
      ['py', 'def f():\n    pass'],
      ['sh', 'echo a\n```'],
      ['rust', 'fn open() {}\n```rust\nstill code'],
    ])
  })

  test('the document holds the language as the paragraph style, the code as it is, and no colors', async () => {
    const { engine, events, info } = await open(enc(MD), 'notes.md')
    expect(info.imported).toBe('markdown')
    const ps = paras(await blocksOf(engine, events))
    const js = ps.find((p) => p.code === 'javascript')!
    expect(textOf(js)).toBe(CODE)
    expect(js.sid).toBe('SourceCodeJavascript')
    expect(js.runs.every((r) => r.t !== 'text' || !r.d)).toBe(true)
    expect(js.ps).toBeDefined()
    expect(ps.find((p) => p.code === '')).toBeDefined()
    // opening is not an edit
    expect(await engine.call('dirty', [])).toBe(false)
    const saved = (await engine.call('save', [])) as Blob
    const styles = await part(saved, 'word/styles.xml')
    expect(styles).toContain('<w:name w:val="Source Code (javascript)"/>')
    expect(styles).toContain('<w:name w:val="Source Code"/>')
    const doc = await part(saved, 'word/document.xml')
    const p = /<w:p><w:pPr><w:pStyle w:val="SourceCodeJavascript"\/><\/w:pPr>(.*?)<\/w:p>/.exec(doc)![1]
    expect(p).not.toContain('<w:color')
    expect(p).not.toContain('<w:rPr>')
    expect(p).toContain('<w:tab/>')
    expect(p).toContain('<w:t xml:space="preserve">  const greeting = "Hello";</w:t>')
  })

  test('saved back as Markdown: unchanged, the file as it was; an edited block with its language and the rest of its code', async () => {
    const { engine, events } = await open(enc(MD), 'notes.md')
    const blocks = await blocksOf(engine, events)
    expect(saveMd(MD, blocks)).toBe(MD)
    // an edit inside the code: the block is written from the document, with its fence and language
    const k = blocks.findIndex((b) => b.t === 'p' && b.code === 'javascript')
    const js = blocks[k] as Paragraph
    const runs = js.runs.map((r) => (r.t === 'text' ? { ...r, text: r.text.replace('"Hello"', '"Hi"') } : r))
    blocks[k] = { ...js, runs, edited: true }
    const out = saveMd(MD, blocks)
    expect(out).toBe(MD.replace('"Hello"', '"Hi"'))
    // also when nothing else of the file can be kept (a whole new document)
    expect(docxToMarkdown(blocks)).toContain('```javascript\n' + CODE.replace('"Hello"', '"Hi"') + '\n```')
  })

  test('the Markdown writer: fences, languages, blank lines kept, a longer fence around backticks, consecutive lines joined', () => {
    const p = (text: string, code?: string): Paragraph => ({ t: 'p', ps: {}, runs: text ? [{ t: 'text', text, s: {} }] : [], ...(code !== undefined ? { code } : {}) })
    const md = docxToMarkdown([p('Intro'), p('a\n\n\n\nb', 'py'), p('```\ninner\n```', 'md'), p('one', 'go'), p('two', 'go'), p(''), p('', 'sh'), p('End')])
    expect(md).toBe('Intro\n\n```py\na\n\n\n\nb\n```\n\n````md\n```\ninner\n```\n````\n\n```go\none\ntwo\n```\n\n```sh\n```\n\nEnd\n')
    const html = docxBodyHtml([p('x < 1\n  y', 'js'), p('plain')])
    expect(html).toBe('<pre><code class="language-js">x &lt; 1\n  y</code></pre>\n<p>plain</p>')
  })

  test('an edit through the engine keeps the style (the language); the saved document has no colors', async () => {
    const { engine, events } = await open(enc(MD), 'notes.md')
    const blocks = await blocksOf(engine, events)
    const js = paras(blocks).find((p) => p.code === 'javascript')!
    const first = js.runs[0] as Extract<Paragraph['runs'][number], { t: 'text' }>
    const edited: Paragraph = { ...js, edited: true, runs: [{ ...first, text: 'let x = 1\nlet y = 2' }] }
    const res = (await engine.call('docx.edit', [{ anchor: js.pid, list: [paraSpec([edited])], want: true }])) as EditResult
    const back = res.blocks[js.pid!][0] as Paragraph
    expect(back.code).toBe('javascript')
    expect(textOf(back)).toBe('let x = 1\nlet y = 2')
    const saved = (await engine.call('save', [])) as Blob
    const again = await open(saved)
    const ps = paras(await blocksOf(again.engine, again.events))
    const p2 = ps.find((p) => p.code === 'javascript')!
    expect(textOf(p2)).toBe('let x = 1\nlet y = 2')
    expect(await part(saved, 'word/document.xml')).not.toContain('<w:color')
  })
})

describe('source files open as code', () => {
  test('by name, MIME type or the host’s hint; Markdown, HTML, text, SVG keep their own view', () => {
    const head = enc('const a = 1\n')
    expect(plainFormat('main.ts', '', head, true)).toBe('code')
    expect(plainFormat('script.py', 'text/x-python', head, true)).toBe('code')
    expect(plainFormat('blob', 'application/json', head, true)).toBe('code')
    expect(plainFormat(undefined, 'text/yaml', head, true)).toBe('code')
    expect(plainFormat('style.css', '', head, true)).toBe('code')
    expect(plainFormat('notes.txt', '', head, true, 'python')).toBe('code')
    expect(plainFormat('notes.txt', '', head, true)).toBe('text')
    expect(plainFormat('README.md', '', head, true)).toBe('markdown')
    expect(plainFormat('page.html', '', head, true)).toBe('html')
    expect(plainFormat('x.svg', '', enc('<svg xmlns="http://www.w3.org/2000/svg"/>'), true)).toBe('svg')
    expect(plainFormat('a.csv', '', head, true)).toBe('csv')
    // a binary file named like source code is not read as text
    expect(() => plainFormat('clip.ts', 'video/mp2t', new Uint8Array([0x47, 0x00, 0x11, 0x10]), true)).toThrow(/not a text file/)
  })

  test('the engine keeps the text and the bytes as they are, and names the language', async () => {
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, ...enc('def f():\r\n    return "é"\r\n')])
    const { engine, info, events } = await open(new Blob([bytes]), 'tool.py')
    expect(info).toMatchObject({ kind: 'docx', imported: 'code', language: 'python' })
    expect(await engine.call('plain.source', [])).toBe('def f():\r\n    return "é"\r\n')
    const original = (await engine.call('plain.original', [])) as Blob
    expect(new Uint8Array(await original.arrayBuffer())).toEqual(bytes)
    // the document it opens as: one code paragraph in its language
    const ps = paras(await blocksOf(engine, events))
    expect(ps).toHaveLength(1)
    expect(ps[0].code).toBe('python')
    expect(textOf(ps[0])).toBe('def f():\n    return "é"\n')
  })

  test('the hint wins over the name; an unknown language is kept as given (shown plain)', async () => {
    expect((await open(enc('SELECT 1'), 'query.txt', 'SQL')).info).toMatchObject({ imported: 'code', language: 'sql' })
    expect((await open(enc('x'), 'a.js', 'c++')).info.language).toBe('cpp')
    expect((await open(enc('MOVE A TO B.'), 'prog.cbl', 'COBOL')).info).toMatchObject({ imported: 'code', language: 'cobol' })
    expect((await open(enc('{"a": 1}'), undefined)).info.imported).not.toBe('code')
    expect((await open(new Blob(['{"a": 1}'], { type: 'application/json' }))).info).toMatchObject({ imported: 'code', language: 'json' })
    // Markdown with no hint stays a document
    expect((await open(enc('# T'), 'r.md')).info.imported).toBe('markdown')
    // other files have no original kept
    const md = await open(enc('# T'), 'r.md')
    expect(await md.engine.call('plain.original', [])).toBeNull()
  })
})

describe('@p10node/markdown: highlight option', () => {
  test('off by default; on: tokens as spans with hl- classes, escaped; unknown languages and diagrams unchanged', () => {
    const md = '```js\nconst a = "<b>" // c\n```'
    expect(render(md)).toBe('<pre><code class="language-js">const a = &quot;&lt;b&gt;&quot; // c\n</code></pre>')
    const html = render(md, { highlight: true })
    expect(html).toBe(
      '<pre><code class="language-js"><span class="hl-keyword">const</span> a = <span class="hl-string">&quot;&lt;b&gt;&quot;</span> <span class="hl-comment">// c</span>\n</code></pre>',
    )
    expect(render('```cobol\nMOVE A\n```', { highlight: true })).toBe('<pre><code class="language-cobol">MOVE A\n</code></pre>')
    expect(render('```\nplain\n```', { highlight: true })).toBe('<pre><code>plain\n</code></pre>')
    expect(render('```mermaid\ngraph LR\n  A --> B\n```', { highlight: true, diagrams: false })).toContain('<span class="hl-keyword">graph</span>')
  })
})
