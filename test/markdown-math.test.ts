/**
 * TeX math in Markdown (#33): which `$` are math, the formulas as equations of the document (the TeX and its
 * delimiters with them), saved back as written, in the Word file, the text / web page exports and the clipboard.
 */
import { describe, expect, test } from 'bun:test'
import { markdownToHtml } from '../src/core/convert'
import { paraSpec } from '../src/core/docx/spec'
import type { Block, EditResult, Inline, Paragraph } from '../src/core/docx/types'
import { htmlToParas } from '../src/core/html'
import { displayEnd, findMath, mathSource } from '../src/core/math/scan'
import { mergeSource } from '../src/core/plainmerge'
import { BlobSource } from '../src/core/source'
import { Engine } from '../src/core/worker/engine'
import { ZipArchive } from '../src/core/zip'
import { docxBodyHtml, docxToHtml, docxToMarkdown, docxToText } from '../src/react/file/text'
import * as markdown from '../packages/markdown/src/index'

const spans = (s: string) => findMath(s).map((m) => [m.delim, m.tex])

describe('which dollars are math', () => {
  test('inline $…$ and display $$…$$, also over lines', () => {
    expect(spans('Euler: $e^{i\\pi} + 1 = 0$.')).toEqual([['$', 'e^{i\\pi} + 1 = 0']])
    expect(spans('$$x^2$$ and $$\n\\frac{a}{b}\n$$')).toEqual([['$$', 'x^2'], ['$$', '\n\\frac{a}{b}\n']])
    expect(spans('a $x\ny$ b')).toEqual([['$', 'x\ny']])
    expect(spans('$a$$b$')).toEqual([['$', 'a'], ['$', 'b']])
    expect(findMath('x $a$ y')[0]).toMatchObject({ start: 2, end: 5 })
  })

  test('currency and other dollars are text (pandoc rules)', () => {
    expect(spans('It costs $5 and $10.')).toEqual([])
    expect(spans('From $20,000 to $30,000')).toEqual([])
    // a space after the opening or before the closing $
    expect(spans('$ x$ and $x $')).toEqual([])
    // the closing $ is followed by a digit
    expect(spans('$x$5')).toEqual([])
    // the first $ after the opening one closes it, or there is no formula there
    expect(spans('costs $5 and $10 or $x$')).toEqual([['$', 'x']])
    expect(spans('$$ $$ and $$')).toEqual([])
    expect(spans('just $')).toEqual([])
  })

  test('escaped dollars', () => {
    expect(spans('\\$x\\$')).toEqual([])
    expect(spans('\\$5 and $y$')).toEqual([['$', 'y']])
    // inside math, an escaped dollar is TeX's
    expect(spans('$a\\$b$')).toEqual([['$', 'a\\$b']])
    // an escaped backslash does not escape the dollar
    expect(spans('\\\\$x$')).toEqual([['$', 'x']])
  })

  test('code spans win', () => {
    expect(spans('`$x$`')).toEqual([])
    expect(spans('``a $x$ b`` and $y$')).toEqual([['$', 'y']])
    // the formula would end inside a code span: no formula
    expect(spans('$a `b$` c')).toEqual([])
    // an unclosed backtick is a backtick
    expect(spans('`a $x$')).toEqual([['$', 'x']])
  })

  test('the Markdown of a formula; a $$ block over lines', () => {
    expect(mathSource({ tex: 'x', delim: '$' })).toBe('$x$')
    expect(mathSource({ tex: '\n x \n', delim: '$$' })).toBe('$$\n x \n$$')
    // what $…$ cannot hold (a pasted formula)
    expect(mathSource({ tex: ' x ', delim: '$' })).toBe('$$ x $$')
    expect(mathSource({ tex: 'a$b', delim: '$' })).toBe('$$a$b$$')
    expect(mathSource({ tex: 'a\\$b', delim: '$' })).toBe('$a\\$b$')
    expect(mathSource({ tex: 'a\\', delim: '$' })).toBe('$$a\\$$')
    expect(displayEnd(['$$', '- a', '# b', '$$'], 0)).toBe(3)
    expect(displayEnd(['$$', 'a', '', '$$'], 0)).toBe(-1)
    expect(displayEnd(['$$x$$'], 0)).toBe(-1)
    expect(displayEnd(['    $$', 'a', '$$'], 0)).toBe(-1)
  })
})

describe('Markdown → HTML of the document', () => {
  const ann = (tex: string, display = false) => `<math${display ? ' display="block"' : ''}><annotation encoding="application/x-tex">${tex}</annotation></math>`

  test('paragraphs, headings, list items, table cells, quotes', () => {
    expect(markdownToHtml('A $x^2$ b.')).toBe(`<p>A ${ann('x^2')} b.</p>`)
    expect(markdownToHtml('# Title $a<b$')).toBe(`<h1>Title ${ann('a&lt;b')}</h1>`)
    expect(markdownToHtml('- item $x$')).toBe(`<ul>\n<li>item ${ann('x')}</li>\n</ul>`)
    expect(markdownToHtml('| a | b |\n| --- | --- |\n| $x$ | 2 |')).toContain(`<td>${ann('x')}</td>`)
    expect(markdownToHtml('> quote $q$')).toBe(`<blockquote><i>quote ${ann('q')}</i></blockquote>`)
    // the rest of the Markdown around it still is Markdown
    expect(markdownToHtml('**bold** $a*b*c$ *it*')).toBe(`<p><b>bold</b> ${ann('a*b*c')} <i>it</i></p>`)
  })

  test('a $$ block keeps its lines, whatever they start with', () => {
    expect(markdownToHtml('$$\n  a\n- b\n# c\n$$')).toBe(`<p>${ann('\n  a\n- b\n# c\n', true)}</p>`)
    expect(markdownToHtml('Text before\n$$\nx\n$$\nafter')).toBe(`<p>Text before ${ann('\nx\n', true)} after</p>`)
  })

  test('code blocks and indented code have no math; escaped dollars are dollars', () => {
    expect(markdownToHtml('```\n$x$\n```')).toBe('<pre><code>$x$</code></pre>')
    expect(markdownToHtml('    $x$ indented')).toBe('<p>$x$ indented</p>')
    expect(markdownToHtml('`$x$` and \\$5')).toBe('<p><code>$x$</code> and $5</p>')
    expect(markdownToHtml('It costs $5 and $10.')).toBe('<p>It costs $5 and $10.</p>')
  })

  test('MathML with a TeX annotation is a formula (the clipboard, KaTeX, Wikipedia), other MathML is its text', () => {
    const katex = '<p>Area <span class="katex"><math><semantics><mrow><mi>π</mi><msup><mi>r</mi><mn>2</mn></msup></mrow><annotation encoding="application/x-tex">\\pi r^2</annotation></semantics></math></span> m</p>'
    const [p] = htmlToParas(katex, false)
    expect(p.runs.map((r) => (r.t === 'ref' ? r.math : r.t === 'text' ? r.text : r.t))).toEqual(['Area ', { tex: '\\pi r^2', delim: '$' }, ' m'])
    const [d] = htmlToParas('<math display="block"><semantics><mi>x</mi><annotation encoding="TeX">x</annotation></semantics></math>', false)
    expect((d.runs[0] as Inline & { t: 'ref' }).math).toEqual({ tex: 'x', delim: '$$' })
    const [plain] = htmlToParas('<p><math><mi>y</mi></math></p>', false)
    expect(plain.runs).toMatchObject([{ t: 'text', text: 'y' }])
  })
})

// ---------------------------------------------------------------- the document

const enc = (s: string) => new TextEncoder().encode(s)

async function open(text: string | Blob, name = 'notes.md') {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  await engine.open(typeof text === 'string' ? enc(text) : text, { name })
  await engine.call('docx.start', [])
  const blocks: Block[] = events.filter((e) => e.event === 'docx:blocks').flatMap((e) => e.data)
  return { engine, blocks }
}

const paras = (bs: Block[]) => bs.filter((b): b is Paragraph => b.t === 'p')
const maths = (bs: Block[]) => paras(bs).flatMap((p) => p.runs.flatMap((r) => (r.t === 'ref' && r.math ? [r.math] : [])))
const saveMd = (src: string, blocks: Block[]) => mergeSource(src, 'markdown', blocks, (bs) => docxToMarkdown(bs)) ?? docxToMarkdown(blocks)

async function partText(blob: Blob, path: string) {
  return (await ZipArchive.open(new BlobSource(blob))).text(path)
}

const MD = `# Notes

Euler’s identity: $e^{i\\pi} + 1 = 0$.

$$
\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}
$$

It costs $5 and \\$10, see \`$x$\`.

$$
  \\begin{aligned}
    a &= b \\\\
    c &= d
  \\end{aligned}
$$

Broken: $\\frac{a}{$ and $\\foo x$.
`

describe('a Markdown file with math opened as a document', () => {
  test('formulas are equations holding their TeX and delimiters, byte for byte; the rest is text', async () => {
    const { engine, blocks } = await open(MD)
    expect(maths(blocks)).toEqual([
      { tex: 'e^{i\\pi} + 1 = 0', delim: '$' },
      { tex: '\n\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}\n', delim: '$$' },
      { tex: '\n  \\begin{aligned}\n    a &= b \\\\\n    c &= d\n  \\end{aligned}\n', delim: '$$' },
      { tex: '\\frac{a}{', delim: '$' },
      { tex: '\\foo x', delim: '$' },
    ])
    const ps = paras(blocks)
    expect(ps[1].runs.map((r) => r.t)).toEqual(['text', 'ref', 'text'])
    // the equation reads as its linear form (find, word count)
    expect((ps[1].runs[1] as Inline & { t: 'ref' }).text).toBe('e^iπ+1=0')
    expect(ps[3].runs.map((r) => (r.t === 'text' ? r.text : r.t)).join('')).toBe('It costs $5 and $10, see $x$.')
    // opened is not edited
    expect(await engine.call('dirty', [])).toBe(false)
    // saved unedited: the file as it was
    expect(saveMd(MD, blocks)).toBe(MD)
  })

  test('an edit next to an equation writes the paragraph again, its TeX as written', async () => {
    const { blocks } = await open(MD)
    const p = paras(blocks)[1]
    ;(p.runs[0] as { text: string }).text = 'Euler’s famous identity: '
    ;(p.runs[2] as { text: string }).text = '!'
    expect(saveMd(MD, blocks)).toBe(MD.replace('Euler’s identity: $e^{i\\pi} + 1 = 0$.', 'Euler’s famous identity: $e^{i\\pi} + 1 = 0$!'))
    // the whole file written from the document (Download as Markdown): the formulas as written too
    const whole = docxToMarkdown(blocks)
    expect(whole).toContain('$e^{i\\pi} + 1 = 0$!')
    expect(whole).toContain('$$\n\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}\n$$')
    expect(whole).toContain('$$\n  \\begin{aligned}\n    a &= b \\\\\n    c &= d\n  \\end{aligned}\n$$')
    expect(whole).toContain('Broken: $\\frac{a}{$ and $\\foo x$.')
  })

  test('an edit in the worker (as the editor sends it): the equation is copied, its TeX comes back', async () => {
    const { engine, blocks } = await open(MD)
    const p = paras(blocks)[1]
    const edited: Paragraph = { ...p, edited: true, runs: [{ ...(p.runs[0] as Inline & { t: 'text' }), text: 'Identity: ' }, p.runs[1], p.runs[2]] }
    const res = (await engine.call('docx.edit', [{ anchor: p.pid, list: [paraSpec([edited])], want: true }])) as EditResult
    const back = res.blocks[p.pid!][0] as Paragraph
    expect(back.runs.map((r) => (r.t === 'ref' ? r.math : r.t === 'text' ? r.text : r.t))).toEqual(['Identity: ', { tex: 'e^{i\\pi} + 1 = 0', delim: '$' }, '.'])
    expect(await engine.call('dirty', [])).toBe(true)
  })

  test('a pasted formula (no source run) is written from its TeX; the Word file says the TeX namespace is ignorable', async () => {
    const { engine, blocks } = await open('Plain *text* here.\n', 'plain.md')
    const p = paras(blocks)[0]
    const math: Inline = { t: 'ref', text: 'x²', s: {}, math: { tex: 'x^2 < y', delim: '$' } }
    const edited: Paragraph = { ...p, edited: true, runs: [...p.runs, math] }
    const res = (await engine.call('docx.edit', [{ anchor: p.pid, list: [paraSpec([edited])], want: true }])) as EditResult
    const back = res.blocks[p.pid!][0] as Paragraph
    expect(back.runs.at(-1)).toMatchObject({ t: 'ref', math: { tex: 'x^2 < y', delim: '$' } })
    const docx = (await engine.call('save', [])) as Blob
    const xml = await partText(docx, 'word/document.xml')
    expect(xml).toMatch(/<w:document [^>]*xmlns:oe="urn:p10node:office-engine:math"[^>]*mc:Ignorable="[^"]*\boe\b/)
    expect(xml).toContain('oe:tex="x^2 &lt; y" oe:delim="$"')
    const again = await open(docx, 'plain.docx')
    expect(maths(again.blocks)).toEqual([{ tex: 'x^2 < y', delim: '$' }])
  })

  test('the Word file: real equations (OMML) Word can show, the TeX kept for a round trip', async () => {
    const { engine } = await open(MD)
    const docx = (await engine.call('save', [])) as Blob
    const xml = await partText(docx, 'word/document.xml')
    expect(xml).toContain(' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"')
    expect(xml).toContain(' mc:Ignorable="oe"')
    expect(xml).toContain('<m:oMath oe:tex="e^{i\\pi} + 1 = 0" oe:delim="$"><m:sSup>')
    expect(xml).toContain('<m:oMathPara oe:tex="&#10;\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}&#10;" oe:delim="$$"><m:oMath><m:f>')
    // opened again as a Word file: the same formulas, saved as Markdown as written
    const again = await open(docx, 'notes.docx')
    expect(maths(again.blocks)).toEqual(maths((await open(MD)).blocks))
    expect(docxToMarkdown(again.blocks)).toContain('$$\n  \\begin{aligned}\n    a &= b \\\\\n    c &= d\n  \\end{aligned}\n$$')
  })

  test('Windows line breaks: the formula keeps them when its neighbour is edited', async () => {
    const crlf = MD.replace(/\n/g, '\r\n')
    const { blocks } = await open(crlf)
    const p = paras(blocks).find((x) => x.runs.some((r) => r.t === 'ref'))!
    ;(p.runs[2] as { text: string }).text = '!'
    const out = saveMd(crlf, blocks)
    expect(out).toBe(crlf.replace('= 0$.', '= 0$!'))
    expect(out).toContain('$$\r\n\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}\r\n$$')
  })

  test('dollar signs in edited text are escaped only when they would be math', async () => {
    const md = 'Price $5 here.\n\nTyped: x.\n'
    const { blocks } = await open(md)
    const [a, b] = paras(blocks)
    ;(a.runs[0] as { text: string }).text = 'Price $5 and $6 here.'
    ;(b.runs[0] as { text: string }).text = 'Typed: $x$ and $y$.'
    expect(saveMd(md, blocks)).toBe('Price $5 and $6 here.\n\nTyped: \\$x\\$ and \\$y\\$.\n')
    // next to a formula: escaped when they would pair with it
    const m = await open('Cost $x$ only.\n')
    const p = paras(m.blocks)[0]
    ;(p.runs[0] as { text: string }).text = 'Cost $5 '
    expect(docxToMarkdown(m.blocks)).toBe('Cost $5 $x$ only.\n')
    ;(p.runs[0] as { text: string }).text = 'Cost $b'
    expect(docxToMarkdown(m.blocks)).toBe('Cost \\$b$x$ only.\n')
    ;(p.runs[0] as { text: string }).text = 'Cost '
    ;(p.runs[2] as { text: string }).text = ' only $5.'
    expect(docxToMarkdown(m.blocks)).toBe('Cost $x$ only $5.\n')
  })

  test('plain text and web page exports; the clipboard', async () => {
    const { blocks } = await open(MD)
    const text = docxToText(blocks)
    expect(text).toContain('Euler’s identity: $e^{i\\pi} + 1 = 0$.')
    expect(text).toContain('$$\n\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}\n$$')
    const html = docxToHtml(blocks, { title: 'Notes' })
    expect(html).toContain('<math xmlns="http://www.w3.org/1998/Math/MathML"><semantics><mrow><msup><mi>e</mi>')
    expect(html).toContain('<annotation encoding="application/x-tex">e^{i\\pi} + 1 = 0</annotation>')
    expect(html).toContain('<math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><semantics><mfrac>')
    // a broken one: its source, marked
    expect(html).toContain('<merror><mtext>\\frac{a}{</mtext></merror>')
    expect(html).toContain('<merror><mtext>\\foo</mtext></merror>')
    // what the clipboard gets is read back as the same formulas
    const back = htmlToParas(docxBodyHtml(paras(blocks).slice(1, 3)), false)
    expect(back.flatMap((p) => p.runs.flatMap((r) => (r.t === 'ref' && r.math ? [r.math] : [])))).toEqual(maths(blocks).slice(0, 2))
  })

  test('an HTML file with MathML holding TeX opens with the formulas, saved back as MathML', async () => {
    const page = '<p>Area: <math><semantics><mi>A</mi><annotation encoding="application/x-tex">\\pi r^2</annotation></semantics></math>.</p>\n'
    const { blocks } = await open(page, 'page.html')
    expect(maths(blocks)).toEqual([{ tex: '\\pi r^2', delim: '$' }])
    expect(mergeSource(page, 'html', blocks, (bs) => docxBodyHtml(bs))).toBe(page)
  })
})

describe('@p10node/markdown: math', () => {
  const md = markdown.render

  test('inline and display formulas as MathML, the TeX in an annotation', () => {
    expect(md('Euler: $e^{i\\pi}=-1$.')).toBe(
      '<p>Euler: <math xmlns="http://www.w3.org/1998/Math/MathML"><semantics><mrow><msup><mi>e</mi><mrow><mi>i</mi><mi>π</mi></mrow></msup><mo>=</mo><mo>−</mo><mn>1</mn></mrow><annotation encoding="application/x-tex">e^{i\\pi}=-1</annotation></semantics></math>.</p>',
    )
    const block = md('$$\n\\frac{a}{b}\n- c\n$$')
    expect(block).toStartWith('<p><math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><semantics><mrow><mfrac><mi>a</mi><mi>b</mi></mfrac><mo>−</mo><mi>c</mi></mrow>')
    // in a list item, a heading (the id from the text, not the TeX), a table
    expect(md('- $x$')).toContain('<li><math')
    expect(md('## Area $\\pi r^2$')).toStartWith('<h2 id="area-πr2">')
    expect(md('| a |\n|---|\n| $x$ |')).toContain('<td><math')
  })

  test('currency, escapes, code are not math; math: false turns it off', () => {
    expect(md('It costs $5 and $10.')).toBe('<p>It costs $5 and $10.</p>')
    expect(md('\\$x\\$ and `$y$`')).toBe('<p>$x$ and <code>$y$</code></p>')
    expect(md('```\n$x$\n```')).toBe('<pre><code>$x$\n</code></pre>')
    expect(md('    $x$')).toBe('<pre><code>$x$\n</code></pre>')
    expect(md('$x$', { math: false })).toBe('<p>$x$</p>')
  })

  test('still safe: the TeX is text, a color is a color, broken TeX is shown as text', () => {
    const html = md('$<script>alert(1)</script>$ $\\color{red" onclick="x}{y}$ $\\href{javascript:alert(1)}{x}$ $\\frac{a}{$')
    // (no script element, no attribute that runs or loads something)
    expect(html).not.toMatch(/<script|<[^>]*\s(?:on\w+|href|src)=/i)
    expect(html).toContain('&lt;script&gt;')
    expect(html).toContain('<merror><mtext>\\href</mtext></merror>')
    expect(html).toContain('<merror><mtext>\\frac{a}{</mtext></merror>')
    // only MathML elements of our own, and only their attributes
    const tags = new Set([...html.matchAll(/<([a-z]+)/g)].map((m) => m[1]))
    for (const t of tags) expect(['p', 'math', 'semantics', 'mrow', 'mi', 'mo', 'mn', 'mtext', 'merror', 'mstyle', 'annotation', 'mfrac', 'msup']).toContain(t)
    const attrs = new Set([...html.matchAll(/<[a-z]+([^>]*)>/g)].flatMap((t) => [...t[1].matchAll(/\s([a-z-]+)="/g)].map((m) => m[1])))
    for (const a of attrs) expect(['xmlns', 'encoding', 'stretchy', 'mathcolor', 'display', 'mathvariant']).toContain(a)
  })
})
