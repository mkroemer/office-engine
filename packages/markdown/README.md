# @p10node/markdown

Markdown → HTML that is safe to show, with ` ```mermaid `, ` ```plantuml ` and ` ```dot ` code blocks drawn as SVG and `$…$` / `$$…$$` TeX math as MathML. No DOM needed, no server, no dependency.

Part of [office-engine](https://github.com/p10node/office-engine). Diagrams alone: [`@p10node/diagram`](../diagram).

## Install

```sh
npm i @p10node/markdown
```

## Use

```ts
import { render, renderTo } from '@p10node/markdown'

const html = render(markdownText)                       // string
renderTo(document.getElementById('doc')!, markdownText) // into an element
render(text, { diagrams: false, headingIds: false, math: false })
```

| Option       | Default | Notes                                                                            |
|--------------|---------|----------------------------------------------------------------------------------|
| `diagrams`   | `true`  | Draw ` ```mermaid `, ` ```plantuml ` / ` ```puml `, ` ```dot ` / ` ```graphviz ` |
| `headingIds` | `true`  | `id` on headings from their text (`# Getting started` → `#getting-started`)      |
| `math`       | `true`  | `$…$` and `$$…$$` TeX formulas as MathML                                         |

## What is read

- CommonMark blocks: ATX and setext headings, paragraphs, block quotes (nested, lazy lines), bullet and numbered lists (nested, tight / loose, `start`), fenced and indented code, thematic breaks
- Inline: emphasis, strong, code spans, links and pictures (inline and reference), autolinks, backslash escapes, hard line breaks
- GitHub: tables (column alignment), task lists, strikethrough, bare `https://` / `www.` links
- Code blocks get `class="language-…"`; diagrams `<figure class="md-diagram md-<lang>">`; a diagram that cannot be drawn shows why and its code (`<div class="md-diagram-error">`)
- TeX math, read as pandoc reads it: `$…$` inline (a non-space after the opening `$`, a non-space before the closing one and no digit after it, so `$5 and $10` stays text), `$$…$$` display (over several lines too). `\$` is a dollar sign; nothing in code is math. Each formula becomes `<math>` (`display="block"` for `$$`) with the drawing in MathML and the TeX in `<annotation encoding="application/x-tex">`, which the browser draws by itself (no script, no font to load). What is read: Greek letters and symbols, operators, relations, arrows, scripts and primes, big operators with their limits, fractions, binomials, roots, `\left … \right`, accents, `\text`, the font commands (`\mathbb`, `\mathbf`, `\mathcal`…), spaces, `\color`, `\operatorname`, the named functions, and the `matrix` / `pmatrix` / `bmatrix` / `cases` / `aligned` / `gathered` / `array` environments. A command it does not know is shown as written in an `<merror>`, the rest drawn; a formula it cannot read (an unclosed brace…) is its source in an `<merror>`

Chromium follows MathML Core, which leaves out the column alignment of tables (`aligned`, `cases`, `array`): add these rules to your page to have it there too.

```css
mtd[columnalign=left] { text-align: left }
mtd[columnalign=right] { text-align: right }
mtable[columnspacing="0em 2em"] > mtr > mtd { padding-left: 0; padding-right: 0 }
mtable[columnspacing="0em 2em"] > mtr > mtd[columnalign=right]:not(:first-child) { padding-left: 2em }
```

## Safety

The HTML can go straight into `innerHTML`: all text is escaped and **HTML written in the Markdown is shown as text**, links keep web, mail, phone and relative targets (`javascript:` and the like become plain text), and pictures load web, relative or `data:image/…` sources only. TeX is text too: the MathML has only MathML elements, with no link and no style, and a `\color` is kept only when it is a color name or `#rgb` value.

## License

MIT
