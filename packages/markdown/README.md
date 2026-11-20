# @p10node/markdown

Markdown → HTML that is safe to show, with ` ```mermaid `, ` ```plantuml ` and ` ```dot ` code blocks drawn as SVG and, on request, other code blocks colored by their language. No DOM needed, no server, no dependency.

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
render(text, { diagrams: false, headingIds: false })
render(text, { highlight: true })                       // code blocks colored
```

| Option       | Default | Notes                                                                                |
|--------------|---------|--------------------------------------------------------------------------------------|
| `diagrams`   | `true`  | Draw ` ```mermaid `, ` ```plantuml ` / ` ```puml `, ` ```dot ` / ` ```graphviz `     |
| `headingIds` | `true`  | `id` on headings from their text (`# Getting started` → `#getting-started`)          |
| `highlight`  | `false` | Color code blocks in a known language: their tokens in `<span class="hl-…">` (below) |

## Code colors

With `highlight: true`, a code block in a known language has its tokens wrapped in spans, for a stylesheet of the page's own (the text stays escaped, the spans carry a class only). The languages: JavaScript / TypeScript / JSX / TSX, JSON, Python, CSS / SCSS / Less, HTML / XML, YAML, shell, SQL, Go, Rust, Java, C / C++ / C#, Markdown, diff, TOML / INI, Mermaid, PlantUML, Graphviz, with their usual names (` ```js `, ` ```ts `, ` ```py `, ` ```sh `, ` ```yml `, ` ```c++ `…). Another language is left plain.

```css
.hl-comment { color: #6a737d; font-style: italic }
.hl-keyword { color: #8e3bc7 }
.hl-string { color: #1a7f37 }
.hl-number { color: #0b7a8a }
.hl-function { color: #2a51a8 }
.hl-type { color: #a3236b }
.hl-attr { color: #0b57d0 }
.hl-tag { color: #b42318 }
.hl-meta { color: #8a5a00 }
.hl-operator { color: #c2410c }
```

## What is read

- CommonMark blocks: ATX and setext headings, paragraphs, block quotes (nested, lazy lines), bullet and numbered lists (nested, tight / loose, `start`), fenced and indented code, thematic breaks
- Inline: emphasis, strong, code spans, links and pictures (inline and reference), autolinks, backslash escapes, hard line breaks
- GitHub: tables (column alignment), task lists, strikethrough, bare `https://` / `www.` links
- Code blocks get `class="language-…"` (and, with `highlight`, colored tokens); diagrams `<figure class="md-diagram md-<lang>">`; a diagram that cannot be drawn shows why and its code (`<div class="md-diagram-error">`)

## Safety

The HTML can go straight into `innerHTML`: all text is escaped and **HTML written in the Markdown is shown as text**, links keep web, mail, phone and relative targets (`javascript:` and the like become plain text), and pictures load web, relative or `data:image/…` sources only.

## License

MIT
