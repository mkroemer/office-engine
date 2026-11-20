/**
 * @p10node/markdown: Markdown → safe HTML (CommonMark blocks, GitHub tables, task lists, strikethrough, bare links),
 * with ```mermaid / ```plantuml / ```dot code blocks drawn as SVG and, on request, other code blocks colored. No DOM
 * needed, no dependency.
 */
import { markdownToHtml } from './markdown'

export interface MarkdownOptions {
  /** Draw ```mermaid, ```plantuml / ```puml and ```dot / ```graphviz code blocks as SVG. Default true. */
  diagrams?: boolean
  /** `id` attributes on headings (from their text) so `#links` reach them. Default true. */
  headingIds?: boolean
  /**
   * Color code blocks in a known language (```js, ```python, ```yaml…): their tokens in `<span class="hl-…">`
   * (hl-keyword, hl-string, hl-comment, hl-number, hl-function, hl-type, hl-attr, hl-tag, hl-meta, hl-operator), for
   * a stylesheet of the page's own. Default false.
   */
  highlight?: boolean
}

/**
 * HTML of a Markdown text, safe to set as `innerHTML`: text is escaped, HTML written in the Markdown is shown as
 * text, links keep web / mail / relative targets only and pictures web, relative or data: images.
 */
export function render(markdown: string, options: MarkdownOptions = {}): string {
  return markdownToHtml(markdown, options)
}

/** Renders Markdown into an element (its content is replaced). */
export function renderTo(target: Element, markdown: string, options: MarkdownOptions = {}): void {
  target.innerHTML = markdownToHtml(markdown, options)
}
