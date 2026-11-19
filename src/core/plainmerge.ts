/**
 * A Markdown / HTML file saved after it was edited as a document. The document model holds only what the importer
 * understands (paragraphs of formatted runs: no nested lists, tables, code blocks, quotes, front matter, pictures
 * inside text), so writing the file from the model alone drops the rest. Instead the source is cut into chunks, each
 * knowing the paragraphs it imported as; the edited document is aligned with the imported paragraphs, a chunk whose
 * paragraphs all came back unchanged is kept verbatim, and only what was edited is written from the model.
 */
import type { Block, Inline, Paragraph } from './docx/types'
import type { ImportPara } from './convert'
import { htmlParas, markdownParas } from './plain'

export type MergeFormat = 'markdown' | 'html'

const MONO = /courier|mono|consolas|menlo|fira code|source code/i
const ORDERED = /^[\da-zA-Z]{1,4}[.)]$/

interface Part {
  key: string
  text: string
}

/** runs of equal look as one, whitespace taking the look of what is before it; '' when there is no text */
function runsKey(parts: { key: string; text: string }[], heading: boolean): string {
  const out: Part[] = []
  for (const p of parts) {
    if (!p.text) continue
    const last = out[out.length - 1]
    if (last && (last.key === p.key || !p.text.trim())) last.text += p.text
    else if (last && !last.text.trim()) {
      last.key = p.key
      last.text += p.text
    } else out.push({ key: p.key, text: p.text })
  }
  const sig = out.map((p) => `${heading ? '' : p.key}\u0001${p.text.replace(/\s+/g, ' ')}`).join('\u0002')
  return sig.replace(/\u0001|\u0002/g, '').trim() ? sig : ''
}

const norm = (c: string | undefined) => (c ?? '').replace('#', '').toLowerCase()

/** what the paragraph reads and looks like, for telling an edited one from one that came back as imported */
function importSig(p: ImportPara): string {
  if (p.image) return 'IMG'
  const link = (r: { href?: string }) => (r.href ? 'a' + r.href : '')
  const parts = p.runs.map((r) => {
    const s = r.set ?? {}
    const href = link(r)
    const key = [s.b ? 'b' : '', s.i ? 'i' : '', s.strike ? 's' : '', !href && s.u && s.u !== 'none' ? 'u' : '', s.font && MONO.test(s.font) ? 'c' : '', href, !href && s.color ? 'k' + norm(s.color) : '', s.hl ? 'h' + norm(s.hl) : ''].join('')
    return { key, text: r.text }
  })
  const text = runsKey(parts, !!p.heading)
  return text ? `${p.heading ?? 0}|${p.list ?? ''}|${text}` : ''
}

function blockSig(b: Block): string {
  if (b.t !== 'p') return 'TABLE' + Math.random()
  const p: Paragraph = b
  const lvl = p.ps.outline
  const heading = lvl !== undefined && lvl < 6
  const list = p.marker ? (ORDERED.test(p.marker.text.trim()) ? 'number' : 'bullet') : ''
  const parts: { key: string; text: string }[] = []
  let img = false
  for (const r of p.runs as Inline[]) {
    if (r.t === 'img') {
      if (r.tc?.k !== 'del') img = true
      continue
    }
    if ((r.t !== 'text' && r.t !== 'ref') || r.tc?.k === 'del' || r.s.vanish) continue
    const href = r.t === 'text' && r.href ? 'a' + r.href : ''
    const s = r.s
    const key = [s.b ? 'b' : '', s.i ? 'i' : '', s.strike || s.dstrike ? 's' : '', !href && s.u && s.u !== 'none' ? 'u' : '', s.font && MONO.test(s.font) ? 'c' : '', href, !href && s.color && s.color !== 'auto' ? 'k' + norm(s.color) : '', s.hl && s.hl !== 'none' ? 'h' + norm(s.hl) : ''].join('')
    parts.push({ key, text: r.text })
  }
  if (img && !parts.length) return 'IMG'
  const text = runsKey(parts, heading)
  return text ? `${heading ? lvl + 1 : 0}|${list}|${text}` : ''
}

// ---------------------------------------------------------------- chunks

const lines = (s: string) => (s.match(/[^\r\n]*(?:\r\n|\r|\n)|[^\r\n]+$/g) ?? []) as string[]
const LI = /^\s*(?:[-*+]|\d+[.)])\s/
const FENCE = /^\s*(```|~~~)/

/** Markdown cut at its block boundaries (every character belongs to a chunk); `fine`: also at headings, rules and code blocks */
function markdownChunks(src: string, fine: boolean): string[] {
  const ls = lines(src)
  const chunks: string[] = []
  let cur = ''
  let curStart = ''
  const push = () => {
    if (cur) chunks.push(cur)
    cur = ''
    curStart = ''
  }
  let blankBefore = false
  for (let k = 0; k < ls.length; k++) {
    const line = ls[k]
    const blank = !line.trim()
    if (!blank && blankBefore) {
      // after blank lines: a new chunk, unless the block goes on (the same list, a quote)
      const cont = (LI.test(curStart) && (LI.test(line) || /^(?: {2,}|\t)/.test(line))) || (/^\s*>/.test(curStart) && /^\s*>/.test(line))
      if (!cont) push()
    }
    blankBefore = blank && !!cur
    const fence = FENCE.exec(line)
    if (fence) {
      // a code block is one unit, blank lines in it or not
      if (fine) push()
      let j = k + 1
      while (j < ls.length && !ls[j].trimStart().startsWith(fence[1])) j++
      if (!cur) curStart = line
      cur += ls.slice(k, j + 1).join('')
      k = j
      if (fine) {
        // blank lines after the block belong to it
        while (k + 1 < ls.length && !ls[k + 1].trim()) cur += ls[++k]
        push()
      }
      continue
    }
    const heading = /^\s{0,3}#{1,6}\s/.test(line) || /^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(line)
    if (fine && heading) push()
    if (!cur) curStart = line
    cur += line
    if (fine && heading) {
      while (k + 1 < ls.length && !ls[k + 1].trim()) cur += ls[++k]
      push()
    }
  }
  push()
  return chunks
}

/** front matter stays with the first chunk: it is one block however many blank lines it has */
function withFrontMatter(chunks: string[]): string[] {
  if (!/^﻿?---[ \t]*(\r\n|\r|\n|$)/.test(chunks[0] ?? '')) return chunks
  let acc = ''
  for (let at = 0; at < chunks.length; at++) {
    acc += chunks[at]
    if (/(^|[\r\n])(---|\.\.\.)[ \t]*(\r\n|\r|\n|$)/.test(acc.replace(/^﻿?---[ \t]*(\r\n|\r|\n)/, ''))) return [acc, ...chunks.slice(at + 1)]
  }
  return chunks
}

const BLOCKS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'TR', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'UL', 'OL', 'TABLE', 'TBODY', 'THEAD', 'DL', 'DT', 'DD', 'FIGURE', 'HR', 'HEAD', 'SCRIPT', 'STYLE', 'MAIN', 'NAV', 'ASIDE', 'FORM', 'FIELDSET', 'DETAILS', 'ADDRESS', 'CENTER', 'NOSCRIPT', 'TEMPLATE', 'BODY', 'HTML'])
const CONTAINERS = new Set(['DIV', 'SECTION', 'ARTICLE', 'MAIN', 'HEADER', 'FOOTER', 'NAV', 'ASIDE', 'CENTER', 'BODY', 'HTML'])
const VOID = new Set(['AREA', 'BASE', 'BR', 'COL', 'EMBED', 'HR', 'IMG', 'INPUT', 'LINK', 'META', 'PARAM', 'SOURCE', 'TRACK', 'WBR'])
const RAW = new Set(['SCRIPT', 'STYLE', 'TEXTAREA', 'TITLE', 'XMP', 'IFRAME', 'NOEMBED', 'NOFRAMES', 'NOSCRIPT'])
const ENDS_P = new Set(['ADDRESS', 'ARTICLE', 'ASIDE', 'BLOCKQUOTE', 'DIV', 'DL', 'FIGURE', 'FOOTER', 'FORM', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'HEADER', 'HR', 'MAIN', 'NAV', 'OL', 'P', 'PRE', 'SECTION', 'TABLE', 'UL'])

/** an HTML file cut at its top-level blocks (every character belongs to a chunk); `descend`: containers (div, section…) are cut inside */
function htmlChunks(src: string, descend: boolean): string[] {
  const n = src.length
  const cuts = new Set<number>()
  const stack: string[] = []
  const wrap: string[] = []
  // an inline run (text, <b>, <a>, <br>…) at the top level, its first character's position
  let run = -1
  let i = 0
  const closeRun = (at: number) => {
    if (run >= 0) cuts.add(at)
    run = -1
  }
  while (i < n) {
    const lt = src.indexOf('<', i)
    const textEnd = lt < 0 ? n : lt
    if (textEnd > i && !stack.length && run < 0 && /\S/.test(src.slice(i, textEnd))) {
      cuts.add(i + (src.slice(i, textEnd).length - src.slice(i, textEnd).trimStart().length))
      run = i
    }
    if (lt < 0) break
    i = lt
    if (src.startsWith('<!--', i)) {
      const e = src.indexOf('-->', i + 4)
      i = e < 0 ? n : e + 3
      continue
    }
    const next = src[i + 1]
    if (next === '!' || next === '?') {
      const e = src.indexOf('>', i)
      i = e < 0 ? n : e + 1
      continue
    }
    const end = next === '/'
    const m = /^[A-Za-z][^\s/>]*/.exec(src.slice(i + (end ? 2 : 1), i + 66))
    if (!m) {
      i++
      continue
    }
    let j = i + (end ? 2 : 1) + m[0].length
    for (let q = ''; j < n; j++) {
      const ch = src[j]
      if (q) {
        if (ch === q) q = ''
      } else if (ch === '"' || ch === "'") q = ch
      else if (ch === '>') break
    }
    const tagEnd = Math.min(j + 1, n)
    const tag = m[0].toUpperCase()
    if (end) {
      const k = stack.lastIndexOf(tag)
      if (k >= 0) {
        stack.length = k
        if (!stack.length) cuts.add(tagEnd)
      } else if (!stack.length && descend && wrap[wrap.length - 1] === tag) {
        wrap.pop()
        closeRun(i)
        cuts.add(i)
        cuts.add(tagEnd)
      } else if (!stack.length && BLOCKS.has(tag)) {
        closeRun(i)
        cuts.add(i)
        cuts.add(tagEnd)
      }
      i = tagEnd
      continue
    }
    if (stack.length && stack[stack.length - 1] === 'P' && ENDS_P.has(tag)) {
      stack.pop()
      if (!stack.length) cuts.add(i)
    }
    if (!stack.length) {
      if (BLOCKS.has(tag)) {
        closeRun(i)
        cuts.add(i)
        if (descend && CONTAINERS.has(tag)) {
          wrap.push(tag)
          cuts.add(tagEnd)
          i = tagEnd
          continue
        }
      } else if (run < 0) {
        cuts.add(i)
        run = i
      }
    }
    if (RAW.has(tag)) {
      const close = new RegExp(`</${tag}[\\s/>]`, 'ig')
      close.lastIndex = tagEnd
      const e = close.exec(src)
      i = e ? src.indexOf('>', e.index) + 1 || n : n
      if (!stack.length && BLOCKS.has(tag)) cuts.add(i)
      continue
    }
    if (!VOID.has(tag)) stack.push(tag)
    else if (!stack.length && BLOCKS.has(tag)) cuts.add(tagEnd)
    i = tagEnd
  }
  const at = [...cuts].filter((c) => c > 0 && c < n).sort((a, b) => a - b)
  const out: string[] = []
  let from = 0
  for (const c of at) {
    if (c > from) out.push(src.slice(from, c))
    from = c
  }
  if (from < n) out.push(src.slice(from))
  return out
}

// ---------------------------------------------------------------- merge

interface Chunk {
  text: string
  /** the paragraphs it imports as (signatures) */
  sigs: string[]
}

function chunked(src: string, format: MergeFormat): Chunk[] | null {
  const parse = (t: string) => (format === 'markdown' ? markdownParas(t, false) : htmlParas(t, false))
  const sigsOf = (t: string) => parse(t).map(importSig).filter(Boolean)
  const whole = sigsOf(src).join('\u0003')
  const attempts = format === 'markdown' ? [withFrontMatter(markdownChunks(src, true)), withFrontMatter(markdownChunks(src, false))] : [htmlChunks(src, true), htmlChunks(src, false)]
  for (const texts of attempts) {
    if (texts.join('') !== src) continue
    const chunks = texts.map((text) => ({ text, sigs: sigsOf(text) }))
    if (chunks.flatMap((c) => c.sigs).join('\u0003') === whole) return chunks
  }
  return null
}

/** the longest common subsequence of two lists, as for each of `b` the index in `a` it pairs with (-1: none); null when too big */
function align(a: string[], b: string[]): Int32Array | null {
  const m = new Int32Array(b.length).fill(-1)
  let lo = 0
  while (lo < a.length && lo < b.length && a[lo] === b[lo]) m[lo] = lo, lo++
  let ea = a.length
  let eb = b.length
  while (ea > lo && eb > lo && a[ea - 1] === b[eb - 1]) ea--, eb--, m[eb] = ea
  const n = ea - lo
  const k = eb - lo
  if (n * k > 4e6) return null
  // lengths of the common subsequence of the tails, a[lo + x..] and b[lo + y..]
  const w = k + 1
  const t = new Int32Array((n + 1) * w)
  for (let x = n - 1; x >= 0; x--)
    for (let y = k - 1; y >= 0; y--) t[x * w + y] = a[lo + x] === b[lo + y] ? t[(x + 1) * w + y + 1] + 1 : Math.max(t[(x + 1) * w + y], t[x * w + y + 1])
  for (let x = 0, y = 0; x < n && y < k; ) {
    if (a[lo + x] === b[lo + y]) m[lo + y++] = lo + x++
    else if (t[(x + 1) * w + y] >= t[x * w + y + 1]) x++
    else y++
  }
  return m
}

/**
 * The file's text after the document was edited: what came back unchanged as it was in `source`, the rest written by
 * `write` from the edited blocks. null when the source cannot be told apart that way (the caller then writes the
 * whole file from the model).
 */
export function mergeSource(source: string, format: MergeFormat, blocks: Block[], write: (blocks: Block[]) => string): string | null {
  const chunks = chunked(source, format)
  if (!chunks) return null
  const a: string[] = []
  const owner: number[] = []
  chunks.forEach((c, k) =>
    c.sigs.forEach((s) => {
      a.push(s)
      owner.push(k)
    }),
  )
  const units = blocks.map((b) => ({ b, sig: blockSig(b) })).filter((u) => u.sig)
  const m = align(a, units.map((u) => u.sig))
  if (!m) return null
  const matched = new Uint8Array(a.length)
  const at = new Int32Array(a.length).fill(-1)
  m.forEach((i, j) => {
    if (i >= 0) {
      matched[i] = 1
      at[i] = j
    }
  })
  // a chunk is kept when all its paragraphs came back, one after the other
  const keep = chunks.map((_, k) => {
    const mine = owner.flatMap((o, i) => (o === k ? [i] : []))
    return mine.every((i, x) => matched[i] && (x === 0 || at[i] === at[mine[x - 1]] + 1))
  })
  // chunks with no paragraph (front matter dropped, markup, comments) wait for the paragraph that follows them
  const before = new Map<number, number[]>()
  let seen = 0
  chunks.forEach((c, k) => {
    if (!c.sigs.length) before.set(seen, [...(before.get(seen) ?? []), k])
    seen += c.sigs.length
  })
  const parts: ({ v: string } | { b: Block[] })[] = []
  const done = new Set<number>()
  const glue = (i: number) => {
    for (const k of before.get(i) ?? []) parts.push({ v: chunks[k].text })
    before.delete(i)
  }
  const add = (b: Block) => {
    const last = parts[parts.length - 1]
    if (last && 'b' in last) last.b.push(b)
    else parts.push({ b: [b] })
  }
  let ai = 0
  let bj = 0
  glue(0)
  while (ai < a.length || bj < units.length) {
    if (bj < units.length && m[bj] < 0) add(units[bj++].b)
    else if (ai < a.length && !matched[ai]) glue(++ai)
    else if (ai < a.length && bj < units.length) {
      const k = owner[ai]
      if (keep[k]) {
        if (!done.has(k)) parts.push({ v: chunks[k].text })
        done.add(k)
      } else add(units[bj].b)
      glue(++ai)
      bj++
    } else break
  }
  for (const i of [...before.keys()]) glue(i)
  const eol = /\r\n/.test(source) ? '\r\n' : /\r/.test(source) && !/\n/.test(source) ? '\r' : '\n'
  let out = ''
  for (const p of parts) {
    let t = 'v' in p ? p.v : write(p.b).replace(/\r?\n/g, eol)
    if ('b' in p) {
      t = t.replace(/\s+$/, '') + eol + (format === 'markdown' ? eol : '')
      // a written block follows a kept one that ends without its line break
      if (out && !/(\r\n|\r|\n)$/.test(out)) out += eol
      if (out && format === 'markdown' && !/(\r\n|\r|\n)(\r\n|\r|\n)$/.test(out)) out += eol
    } else if (out && !/(\r\n|\r|\n)$/.test(out) && /^\S/.test(t) && format === 'markdown') out += eol + eol
    out += t
  }
  // the file ends as it did: one line break, or none
  const last = parts[parts.length - 1]
  if (last && 'b' in last) out = out.replace(/(\r\n|\r|\n)+$/, /(\r\n|\r|\n)$/.test(source) ? eol : '')
  return out
}
