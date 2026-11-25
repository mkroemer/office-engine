/**
 * The font files of a server: found in the system's font folders (and any the host names), read once for their
 * names, styles and weights, and declared by their paths. Text is then set in them when a PDF is written; a file is
 * read from the disk when text first needs it.
 */
import { readdir, readFile, stat } from 'node:fs/promises'
import { homedir } from 'node:os'
import { join, resolve } from 'node:path'
import { Font } from '../../../src/core/font'

/** a font file PDFs set text in: the engine's `FontSource`, the list `DocViewer`'s `fonts` takes */
export interface FontFile {
  /** the family a document names (Calibri), or the one its font list falls back to (Carlito, sans-serif) */
  family: string
  /** a path, or the file's bytes */
  src: string | ArrayBuffer | ArrayBufferView | Blob
  /** CSS weight: 400, 700, '100 900' */
  weight?: number | string
  /** 'normal', 'italic' */
  style?: string
  unicodeRange?: string
}

/** where Linux, macOS and Windows keep their fonts */
export function systemFontDirs(): string[] {
  const home = homedir()
  const win = process.env.WINDIR ?? 'C:\\Windows'
  return [
    '/usr/share/fonts',
    '/usr/local/share/fonts',
    join(home, '.fonts'),
    join(home, '.local/share/fonts'),
    '/System/Library/Fonts',
    '/Library/Fonts',
    join(home, 'Library/Fonts'),
    join(win, 'Fonts'),
  ]
}

const FILE = /\.(ttf|otf|ttc|woff2?)$/i
/** files looked at, at most: a font folder is not a disk */
const MAX_FILES = 5000
const MAX_DEPTH = 6

/**
 * The generic families of a CSS font list (the last name of every list the engine writes), each served by the first
 * of these that is installed, then by fonts that cover other scripts
 */
const GENERIC: Record<string, string[]> = {
  'sans-serif': ['Liberation Sans', 'Arimo', 'Arial', 'Helvetica Neue', 'Helvetica', 'DejaVu Sans', 'Noto Sans', 'Verdana'],
  serif: ['Liberation Serif', 'Tinos', 'Times New Roman', 'Times', 'DejaVu Serif', 'Noto Serif', 'Georgia'],
  monospace: ['Liberation Mono', 'Cousine', 'Courier New', 'Courier', 'DejaVu Sans Mono', 'Menlo', 'Consolas', 'Noto Sans Mono'],
}
/** fonts of the scripts the ones above do not cover, after them in every generic family */
const COVERAGE = /^(Noto Sans (CJK|JP|SC|TC|KR|Arabic|Hebrew|Devanagari|Thai|Symbols)|Noto Serif CJK|Arial Unicode|PingFang|Hiragino Sans|Microsoft YaHei|Malgun Gothic|Apple SD Gothic|Segoe UI Symbol|Apple Symbols)/i

/**
 * Every font file in `dirs` (searched into their subfolders), declared under its family (and its typographic family,
 * "Avenir Next" for "Avenir Next Demi Bold"), weight and style, by its path; then the generic families. A collection
 * (.ttc) is declared by its first font. Files that cannot be read are left out.
 */
export async function findFonts(dirs: string[]): Promise<FontFile[]> {
  const files: string[] = []
  const seen = new Set<string>()
  const walk = async (dir: string, depth: number) => {
    if (depth > MAX_DEPTH || files.length >= MAX_FILES) return
    let entries
    try {
      entries = await readdir(dir, { withFileTypes: true })
    } catch {
      return
    }
    for (const e of entries) {
      const path = join(dir, e.name)
      if (e.isDirectory()) await walk(path, depth + 1)
      else if (FILE.test(e.name) && !seen.has(path) && files.length < MAX_FILES) {
        seen.add(path)
        files.push(path)
      }
    }
  }
  for (const d of dirs) {
    const dir = resolve(d)
    const st = await stat(dir).catch(() => null)
    if (st?.isDirectory()) await walk(dir, 0)
    else if (st?.isFile() && FILE.test(dir) && !seen.has(dir)) (seen.add(dir), files.push(dir))
  }
  const out: FontFile[] = []
  for (const path of files) {
    try {
      const font = await Font.open(await readFile(path))
      const face = { src: path, weight: font.weight, style: font.italic ? 'italic' : 'normal' }
      for (const family of new Set([font.names.family, font.names.typoFamily].filter(Boolean) as string[])) out.push({ family, ...face })
    } catch {
      // (not a font, damaged, or of a kind that is not read)
    }
  }
  return withGenerics(out)
}

/** `fonts` and, after them, the generic families served by the best of them */
export function withGenerics(fonts: FontFile[]): FontFile[] {
  const byFamily = new Map<string, FontFile[]>()
  for (const f of fonts) {
    const key = f.family.toLowerCase()
    const list = byFamily.get(key)
    if (list) list.push(f)
    else byFamily.set(key, [f])
  }
  const coverage = [...byFamily.keys()].filter((k) => COVERAGE.test(k))
  const out = [...fonts]
  for (const [generic, names] of Object.entries(GENERIC)) {
    if (byFamily.has(generic)) continue
    const first = names.find((n) => byFamily.has(n.toLowerCase()))
    for (const key of [...(first ? [first.toLowerCase()] : []), ...coverage]) for (const f of byFamily.get(key)!) out.push({ ...f, family: generic })
  }
  return out
}

let system: Promise<FontFile[]> | undefined

/** the system's fonts, found once */
export function systemFonts(): Promise<FontFile[]> {
  return (system ??= findFonts(systemFontDirs()))
}

/** bytes of font files read lately, by path, up to this many bytes: a server reads the same few for every file */
const CACHE_BYTES = 256 * 1024 * 1024
const cache = new Map<string, Promise<Uint8Array>>()
const sizes = new Map<string, number>()
let cached = 0

/** a font file by its path, from a cache of the ones read lately */
export function readFontFile(path: string): Promise<Uint8Array> {
  const hit = cache.get(path)
  if (hit) {
    // (the most recently used last)
    cache.delete(path)
    cache.set(path, hit)
    return hit
  }
  const p = readFile(path).then((b) => {
    const bytes = new Uint8Array(b.buffer, b.byteOffset, b.byteLength)
    sizes.set(path, bytes.byteLength)
    cached += bytes.byteLength
    for (const [k] of cache) {
      if (cached <= CACHE_BYTES || k === path) break
      cache.delete(k)
      cached -= sizes.get(k) ?? 0
      sizes.delete(k)
    }
    return bytes
  })
  p.catch(() => cache.delete(path))
  cache.set(path, p)
  return p
}
