/**
 * The host's fonts: which declared files a document's text can use (through the fallbacks `cssFont()` names), the
 * FontFace each becomes, and the page-wide set of them that viewers share (load, timeout, failure, release).
 */
import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import type { Block, DocxMeta, Paragraph } from '../src/core/docx/types'
import { cssFont, cssUrl, docxFonts, docxPageFonts, faceDescriptors, faceFamily, familyKey, fontFamilies, fontsFor, type FontSource } from '../src/core/fonts'
import { acquireFont, copyFonts, FONT_LIMITS, fontKey, loadFont, onFonts, releaseFont, type HostFace } from '../src/react/common/fonts'
import { testFont } from './font-builder'

describe('font stacks and declared fonts', () => {
  test('the families of a stack, as CSS compares them', () => {
    expect(fontFamilies(cssFont('Calibri')!)).toEqual(['calibri', 'carlito', 'sans-serif'])
    expect(fontFamilies(cssFont('Times New Roman')!)).toEqual(['times new roman', 'liberation serif', 'tinos', 'serif'])
    expect(fontFamilies(`'A, B',C ,  "Two   Words" , serif`)).toEqual(['a, b', 'c', 'two words', 'serif'])
    expect(fontFamilies('')).toEqual([])
    expect(familyKey(' "Liberation  Sans" ')).toBe('liberation sans')
    // ASCII case only, as CSS
    expect(familyKey('ÉLAN Sans')).toBe('Élan sans')
  })

  const OFL: FontSource[] = [
    { family: 'Carlito', src: '/fonts/Carlito-Regular.ttf' },
    { family: 'Carlito', src: '/fonts/Carlito-Bold.ttf', weight: 700 },
    { family: 'Caladea', src: '/fonts/Caladea-Regular.ttf' },
    { family: 'Liberation Sans', src: '/fonts/LiberationSans-Regular.ttf' },
    { family: 'Liberation Serif', src: '/fonts/LiberationSerif-Regular.ttf' },
    { family: 'Liberation Mono', src: '/fonts/LiberationMono-Regular.ttf' },
  ]
  const used = (...names: string[]) => fontsFor(OFL, names.map((n) => cssFont(n)!)).map((f) => f.src)

  test('a fallback declared serves the documents set in the original, through the font list', () => {
    expect(used('Calibri')).toEqual(['/fonts/Carlito-Regular.ttf', '/fonts/Carlito-Bold.ttf'])
    expect(used('Calibri Light')).toEqual(['/fonts/Carlito-Regular.ttf', '/fonts/Carlito-Bold.ttf'])
    expect(used('Cambria')).toEqual(['/fonts/Caladea-Regular.ttf'])
    expect(used('Arial')).toEqual(['/fonts/LiberationSans-Regular.ttf'])
    expect(used('Helvetica')).toEqual(['/fonts/LiberationSans-Regular.ttf'])
    expect(used('Times New Roman')).toEqual(['/fonts/LiberationSerif-Regular.ttf'])
    expect(used('Courier New')).toEqual(['/fonts/LiberationMono-Regular.ttf'])
    expect(used('Arial', 'Cambria')).toEqual(['/fonts/Caladea-Regular.ttf', '/fonts/LiberationSans-Regular.ttf'])
    // nothing names them: nothing to load
    expect(used('Georgia', 'Wingdings')).toEqual([])
    expect(fontsFor([], [cssFont('Calibri')!])).toEqual([])
  })

  test('a font declared under the name documents use, or the fallback name in any case', () => {
    expect(fontsFor([{ family: 'Calibri', src: 'c.woff2' }], [cssFont('Calibri')!])).toHaveLength(1)
    expect(fontsFor([{ family: 'carlito', src: 'c.woff2' }], [cssFont('CALIBRI')!])).toHaveLength(1)
    expect(fontsFor([{ family: '"Calibri"', src: 'c.woff2' }], [cssFont('Calibri')!])).toHaveLength(1)
    // another document's font: not this one
    expect(fontsFor([{ family: 'Calibri', src: 'c.woff2' }], [cssFont('Aptos')!])).toHaveLength(0)
  })

  test('descriptors: what is given, as CSS takes it', () => {
    expect(faceDescriptors({})).toEqual({})
    expect(faceDescriptors({ weight: 700, style: 'italic' })).toEqual({ weight: '700', style: 'italic' })
    expect(faceDescriptors({ weight: '100 900', unicodeRange: 'U+0000-00FF, U+0131' })).toEqual({ weight: '100 900', unicodeRange: 'U+0000-00FF, U+0131' })
    expect(faceDescriptors({ weight: 'bold', style: '' })).toEqual({ weight: 'bold' })
    expect(faceFamily(' "Liberation Sans" ')).toBe('Liberation Sans')
    expect(faceFamily('Evil"\\\nName')).toBe('EvilName')
  })

  test('a URL is quoted for url() without what would end it', () => {
    expect(cssUrl('https://cdn.example/f.woff2')).toBe('url("https://cdn.example/f.woff2")')
    expect(cssUrl('a") , url("b\\c\nd')).toBe('url("a\\22 ) , url(\\22 b\\5c c\\a d")')
  })

  test('the declarations of one face share a key; another file, weight or style does not', () => {
    const bytes = new Uint8Array(4)
    expect(fontKey({ family: 'Carlito', src: '/f.ttf' })).toBe(fontKey({ family: ' carlito ', src: '/f.ttf' }))
    expect(fontKey({ family: 'Carlito', src: '/f.ttf' })).not.toBe(fontKey({ family: 'Carlito', src: '/g.ttf' }))
    expect(fontKey({ family: 'Carlito', src: '/f.ttf' })).not.toBe(fontKey({ family: 'Carlito', src: '/f.ttf', weight: 700 }))
    expect(fontKey({ family: 'Carlito', src: '/f.ttf', style: 'italic' })).not.toBe(fontKey({ family: 'Carlito', src: '/f.ttf' }))
    expect(fontKey({ family: 'Carlito', src: bytes })).toBe(fontKey({ family: 'Carlito', src: bytes }))
    expect(fontKey({ family: 'Carlito', src: bytes })).not.toBe(fontKey({ family: 'Carlito', src: new Uint8Array(4) }))
  })

  test('the stacks a DOCX shows: runs, markers, empty paragraphs, cells, text boxes, headers and footers', () => {
    const p = (font: string | undefined, more: Partial<Paragraph> = {}): Paragraph => ({ t: 'p', ps: {}, runs: [{ t: 'text', text: 'x', s: { font } }], ...more })
    const blocks: Block[] = [
      p('"A", sans-serif'),
      p(undefined, { marker: { text: '•', s: { font: '"Symbol", sans-serif' }, tab: true }, markFont: '"Mark", serif' }),
      { t: 'p', ps: {}, runs: [{ t: 'box', w: 10, h: 10, blocks: [p('"Boxed", serif')] }, { t: 'field', kind: 'PAGE', s: { font: '"Field", serif' } }] },
      { t: 'tbl', cols: [10], rows: [{ cells: [{ blocks: [p('"Cell", serif')], colSpan: 1, rowSpan: 1, borders: {}, pad: [0, 0, 0, 0] }] }] } as unknown as Block,
    ]
    expect([...docxFonts(blocks)].sort()).toEqual(['"A", sans-serif', '"Boxed", serif', '"Cell", serif', '"Field", serif', '"Mark", serif', '"Symbol", sans-serif'])
    const meta = { page: { w: 1, h: 1, mt: 0, mr: 0, mb: 0, ml: 0, header: 0, footer: 0 }, font: '"Body", serif', header: [p('"Head", serif')], evenFooter: [p('"Foot", serif')] } as DocxMeta
    expect([...docxPageFonts(meta)].sort()).toEqual(['"Body", serif', '"Foot", serif', '"Head", serif'])
  })
})

// ---- the page-wide set, with a FontFace and a font set that behave as a browser's

class FakeFace {
  static made: FakeFace[] = []
  status = 'unloaded'
  settle?: { ok: () => void; fail: (e: Error) => void }
  constructor(
    readonly family: string,
    readonly source: string | BufferSource,
    readonly desc: FontFaceDescriptors,
  ) {
    if (desc.weight === 'heavy') throw new SyntaxError("Failed to construct 'FontFace': Failed to set 'heavy' as a property value.")
    FakeFace.made.push(this)
  }
  load() {
    this.status = 'loading'
    return new Promise<FakeFace>((ok, fail) => {
      this.settle = {
        ok: () => {
          this.status = 'loaded'
          ok(this)
        },
        fail: (e) => {
          this.status = 'error'
          fail(e)
        },
      }
    })
  }
}

describe('the fonts of the page', () => {
  const set = new Set<FakeFace>()
  const warnings: string[] = []
  const saved = { document: globalThis.document, FontFace: (globalThis as { FontFace?: unknown }).FontFace, warn: console.warn, timeout: FONT_LIMITS.timeout }
  const tick = (ms = 0) => new Promise((r) => setTimeout(r, ms))

  beforeEach(() => {
    set.clear()
    warnings.length = 0
    FakeFace.made = []
    Object.assign(globalThis, {
      FontFace: FakeFace,
      document: { baseURI: 'https://host.example/app/', fonts: { add: (f: FakeFace) => set.add(f), delete: (f: FakeFace) => set.delete(f), addEventListener() {}, removeEventListener() {} } },
    })
    console.warn = (...a: unknown[]) => void warnings.push(a.map(String).join(' '))
  })
  afterEach(() => {
    Object.assign(globalThis, { document: saved.document, FontFace: saved.FontFace })
    console.warn = saved.warn
    FONT_LIMITS.timeout = saved.timeout
  })

  test('viewers declaring one file share one face; it leaves the page with the last of them', async () => {
    const events: [string, boolean][] = []
    const off = onFonts((f, removed) => events.push([f.state, removed]))
    const a = acquireFont({ family: 'Carlito', src: 'fonts/carlito.ttf' })
    const b = acquireFont({ family: 'carlito', src: 'fonts/carlito.ttf' })
    expect(b).toBe(a)
    expect(FakeFace.made).toHaveLength(1)
    // resolved against the page; nothing fetched yet (the face is not loaded until text needs it)
    expect(FakeFace.made[0].source).toBe('url("https://host.example/app/fonts/carlito.ttf")')
    expect(FakeFace.made[0].family).toBe('Carlito')
    expect(FakeFace.made[0].status).toBe('unloaded')
    expect([...set]).toEqual(FakeFace.made)
    releaseFont(a)
    await tick()
    expect(set.size).toBe(1)
    releaseFont(b)
    await tick()
    expect(set.size).toBe(0)
    expect(events.at(-1)).toEqual(['idle', true])
    // let go and held again at once (an effect run twice): the same face, not a second download
    const c = acquireFont({ family: 'Carlito', src: 'fonts/carlito.ttf' })
    releaseFont(c)
    const d = acquireFont({ family: 'Carlito', src: 'fonts/carlito.ttf' })
    await tick()
    expect(d).toBe(c)
    expect(set.size).toBe(1)
    releaseFont(d)
    off()
  })

  test('loaded once, however many ask', async () => {
    const f = acquireFont({ family: 'Carlito', src: '/c.ttf', weight: 700, style: 'italic' })
    expect(FakeFace.made[0].desc).toEqual({ weight: '700', style: 'italic' })
    loadFont(f)
    loadFont(f)
    expect(f.state).toBe('loading')
    await tick()
    FakeFace.made[0].settle!.ok()
    await tick()
    expect(f.state).toBe('loaded')
    loadFont(f)
    expect(f.state).toBe('loaded')
    releaseFont(f)
  })

  test('a slow font times out and arrives later; a failing one says which and why', async () => {
    FONT_LIMITS.timeout = 20
    const slow = acquireFont({ family: 'Slow', src: '/slow.ttf' })
    loadFont(slow)
    await tick(40)
    expect(slow.state).toBe('timeout')
    expect(warnings.join('\n')).toMatch(/^\[office-engine\] font "Slow" \(\/slow\.ttf\) is still loading after 0\.02 s/)
    FakeFace.made[0].settle!.ok()
    await tick()
    expect(slow.state).toBe('loaded')

    const bad = acquireFont({ family: 'Bad', src: '/missing.ttf' })
    loadFont(bad)
    await tick()
    FakeFace.made[1].settle!.fail(new Error('A network error occurred.'))
    await tick()
    expect(bad.state).toBe('error')
    expect(bad.error?.message).toBe('Font "Bad" (/missing.ttf) could not be loaded: A network error occurred.')
    expect(warnings.at(-1)).toStartWith('[office-engine] Font "Bad" (/missing.ttf) could not be loaded')
    releaseFont(slow)
    releaseFont(bad)
  })

  test('the print frame gets the loaded faces of its own, from the same file', async () => {
    const f = acquireFont({ family: 'Carlito', src: '/c.ttf', weight: 700 })
    const idle = acquireFont({ family: 'Caladea', src: '/d.ttf' })
    loadFont(f)
    await tick()
    FakeFace.made[0].settle!.ok()
    await tick()
    const added: FakeFace[] = []
    const frame = { defaultView: { FontFace: FakeFace }, fonts: { add: (x: FakeFace) => added.push(x) } } as unknown as Document
    copyFonts(frame)
    // the loaded one only, as a face of the frame's (and loading there)
    expect(added).toHaveLength(1)
    expect(added[0]).not.toBe(FakeFace.made[0])
    expect([added[0].family, added[0].source, added[0].desc, added[0].status]).toEqual(['Carlito', 'url("https://host.example/c.ttf")', { weight: '700' }, 'loading'])
    releaseFont(f)
    releaseFont(idle)
  })

  test('descriptors a FontFace refuses: a failed font, not an exception', () => {
    const f: HostFace = acquireFont({ family: 'Heavy', src: '/h.ttf', weight: 'heavy' })
    expect(f.state).toBe('error')
    expect(f.face).toBeUndefined()
    loadFont(f)
    expect(f.state).toBe('error')
    releaseFont(f)
  })

  test('bytes: a Blob is read when first needed, a buffer is used as it is', async () => {
    const bytes = testFont({ family: 'Wide Test' })
    const blob = acquireFont({ family: 'Wide Test', src: new Blob([bytes]) })
    expect(FakeFace.made).toHaveLength(0)
    loadFont(blob)
    await tick(10)
    expect(FakeFace.made).toHaveLength(1)
    expect(new Uint8Array(FakeFace.made[0].source as ArrayBuffer)).toEqual(bytes)
    FakeFace.made[0].settle!.ok()
    await tick()
    expect(blob.state).toBe('loaded')
    const view = acquireFont({ family: 'Wide Test', src: bytes, weight: 700 })
    expect(FakeFace.made[1].source).toBe(bytes)
    releaseFont(blob)
    releaseFont(view)
  })
})

describe('the test font', () => {
  test('a TrueType file with its tables in order and checksums that add up', () => {
    const f = testFont({ family: 'Wide Test' })
    const dv = new DataView(f.buffer)
    expect(dv.getUint32(0)).toBe(0x00010000)
    const n = dv.getUint16(4)
    const tags: string[] = []
    let sum = 0
    for (let i = 0; i < f.length; i += 4) sum = (sum + dv.getUint32(i)) >>> 0
    for (let i = 0; i < n; i++) {
      const at = 12 + 16 * i
      tags.push(String.fromCharCode(...f.subarray(at, at + 4)))
      const off = dv.getUint32(at + 8)
      expect(off % 4).toBe(0)
      expect(off + dv.getUint32(at + 12)).toBeLessThanOrEqual(f.length)
    }
    expect(tags).toEqual(['OS/2', 'cmap', 'glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp', 'name', 'post'])
    expect(sum).toBe(0xb1b0afba)
    expect(f.length).toBeLessThan(8000)
  })
})
