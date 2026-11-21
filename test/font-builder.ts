/**
 * Writes a tiny TrueType font for tests: printable ASCII as one box glyph each, every one of them `advance` units wide
 * (2 em by default: text in it is about four times as wide as in an ordinary font), the space `space` units wide.
 * Valid as browsers check fonts (OpenType Sanitizer): head, hhea, maxp, OS/2, hmtx, cmap (format 4), loca, glyf,
 * name, post, with their checksums. A few kilobytes, made when a test runs: no font file in the repository.
 */

export interface TestFontOptions {
  family: string
  /** advance width of the printable characters, in units of 1000 per em */
  advance?: number
  /** advance width of the space */
  space?: number
}

const UPEM = 1000
const ASCENT = 800
const DESCENT = 200
const FIRST = 0x20
const LAST = 0x7e

class Out {
  b: number[] = []
  u8(v: number) {
    this.b.push(v & 0xff)
    return this
  }
  u16(v: number) {
    return this.u8(v >> 8).u8(v)
  }
  i16(v: number) {
    return this.u16(v < 0 ? v + 0x10000 : v)
  }
  u32(v: number) {
    return this.u16((v >>> 16) & 0xffff).u16(v & 0xffff)
  }
  tag(s: string) {
    for (const c of s) this.u8(c.charCodeAt(0))
    return this
  }
  zeros(n: number) {
    for (let i = 0; i < n; i++) this.u8(0)
    return this
  }
}

const table = (fill: (o: Out) => void) => {
  const o = new Out()
  fill(o)
  return o.b
}

function checksum(bytes: number[], from = 0, length = bytes.length): number {
  let sum = 0
  for (let i = 0; i < length; i += 4) {
    const at = from + i
    sum = (sum + (((bytes[at] ?? 0) << 24) | ((bytes[at + 1] ?? 0) << 16) | ((bytes[at + 2] ?? 0) << 8) | (bytes[at + 3] ?? 0))) >>> 0
  }
  return sum
}

/** the font's bytes */
export function testFont({ family, advance = 2000, space = 500 }: TestFontOptions): Uint8Array<ArrayBuffer> {
  // glyph 0 .notdef (a box), 1 space (empty), then one box per printable character after the space
  const count = LAST - FIRST + 2
  const box = table((o) => {
    o.i16(1).i16(50).i16(0).i16(advance - 50).i16(ASCENT - 100)
    o.u16(3).u16(0)
    for (let i = 0; i < 4; i++) o.u8(1)
    for (const dx of [50, 0, advance - 100, 0]) o.i16(dx)
    for (const dy of [0, ASCENT - 100, 0, -(ASCENT - 100)]) o.i16(dy)
    o.zeros(2)
  })
  const glyphs = [box, [], ...Array.from({ length: count - 2 }, () => box)]
  const glyf = glyphs.flat()
  const loca = table((o) => {
    let at = 0
    o.u16(0)
    for (const g of glyphs) o.u16((at += g.length) / 2)
  })
  const head = table((o) => {
    o.u16(1).u16(0).u32(0x00010000).u32(0).u32(0x5f0f3cf5).u16(0x0003).u16(UPEM)
    o.zeros(16)
    o.i16(50).i16(0).i16(advance - 50).i16(ASCENT - 100)
    o.u16(0).u16(8).i16(2).i16(0).i16(0)
  })
  const hhea = table((o) => {
    o.u16(1).u16(0).i16(ASCENT).i16(-DESCENT).i16(0).u16(Math.max(advance, space))
    o.i16(50).i16(50).i16(advance - 50).i16(1).i16(0).i16(0)
    o.zeros(8).i16(0).u16(count)
  })
  const maxp = table((o) => {
    o.u32(0x00010000).u16(count).u16(4).u16(1).u16(0).u16(0).u16(2)
    o.zeros(16)
  })
  const os2 = table((o) => {
    o.u16(4).i16(advance).u16(400).u16(5).u16(0)
    o.i16(650).i16(600).i16(0).i16(75).i16(650).i16(600).i16(0).i16(350).i16(50).i16(300).i16(0)
    o.zeros(10)
    o.u32(1).u32(0).u32(0).u32(0)
    o.tag('NONE').u16(0x0040).u16(FIRST).u16(LAST)
    o.i16(ASCENT).i16(-DESCENT).i16(0).u16(ASCENT).u16(DESCENT)
    o.u32(1).u32(0)
    o.i16(500).i16(ASCENT - 100).u16(0).u16(0x20).u16(0)
  })
  const hmtx = table((o) => {
    o.u16(advance).i16(50)
    o.u16(space).i16(0)
    for (let i = 2; i < count; i++) o.u16(advance).i16(50)
  })
  // format 4: the space and the printable characters as one segment (glyph = code - 0x1f), then the closing one
  const cmap = table((o) => {
    o.u16(0).u16(1).u16(3).u16(1).u32(12)
    o.u16(4).u16(32).u16(0).u16(4).u16(4).u16(1).u16(0)
    o.u16(LAST).u16(0xffff).u16(0)
    o.u16(FIRST).u16(0xffff)
    o.i16(1 - FIRST).i16(1)
    o.u16(0).u16(0)
  })
  const ps = family.replace(/[^A-Za-z0-9-]/g, '') || 'TestFont'
  const names: [number, string][] = [
    [1, family],
    [2, 'Regular'],
    [3, `${ps}-test`],
    [4, family],
    [5, 'Version 1.000'],
    [6, ps],
  ]
  const name = table((o) => {
    o.u16(0).u16(names.length).u16(6 + 12 * names.length)
    let at = 0
    for (const [id, text] of names) {
      o.u16(3).u16(1).u16(0x0409).u16(id).u16(text.length * 2).u16(at)
      at += text.length * 2
    }
    for (const [, text] of names) for (const c of text) o.u16(c.charCodeAt(0))
  })
  const post = table((o) => {
    o.u32(0x00030000).u32(0).i16(-100).i16(50).u32(0)
    o.zeros(16)
  })

  const tables: [string, number[]][] = [
    ['OS/2', os2],
    ['cmap', cmap],
    ['glyf', glyf],
    ['head', head],
    ['hhea', hhea],
    ['hmtx', hmtx],
    ['loca', loca],
    ['maxp', maxp],
    ['name', name],
    ['post', post],
  ]
  const n = tables.length
  const sel = Math.floor(Math.log2(n))
  const out = new Out()
  out.u32(0x00010000).u16(n).u16(16 << sel).u16(sel).u16(n * 16 - (16 << sel))
  let offset = 12 + 16 * n
  const placed: { at: number; bytes: number[] }[] = []
  for (const [tag, bytes] of tables) {
    out.tag(tag).u32(checksum(bytes)).u32(offset).u32(bytes.length)
    placed.push({ at: offset, bytes })
    offset += (bytes.length + 3) & ~3
  }
  for (const { bytes } of placed) {
    out.b.push(...bytes)
    out.zeros(((bytes.length + 3) & ~3) - bytes.length)
  }
  // head.checkSumAdjustment: the whole file sums to 0xB1B0AFBA
  const headAt = placed[3].at + 8
  const adjust = (0xb1b0afba - checksum(out.b)) >>> 0
  for (let k = 0; k < 4; k++) out.b[headAt + k] = (adjust >>> (24 - 8 * k)) & 0xff
  return Uint8Array.from(out.b)
}
