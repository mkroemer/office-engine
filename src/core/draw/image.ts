/**
 * Pictures read for a backend that draws pixels it is given (a PDF): a JPEG's size and color space, PNG and GIF
 * unpacked to 8-bit color and alpha. No DOM, no dependency: inflate is the runtime's `DecompressionStream`.
 */
import { sniffImage } from '../legacy/odf/odp-media'

export { sniffImage }

/** pixels past which a picture is not unpacked: 1 GB of RGBA would be too much for one page */
export const MAX_PIXELS = 64 * 1024 * 1024

export class ImageError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'ImageError'
  }
}

const be16 = (b: Uint8Array, i: number) => (b[i] << 8) | b[i + 1]
const be32 = (b: Uint8Array, i: number) => ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]) >>> 0
const le16 = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8)

export interface JpegInfo {
  w: number
  h: number
  /** 1 gray, 3 YCbCr / RGB, 4 CMYK */
  components: number
  /** Adobe's CMYK is stored inverted */
  adobe: boolean
}

/** a JPEG's size and color components, from its frame header; null when it has none */
export function jpegInfo(b: Uint8Array): JpegInfo | null {
  let adobe = false
  for (let i = 2, steps = 0; i + 9 < b.length && steps < 10000; steps++) {
    if (b[i] !== 0xff) {
      i++
      continue
    }
    const m = b[i + 1]
    if (m === 0xff) {
      i++
      continue
    }
    if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) {
      i += 2
      continue
    }
    const len = be16(b, i + 2)
    if (m === 0xee && String.fromCharCode(b[i + 4], b[i + 5], b[i + 6], b[i + 7], b[i + 8]) === 'Adobe') adobe = true
    if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
      const h = be16(b, i + 5)
      const w = be16(b, i + 7)
      const components = b[i + 9]
      if (!w || !h || ![1, 3, 4].includes(components)) return null
      return { w, h, components, adobe }
    }
    i += 2 + len
  }
  return null
}

/** pixels as a PDF takes them: 8-bit gray or RGB, and an 8-bit alpha when some are not opaque */
export interface Pixels {
  w: number
  h: number
  /** 1 gray, 3 RGB */
  channels: 1 | 3
  color: Uint8Array
  alpha?: Uint8Array
}

async function inflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new DecompressionStream('deflate'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

/** a PNG unpacked: every bit depth and color type, palettes and their transparency, interlaced or not */
export async function decodePng(b: Uint8Array): Promise<Pixels> {
  if (b.length < 33 || be32(b, 12) !== 0x49484452) throw new ImageError('Not a PNG file')
  const w = be32(b, 16)
  const h = be32(b, 20)
  const depth = b[24]
  const type = b[25]
  const interlaced = b[28] === 1
  if (!w || !h || w * h > MAX_PIXELS) throw new ImageError(`A PNG of ${w} × ${h} pixels is not drawn`)
  const samples = ({ 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 } as Record<number, number>)[type]
  if (!samples || ![1, 2, 4, 8, 16].includes(depth)) throw new ImageError('A PNG of a kind that is not read')
  let palette: Uint8Array | undefined
  let trns: Uint8Array | undefined
  const idat: Uint8Array[] = []
  for (let i = 8; i + 12 <= b.length; ) {
    const len = be32(b, i)
    const t = be32(b, i + 4)
    if (i + 12 + len > b.length) break
    const body = b.subarray(i + 8, i + 8 + len)
    if (t === 0x504c5445) palette = body
    else if (t === 0x74524e53) trns = body
    else if (t === 0x49444154) idat.push(body)
    else if (t === 0x49454e44) break
    i += 12 + len
  }
  let size = 0
  for (const c of idat) size += c.length
  const all = new Uint8Array(size)
  size = 0
  for (const c of idat) all.set(c, size), (size += c.length)
  const raw = await inflate(all)
  const bpp = Math.max(1, (samples * depth) >> 3)
  // the passes of Adam7: x0, y0, dx, dy
  const passes = interlaced ? [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]] : [[0, 0, 1, 1]]
  const gray = type === 0 || type === 4
  const channels = gray ? 1 : 3
  const color = new Uint8Array(w * h * channels)
  const alpha = new Uint8Array(w * h).fill(255)
  let hasAlpha = false
  const maxv = (1 << depth) - 1
  let at = 0
  for (const [x0, y0, dx, dy] of passes) {
    const pw = Math.ceil((w - x0) / dx)
    const ph = Math.ceil((h - y0) / dy)
    if (pw <= 0 || ph <= 0) continue
    const stride = Math.ceil((pw * samples * depth) / 8)
    let prev = new Uint8Array(stride)
    for (let row = 0; row < ph; row++) {
      if (at + 1 + stride > raw.length) throw new ImageError('A PNG whose pixels end early')
      const filter = raw[at]
      const line = raw.slice(at + 1, at + 1 + stride)
      at += 1 + stride
      unfilter(filter, line, prev, bpp)
      prev = line
      const y = y0 + row * dy
      for (let col = 0; col < pw; col++) {
        const x = x0 + col * dx
        const p = y * w + x
        const sample = (k: number): number => {
          const n = col * samples + k
          if (depth === 8) return line[n]
          if (depth === 16) return line[n * 2]
          const bit = n * depth
          return (line[bit >> 3] >> (8 - depth - (bit & 7))) & maxv
        }
        const full = (v: number) => (depth >= 8 ? v : Math.round((v * 255) / maxv))
        if (type === 3) {
          const k = sample(0)
          color[p * 3] = palette?.[k * 3] ?? 0
          color[p * 3 + 1] = palette?.[k * 3 + 1] ?? 0
          color[p * 3 + 2] = palette?.[k * 3 + 2] ?? 0
          if (trns && k < trns.length && trns[k] !== 255) (alpha[p] = trns[k]), (hasAlpha = true)
        } else if (gray) {
          const v = sample(0)
          color[p] = full(v)
          if (type === 4) {
            alpha[p] = full(sample(1))
            if (alpha[p] !== 255) hasAlpha = true
          } else if (trns && trns.length >= 2) {
            // (a 16-bit key is compared on its high byte: what was sampled)
            const key = depth === 16 ? trns[0] : be16(trns, 0)
            if (v === key) (alpha[p] = 0), (hasAlpha = true)
          }
        } else {
          const r = sample(0)
          const g = sample(1)
          const bl = sample(2)
          color[p * 3] = r
          color[p * 3 + 1] = g
          color[p * 3 + 2] = bl
          if (type === 6) {
            alpha[p] = sample(3)
            if (alpha[p] !== 255) hasAlpha = true
          } else if (trns && trns.length >= 6) {
            const k = depth === 16 ? [trns[0], trns[2], trns[4]] : [be16(trns, 0), be16(trns, 2), be16(trns, 4)]
            if (r === k[0] && g === k[1] && bl === k[2]) (alpha[p] = 0), (hasAlpha = true)
          }
        }
      }
    }
  }
  return { w, h, channels, color, ...(hasAlpha ? { alpha } : {}) }
}

function unfilter(filter: number, line: Uint8Array, prev: Uint8Array, bpp: number) {
  const n = line.length
  switch (filter) {
    case 1:
      for (let i = bpp; i < n; i++) line[i] = (line[i] + line[i - bpp]) & 255
      break
    case 2:
      for (let i = 0; i < n; i++) line[i] = (line[i] + prev[i]) & 255
      break
    case 3:
      for (let i = 0; i < n; i++) line[i] = (line[i] + (((i >= bpp ? line[i - bpp] : 0) + prev[i]) >> 1)) & 255
      break
    case 4:
      for (let i = 0; i < n; i++) {
        const a = i >= bpp ? line[i - bpp] : 0
        const b = prev[i]
        const c = i >= bpp ? prev[i - bpp] : 0
        const p = a + b - c
        const pa = Math.abs(p - a)
        const pb = Math.abs(p - b)
        const pc = Math.abs(p - c)
        line[i] = (line[i] + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c)) & 255
      }
      break
  }
}

/** a GIF's first frame unpacked, on its logical screen (transparent around it) */
export function decodeGif(b: Uint8Array): Pixels {
  if (b.length < 13 || String.fromCharCode(b[0], b[1], b[2]) !== 'GIF') throw new ImageError('Not a GIF file')
  const w = le16(b, 6)
  const h = le16(b, 8)
  if (!w || !h || w * h > MAX_PIXELS) throw new ImageError(`A GIF of ${w} × ${h} pixels is not drawn`)
  let i = 13
  let global: Uint8Array | undefined
  if (b[10] & 0x80) {
    const n = 3 << ((b[10] & 7) + 1)
    global = b.subarray(i, i + n)
    i += n
  }
  let transparent = -1
  const color = new Uint8Array(w * h * 3)
  const alpha = new Uint8Array(w * h)
  while (i < b.length) {
    const block = b[i++]
    if (block === 0x21) {
      const label = b[i++]
      if (label === 0xf9 && b[i] >= 4 && b[i + 1] & 1) transparent = b[i + 4]
      while (i < b.length && b[i]) i += b[i] + 1
      i++
    } else if (block === 0x2c) {
      const fx = le16(b, i)
      const fy = le16(b, i + 2)
      const fw = le16(b, i + 4)
      const fh = le16(b, i + 6)
      const flags = b[i + 8]
      i += 9
      let table = global
      if (flags & 0x80) {
        const n = 3 << ((flags & 7) + 1)
        table = b.subarray(i, i + n)
        i += n
      }
      const minCode = b[i++]
      const data: number[] = []
      while (i < b.length && b[i]) {
        for (let k = 1; k <= b[i] && i + k < b.length; k++) data.push(b[i + k])
        i += b[i] + 1
      }
      const idx = lzw(data, minCode, fw * fh)
      const rows = (flags & 0x40 ? interlacedRows(fh) : null) ?? Array.from({ length: fh }, (_, r) => r)
      for (let k = 0; k < idx.length; k++) {
        const row = rows[Math.floor(k / fw)]
        const x = fx + (k % fw)
        const y = fy + row
        if (x >= w || y >= h) continue
        const c = idx[k]
        if (c === transparent) continue
        const p = y * w + x
        color[p * 3] = table?.[c * 3] ?? 0
        color[p * 3 + 1] = table?.[c * 3 + 1] ?? 0
        color[p * 3 + 2] = table?.[c * 3 + 2] ?? 0
        alpha[p] = 255
      }
      break
    } else break
  }
  return { w, h, channels: 3, color, ...(alpha.every((a) => a === 255) ? {} : { alpha }) }
}

function interlacedRows(h: number): number[] {
  const out: number[] = []
  for (const [start, step] of [[0, 8], [4, 8], [2, 4], [1, 2]]) for (let r = start; r < h; r += step) out.push(r)
  return out
}

function lzw(data: number[], minCode: number, count: number): Uint8Array {
  const out = new Uint8Array(count)
  if (minCode < 2 || minCode > 11) return out
  const clear = 1 << minCode
  const end = clear + 1
  let size = minCode + 1
  const prefix = new Int32Array(4096)
  const suffix = new Uint8Array(4096)
  const first = new Uint8Array(4096)
  let next = end + 1
  let old = -1
  let n = 0
  let bits = 0
  let acc = 0
  const stack = new Uint8Array(4097)
  for (let k = 0; k < clear; k++) (prefix[k] = -1), (suffix[k] = k), (first[k] = k)
  for (let p = 0; p < data.length && n < count; ) {
    while (bits < size && p < data.length) (acc |= data[p++] << bits), (bits += 8)
    if (bits < size) break
    const code = acc & ((1 << size) - 1)
    acc >>>= size
    bits -= size
    if (code === clear) {
      size = minCode + 1
      next = end + 1
      old = -1
      continue
    }
    if (code === end) break
    let c = code
    let top = 0
    if (code >= next) {
      if (old < 0) break
      stack[top++] = first[old]
      c = old
    }
    while (c >= 0 && top < 4097) {
      stack[top++] = suffix[c]
      c = c < clear ? -1 : prefix[c]
    }
    while (top && n < count) out[n++] = stack[--top]
    if (old >= 0 && next < 4096) {
      prefix[next] = old
      suffix[next] = code === next ? first[old] : first[code]
      first[next] = first[old]
      next++
      if (next === 1 << size && size < 12) size++
    }
    old = code
  }
  return out
}
