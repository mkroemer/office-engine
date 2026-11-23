import { afterAll, describe, expect, test } from 'bun:test'
import { HttpRangeSource } from '../src/core/source'

/** `data` served by range; `cut(from)`: how many bytes of an answer are sent before it ends (a browser or a proxy that cuts long answers) */
function serve(data: Uint8Array, cut: (from: number) => number) {
  const asked: string[] = []
  const server = Bun.serve({
    port: 0,
    fetch(req) {
      const m = /bytes=(\d+)-(\d+)/.exec(req.headers.get('range') ?? '')
      if (!m) return new Response(data as BodyInit)
      const a = Number(m[1])
      const b = Number(m[2])
      asked.push(m[0])
      const end = Math.min(b + 1, a + cut(a))
      return new Response(data.subarray(a, end) as BodyInit, { status: 206, headers: { 'Content-Range': `bytes ${a}-${b}/${data.length}` } })
    },
  })
  return { url: `http://localhost:${server.port}/big.csv`, asked, stop: () => server.stop(true) }
}

async function readAll(s: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const parts: Uint8Array[] = []
  const r = s.getReader()
  for (;;) {
    const { done, value } = await r.read()
    if (done) break
    parts.push(value)
  }
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

describe('a file read from a URL as one stream', () => {
  const data = new Uint8Array(5_000_000)
  for (let i = 0; i < data.length; i++) data[i] = (i * 31 + (i >> 8)) & 255
  const stops: (() => void)[] = []
  afterAll(() => stops.forEach((f) => f()))

  test('an answer that ends before its range is asked for again from there: the whole file is read', async () => {
    // (as a browser does past 2 GiB: every answer stops after a while)
    const s = serve(data, (from) => (from === 0 ? 1 : 1_300_000))
    stops.push(s.stop)
    const src = await HttpRangeSource.open(s.url)
    expect(src.size).toBe(data.length)
    const got = await readAll(src.stream(0, data.length))
    expect(got.length).toBe(data.length)
    expect(Buffer.compare(Buffer.from(got), Buffer.from(data))).toBe(0)
    // the first answer of the stream gave one byte, the next ones went on from where the last ended
    expect(s.asked.slice(1, 4)).toEqual(['bytes=0-4999999', 'bytes=1-4999999', 'bytes=1300001-4999999'])
  })

  test('an answer that gives nothing more is an error, not the end of the file', async () => {
    const s = serve(data, (from) => (from < 2_000_000 ? 2_000_000 - from : 0))
    stops.push(s.stop)
    const src = await HttpRangeSource.open(s.url)
    await expect(readAll(src.stream(0, data.length))).rejects.toThrow('The download ended early: 2,000,000 of 5,000,000 bytes')
  })

  test('part of the file: the range asked for, no more', async () => {
    const s = serve(data, () => 1_000_000)
    stops.push(s.stop)
    const src = await HttpRangeSource.open(s.url)
    const got = await readAll(src.stream(1_000_000, 3_000_000))
    expect(Buffer.compare(Buffer.from(got), Buffer.from(data.subarray(1_000_000, 4_000_000)))).toBe(0)
  })
})
