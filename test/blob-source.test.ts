import { describe, expect, test } from 'bun:test'
import { BlobSource } from '../src/core/source'
import { ZipArchive } from '../src/core/zip'
import { crc32 } from '../src/core/zipwrite'

const bytes = (n: number) => Uint8Array.from({ length: n }, (_, i) => (i * 7 + (i >> 8)) & 255)

async function readAll(s: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const parts: Uint8Array[] = []
  const r = s.getReader()
  for (;;) {
    const { done, value } = await r.read()
    if (done) break
    parts.push(value)
  }
  return new Uint8Array(await new Blob(parts as BlobPart[]).arrayBuffer())
}

/**
 * A Blob as Bun 1.3 gives it: the stream of a slice does not stop at the end of the slice, it runs on to the end of
 * the whole Blob. `chunk`: the size of the pieces it comes in; `cancelled` counts the streams that were cancelled.
 */
function overrunning(data: Uint8Array, chunk: number, empty = false) {
  const seen = { cancelled: 0 }
  const blob = {
    size: data.length,
    type: '',
    slice(from: number, to: number) {
      return {
        size: Math.max(0, Math.min(to, data.length) - from),
        arrayBuffer: async () => data.slice(from, to).buffer,
        stream() {
          let at = from
          let n = 0
          return new ReadableStream<Uint8Array>({
            pull(ctrl) {
              if (at >= data.length) return ctrl.close()
              // (every other piece empty, when asked for)
              if (empty && n++ % 2 === 0) return ctrl.enqueue(new Uint8Array(0))
              ctrl.enqueue(data.slice(at, (at += chunk)))
            },
            cancel() {
              seen.cancelled++
            },
          })
        },
      }
    },
  }
  return { blob: blob as unknown as Blob, seen }
}

describe('part of a Blob read as a stream', () => {
  const data = bytes(1000)

  test('the bytes asked for, no more, from a real Blob', async () => {
    const src = new BlobSource(new Blob([data as BlobPart]))
    expect(await readAll(src.stream(10, 39))).toEqual(data.slice(10, 49))
    expect(await readAll(src.stream(0, 1000))).toEqual(data)
    expect((await readAll(src.stream(500, 0))).length).toBe(0)
    // past the end of the file: what there is
    expect(await readAll(src.stream(990, 50))).toEqual(data.slice(990))
    expect((await readAll(src.stream(2000, 10))).length).toBe(0)
  })

  // https://github.com/p10node/office-engine/issues/49
  test('a runtime whose slice streams run on to the end of the Blob: cut at the length asked for', async () => {
    for (const chunk of [1, 7, 39, 64, 4096]) {
      const src = new BlobSource(overrunning(data, chunk).blob)
      expect(await readAll(src.stream(10, 39))).toEqual(data.slice(10, 49))
      expect(await readAll(src.stream(0, 1000))).toEqual(data)
      expect((await readAll(src.stream(500, 0))).length).toBe(0)
    }
    // the rest of the runtime's stream is not read
    const { blob, seen } = overrunning(data, 16)
    await readAll(new BlobSource(blob).stream(10, 39))
    expect(seen.cancelled).toBe(1)
  })

  test('empty pieces in the stream are read past', async () => {
    const src = new BlobSource(overrunning(data, 16, true).blob)
    expect(await readAll(src.stream(100, 300))).toEqual(data.slice(100, 400))
  })

  test('cancelling the stream cancels the Blob’s', async () => {
    const { blob, seen } = overrunning(data, 16)
    const r = new BlobSource(blob).stream(0, 500).getReader()
    await r.read()
    await r.cancel()
    expect(seen.cancelled).toBe(1)
  })

  /** a ZIP of stored entries, each with its checksum */
  function storedZip(files: [string, string][]): Uint8Array {
    const enc = new TextEncoder()
    const u16 = (n: number) => [n & 255, (n >> 8) & 255]
    const u32 = (n: number) => [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]
    const local: number[] = []
    const central: number[] = []
    for (const [name, text] of files) {
      const data = enc.encode(text)
      const head = [...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(crc32(data)), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0)]
      central.push(...u32(0x02014b50), ...u16(20), ...head, ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(local.length), ...enc.encode(name))
      local.push(...u32(0x04034b50), ...head, ...enc.encode(name), ...data)
    }
    return new Uint8Array([...local, ...central, ...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(central.length), ...u32(local.length), ...u16(0)])
  }

  test('a stored ZIP entry followed by others matches its checksum (the “mimetype” of an .odt)', async () => {
    const mime = 'application/vnd.oasis.opendocument.text'
    const zipped = storedZip([['mimetype', mime], ['content.xml', '<a/>'.repeat(4000)], ['meta.xml', '<m/>']])
    for (const blob of [new Blob([zipped as BlobPart]), overrunning(zipped, 4096).blob]) {
      const zip = await ZipArchive.open(new BlobSource(blob))
      expect(await zip.text('mimetype')).toBe(mime)
      expect((await zip.text('content.xml')).length).toBe(16000)
      expect(await zip.text('meta.xml')).toBe('<m/>')
    }
  })
})
