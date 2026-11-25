import { afterAll, describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { BusyError, convert, ConvertError, Converter, createHandler, TimeoutError } from '../packages/convert/src/index'
import { ZipArchive } from '../src/core/zip'
import { BlobSource } from '../src/core/source'

const S = (n: string) => new Uint8Array(readFileSync(`${import.meta.dir}/../playground/public/samples/${n}`))
const text = (d: Uint8Array) => new TextDecoder().decode(d)

/** a CSV that takes the engine a few seconds */
const bigCsv = () => {
  const rows: string[] = ['a,b,c,d']
  for (let i = 0; i < 600_000; i++) rows.push(`${i},name ${i},${i * 1.5},x`)
  return new TextEncoder().encode(rows.join('\n'))
}

describe('convert', () => {
  test('legacy and other formats to the Office file they open as', async () => {
    for (const [file, ext, from] of [
      ['sample.doc', 'docx', 'doc'],
      ['sample.xls', 'xlsx', 'xls'],
      ['sample.ppt', 'pptx', 'ppt'],
      ['sample.odt', 'docx', 'odt'],
      ['sample.key', 'pptx', 'keynote'],
      ['sample.md', 'docx', 'markdown'],
      ['sample.csv', 'xlsx', 'csv'],
    ] as const) {
      const r = await convert(S(file), { name: file })
      expect(r.name).toBe(file.replace(/\.\w+$/, `.${ext}`))
      expect(r.kind).toBe(ext)
      expect(r.from).toBe(from)
      // a real package of that kind: it opens again
      const again = await convert(r.data, { name: r.name })
      expect(again.kind).toBe(ext)
      expect(again.from).toBeUndefined()
    }
  })

  test('documents as text, Markdown and HTML', async () => {
    const md = text((await convert(S('sample.doc'), { name: 'sample.doc', to: 'md' })).data)
    expect(md).toContain('# Highlights')
    expect(md).toContain('[example.com](https://example.com/)')
    const html = await convert(S('sample.md'), { name: 'notes.md', to: 'html' })
    expect(html.type).toStartWith('text/html')
    expect(text(html.data)).toContain('<title>notes</title>')
    expect(text((await convert(S('sample.pptx'), { name: 'deck.pptx', to: 'txt' })).data)).toStartWith('Slide 1\n')
  })

  test('a workbook as CSV: formulas the file has no value for are computed', async () => {
    const r = await convert(S('sample.xlsx'), { name: 'sales.xlsx', to: 'csv' })
    expect(r.name).toBe('sales.csv')
    const rows = text(r.data).replace(/^﻿/, '').split('\r\n')
    expect(rows[2]).toStartWith('North,2026-01-16,195,$21.20,"4,134.00"')
    const tsv = await convert(S('sample.xlsx'), { name: 'sales.xlsx', to: 'tsv', sheet: 0 })
    expect(text(tsv.data)).toContain('North\t2026-01-16')
  })

  test('formats a kind cannot be written as, and PDF, are refused', async () => {
    const e = await convert(S('sample.pptx'), { to: 'csv' }).catch((e) => e)
    expect(e).toBeInstanceOf(ConvertError)
    expect(e.code).toBe('unsupported')
    const pdf = await convert(S('sample.docx'), { to: 'pdf' }).catch((e) => e)
    expect(pdf.code).toBe('not-implemented')
    expect(await convert(S('sample.xlsx'), { to: 'csv', sheet: 'No such sheet' }).catch((e) => e.code)).toBe('unsupported')
  })
})

describe('Converter', () => {
  const pool = new Converter({ workers: 2, timeout: 1500 })
  afterAll(() => pool.close())

  test('converts on worker threads', async () => {
    const [a, b, c] = await Promise.all([pool.convert(S('sample.doc'), { name: 'a.doc' }), pool.convert(S('sample.xls'), { name: 'b.xls', to: 'csv' }), pool.convert(S('sample.odp'), { name: 'c.odp' })])
    expect([a.name, b.name, c.name]).toEqual(['a.docx', 'b.csv', 'c.pptx'])
    expect(a.data).toBeInstanceOf(Uint8Array)
    expect((await convert(a.data, { name: a.name })).kind).toBe('docx')
    // errors keep their kind across the thread
    expect(await pool.convert(S('sample.pptx'), { to: 'csv' }).catch((e) => e)).toBeInstanceOf(ConvertError)
  })

  test('a conversion past the time limit ends its thread; the next one gets a new thread', async () => {
    const e = await pool.convert(bigCsv(), { name: 'big.csv' }).catch((e) => e)
    expect(e).toBeInstanceOf(TimeoutError)
    expect(pool.load.running).toBe(0)
    expect((await pool.convert(S('sample.doc'), { name: 'a.doc' })).name).toBe('a.docx')
  })

  test('a caller giving up ends its conversion', async () => {
    const ac = new AbortController()
    const p = pool.convert(bigCsv(), { name: 'big.csv', signal: ac.signal })
    setTimeout(() => ac.abort(new Error('gave up')), 100)
    expect((await p.catch((e) => e)).message).toBe('gave up')
    expect(pool.load.running).toBe(0)
  })

  test('a full queue refuses more', async () => {
    const small = new Converter({ workers: 1, queue: 1, timeout: 0 })
    try {
      const ac = new AbortController()
      const first = small.convert(bigCsv(), { signal: ac.signal }).catch(() => {})
      const second = small.convert(S('sample.doc')).catch(() => {})
      expect(await small.convert(S('sample.doc')).catch((e) => e)).toBeInstanceOf(BusyError)
      ac.abort()
      await Promise.all([first, second])
    } finally {
      await small.close()
    }
  })
})

describe('HTTP service', () => {
  const converter = new Converter({ workers: 0 })
  const handle = createHandler(converter, { maxBytes: 1024 * 1024, maxFiles: 3, version: '9.9.9' })
  const form = (files: [string, Uint8Array][], fields: Record<string, string> = {}) => {
    const fd = new FormData()
    for (const [name, data] of files) fd.append('files', new File([data as BlobPart], name))
    for (const [k, v] of Object.entries(fields)) fd.append(k, v)
    return fd
  }
  const post = (path: string, body: BodyInit, headers: Record<string, string> = {}) => handle(new Request(`http://x${path}`, { method: 'POST', body, headers }))

  test('health and version, as Gotenberg answers them', async () => {
    const health = await handle(new Request('http://x/health'))
    expect(((await health.json()) as { status: string }).status).toBe('up')
    expect(await (await handle(new Request('http://x/version'))).text()).toBe('9.9.9')
  })

  test('one file gives the file', async () => {
    const res = await post('/convert', form([['report.doc', S('sample.doc')]]))
    expect(res.status).toBe(200)
    expect(res.headers.get('content-type')).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    expect(res.headers.get('content-disposition')).toContain('filename="report.docx"')
    expect(res.headers.get('gotenberg-trace')).toBeTruthy()
  })

  test('several files give a ZIP; Gotenberg-Output-Filename names it, Gotenberg-Trace comes back', async () => {
    const res = await post('/convert', form([['a.odt', S('sample.odt')], ['a.rtf', S('sample.rtf')], ['b.ods', S('sample.ods')]]), { 'Gotenberg-Output-Filename': 'out', 'Gotenberg-Trace': 't-1' })
    expect(res.status).toBe(200)
    expect(res.headers.get('content-disposition')).toContain('filename="out.zip"')
    expect(res.headers.get('gotenberg-trace')).toBe('t-1')
    const zip = await ZipArchive.open(new BlobSource(await res.blob()))
    expect(zip.list().map((e) => e.name).sort()).toEqual(['a (2).docx', 'a.docx', 'b.xlsx'])
  })

  test('the file as the body, its name and format in the query', async () => {
    const res = await post('/convert?to=csv&name=sales.xlsx', S('sample.xlsx') as BodyInit)
    expect(res.headers.get('content-disposition')).toContain('filename="sales.csv"')
    expect(await res.text()).toContain('North,2026-01-16')
  })

  test('errors', async () => {
    const status = async (res: Promise<Response>) => (await res).status
    expect(await status(post('/forms/libreoffice/convert', form([['a.docx', S('sample.docx')]])))).toBe(501)
    expect(await status(post('/convert', form([['a.docx', S('sample.docx')]], { to: 'nope' })))).toBe(400)
    expect(await status(post('/convert', form([['a.pptx', S('sample.pptx')]], { to: 'csv' })))).toBe(400)
    expect(await status(post('/convert', form([['a.docx', new TextEncoder().encode('PK\x03\x04 not a zip')]])))).toBe(422)
    expect(await status(post('/convert', form([['a.docx', new Uint8Array(2 * 1024 * 1024)]])))).toBe(413)
    expect(await status(post('/convert', form([['a', S('sample.md')], ['b', S('sample.md')], ['c', S('sample.md')], ['d', S('sample.md')]])))).toBe(400)
    expect(await status(post('/convert', new FormData()))).toBe(400)
    expect(await status(handle(new Request('http://x/convert')))).toBe(405)
    expect(await status(handle(new Request('http://x/nope')))).toBe(404)
  })
})
