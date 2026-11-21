/**
 * Numbers typed into cells in a locale (#17): what is saved, what is read again, formulas, collaboration. Without a
 * locale every input reads as it did before.
 */
import { describe, expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'
import { CollabHub } from '../src/collab/hub'
import { compact, startCollab, type Collab } from '../src/collab'
import { MemoryTransport } from '../src/collab/transport'
import { at, book, openBook, rows, type Opened } from './xlsx-builder'

const NNBSP = '\u202f'
const NBSP = '\u00a0'

/** what each input typed without a locale gave before locales were read: type, value, text shown (from `main`) */
const BEFORE: [string, string, string, string][] = [
  ['1', 'number', '1', '1'],
  ['-1', 'number', '-1', '-1'],
  ['+1', 'number', '1', '1'],
  ['1.5', 'number', '1.5', '1.5'],
  ['-1.5', 'number', '-1.5', '-1.5'],
  ['.5', 'number', '0.5', '0.5'],
  ['-.5', 'number', '-0.5', '-0.5'],
  ['+.5', 'number', '0.5', '0.5'],
  ['5.', 'number', '5', '5'],
  ['1e3', 'number', '1000', '1000'],
  ['1E3', 'number', '1000', '1000'],
  ['1.5e-3', 'number', '0.0015', '0.0015'],
  ['-2.5E+2', 'number', '-250', '-250'],
  ['1e999', 'string', '1e999', '1e999'],
  ['1,5', 'string', '1,5', '1,5'],
  ['1,234', 'number', '1234', '1,234'],
  ['1,234,567', 'number', '1234567', '1,234,567'],
  ['1,234.5', 'number', '1234.5', '1,234.5'],
  ['1,234.56', 'number', '1234.56', '1,234.56'],
  ['12,34,567', 'string', '12,34,567', '12,34,567'],
  ['1,23', 'string', '1,23', '1,23'],
  ['1,2345', 'string', '1,2345', '1,2345'],
  ['1.234,5', 'string', '1.234,5', '1.234,5'],
  ['1 234', 'string', '1 234', '1 234'],
  [`1${NBSP}234`, 'string', `1${NBSP}234`, `1${NBSP}234`],
  ['50%', 'number', '0.5', '50%'],
  ['12.5%', 'number', '0.125', '12.5%'],
  ['-5%', 'number', '-0.05', '-5%'],
  ['.5%', 'number', '0.005', '0.5%'],
  ['5.%', 'number', '0.05', '5%'],
  ['50 %', 'string', '50 %', '50 %'],
  ['1,000%', 'string', '1,000%', '1,000%'],
  ['1e2%', 'string', '1e2%', '1e2%'],
  ['%', 'string', '%', '%'],
  ['5%%', 'string', '5%%', '5%%'],
  ['$50', 'number', '50', '$50'],
  ['-$3.75', 'number', '-3.75', '-$3.75'],
  ['+$5', 'number', '5', '$5'],
  ['$-5', 'string', '$-5', '$-5'],
  ['€1,200.50', 'number', '1200.5', '€1,200.50'],
  ['€1,200.5', 'number', '1200.5', '€1,200.50'],
  ['50 ₫', 'number', '50', '50 ₫'],
  ['50₫', 'number', '50', '50 ₫'],
  ['$5€', 'string', '$5€', '$5€'],
  ['$ 1,234', 'number', '1234', '$1,234'],
  ['- 1,234', 'number', '-1234', '-1,234'],
  ['- 5', 'string', '- 5', '- 5'],
  ['1,234.', 'string', '1,234.', '1,234.'],
  ['1,000.5.5', 'string', '1,000.5.5', '1,000.5.5'],
  ['1..5', 'string', '1..5', '1..5'],
  ['1.2.3', 'string', '1.2.3', '1.2.3'],
  ['192.168.1.1', 'string', '192.168.1.1', '192.168.1.1'],
  ['−5', 'string', '−5', '−5'],
  ['١٢', 'string', '١٢', '١٢'],
  ['1234567890123456', 'string', '1234567890123456', '1234567890123456'],
  ['-1234567890123456', 'string', '-1234567890123456', '-1234567890123456'],
  ['123456789012345', 'number', '123456789012345', '1.23457E+14'],
  ['TRUE', 'boolean', 'TRUE', 'TRUE'],
  ['abc', 'string', 'abc', 'abc'],
  ['12a', 'string', '12a', '12a'],
  ['a12', 'string', 'a12', 'a12'],
  ['0x1F', 'string', '0x1F', '0x1F'],
  ['Infinity', 'string', 'Infinity', 'Infinity'],
  ['NaN', 'string', 'NaN', 'NaN'],
  ['1_000', 'string', '1_000', '1_000'],
  [' 12 ', 'number', '12', '12'],
  [`12${NBSP}`, 'number', '12', '12'],
  ['1/2', 'string', '1/2', '1/2'],
  ['1 1/2', 'string', '1 1/2', '1 1/2'],
  ['(5)', 'string', '(5)', '(5)'],
  ['12:30', 'string', '12:30', '12:30'],
  ['2024-01-15', 'number', '45306', '2024-01-15'],
  ['15/01/2024', 'number', '45306', '15/01/2024'],
  ['#N/A', 'error', '#N/A', '#N/A'],
  ['+84 912 345 678', 'string', '+84 912 345 678', '+84 912 345 678'],
  ['1,234e3', 'string', '1,234e3', '1,234e3'],
  ['$1,234e3', 'string', '$1,234e3', '$1,234e3'],
  ['1.5.', 'string', '1.5.', '1.5.'],
  ['-0', 'number', '0', '0'],
  ['00012', 'number', '12', '12'],
  ['0012.50', 'number', '12.5', '12.5'],
  ['1,234,5678', 'string', '1,234,5678', '1,234,5678'],
  [',5', 'string', ',5', ',5'],
  ['1,', 'string', '1,', '1,'],
  ['-', 'string', '-', '-'],
  ['+', 'string', '+', '+'],
  ['.', 'string', '.', '.'],
  ['e5', 'string', 'e5', 'e5'],
  ['1e', 'string', '1e', '1e'],
  ['1e+', 'string', '1e+', '1e+'],
  ['$', 'string', '$', '$'],
  ['$.5', 'string', '$.5', '$.5'],
  ['$5.', 'string', '$5.', '$5.'],
  ['1 000 000', 'string', '1 000 000', '1 000 000'],
  ["1'234", 'string', "1'234", "1'234"],
  ['1’234', 'string', '1’234', '1’234'],
  ['$1,234.567', 'number', '1234.567', '$1,234.567'],
  ['1,234 €', 'number', '1234', '1,234 €'],
  ['€ 1,234.5', 'number', '1234.5', '€1,234.50'],
  ['12.50 ₫', 'number', '12.5', '12.50 ₫'],
  ['£0.5', 'number', '0.5', '£0.50'],
  ['1,234,567.891', 'number', '1234567.891', '1,234,567.891'],
]

const detail = async (o: Opened, ref: string) => (await o.engine.call('xlsx.cell', [0, at(ref).r, at(ref).c])) as { type: string; value: string; formula?: string }
const typeIn = (o: Opened, ref: string, input: string, locale?: string) => o.engine.call('xlsx.set', [0, [{ ...at(ref), input }], locale ? { locale } : undefined])
const blank = async () => openBook(await book({ sheets: [{ name: 'S', rows: rows([[1]]) }] }))

describe('cells: without a locale, as before', () => {
  for (const locale of [undefined, 'en-US', 'de_DE']) {
    test(`every input reads as it did (${locale ?? 'no locale'})`, async () => {
      const o = await blank()
      const got: unknown[] = []
      for (let i = 0; i < BEFORE.length; i++) {
        const r = 2 + i
        await o.engine.call('xlsx.set', [0, [{ r, c: 0, input: BEFORE[i][0] }], locale ? { locale } : undefined])
        const d = (await o.engine.call('xlsx.cell', [0, r, 0])) as { type: string; value: string }
        got.push([BEFORE[i][0], d?.type, d?.value, await o.shown(0, r, 0)])
      }
      expect(got).toEqual(BEFORE)
    })
  }
})

describe('cells: typed in a locale', () => {
  test('de-DE: numbers, percentages, amounts; what is not one stays text', async () => {
    const o = await blank()
    const typed: [string, string, string, string][] = [
      ['A2', '1,5', 'number', '1.5'],
      ['B2', '-1,25', 'number', '-1.25'],
      ['C2', '1.234,5', 'number', '1234.5'],
      ['D2', '1.234', 'number', '1234'],
      ['E2', '50 %', 'number', '0.5'],
      ['F2', '12,5%', 'number', '0.125'],
      ['G2', '1,5e3', 'number', '1500'],
      ['H2', '1.234,50 €', 'number', '1234.5'],
      ['I2', '1.5', 'string', '1.5'],
      ['J2', '1,234.5', 'string', '1,234.5'],
      ['K2', '1.2345', 'string', '1.2345'],
      ['L2', '12,5pt', 'string', '12,5pt'],
    ]
    for (const [ref, input] of typed) await typeIn(o, ref, input, 'de-DE')
    for (const [ref, input, type, value] of typed) expect([input, (await detail(o, ref)).type, (await detail(o, ref)).value]).toEqual([input, type, value])
    // the formats derived are format codes, the same in every locale
    expect([await o.shown(0, 1, 2), await o.shown(0, 1, 3), await o.shown(0, 1, 4), await o.shown(0, 1, 5), await o.shown(0, 1, 7)]).toEqual(['1,234.5', '1,234', '50%', '12.5%', '1,234.50 €'])
  })

  test('the saved file holds the number: 1,5 typed in de-DE is <v>1.5</v>; without a locale it stays text; read again the same', async () => {
    const o = await blank()
    await typeIn(o, 'A2', '1,5', 'de-DE')
    await typeIn(o, 'B2', '1,5')
    await typeIn(o, 'C2', '1.234,5', 'de-DE')
    await typeIn(o, 'D2', `1${NNBSP}234,5`, 'fr-FR')
    await typeIn(o, 'E2', '12,34,567.5', 'en-IN')
    const saved = await o.save()
    const xml = await saved.text('xl/worksheets/sheet1.xml')
    expect(xml).toMatch(/<c r="A2"( s="\d+")?><v>1\.5<\/v><\/c>/)
    expect(xml).toMatch(/<c r="B2"( s="\d+")? t="(s|inlineStr)">/)
    expect(xml).toMatch(/<c r="C2"( s="\d+")?><v>1234\.5<\/v><\/c>/)
    expect(xml).toMatch(/<c r="D2"( s="\d+")?><v>1234\.5<\/v><\/c>/)
    expect(xml).toMatch(/<c r="E2"( s="\d+")?><v>1234567\.5<\/v><\/c>/)
    const re = await openBook(saved.blob)
    expect([(await detail(re, 'A2')).type, (await detail(re, 'A2')).value]).toEqual(['number', '1.5'])
    expect([(await detail(re, 'B2')).type, (await detail(re, 'B2')).value]).toEqual(['string', '1,5'])
    for (const ref of ['C2', 'D2']) expect(Number((await detail(re, ref)).value)).toBe(1234.5)
    for (const c of [0, 1, 2, 3, 4]) expect(await re.shown(0, 1, c)).toBe(await o.shown(0, 1, c))
  })

  test('1,5 shown for 1.5 in de-DE, typed back, is still 1.5', async () => {
    const o = await openBook(await book({ sheets: [{ name: 'S', rows: rows([[1.5, -1234.25]]) }] }))
    await typeIn(o, 'A1', '1,5', 'de-DE')
    await typeIn(o, 'B1', '-1234,25', 'de-DE')
    expect([(await detail(o, 'A1')).type, (await detail(o, 'A1')).value, (await detail(o, 'B1')).value]).toEqual(['number', '1.5', '-1234.25'])
  })

  test('formulas are not read in the locale: their text is kept as typed', async () => {
    const o = await blank()
    await typeIn(o, 'A2', '=1.5+A1', 'de-DE')
    await typeIn(o, 'B2', '=SUM(1,5)', 'de-DE')
    await typeIn(o, 'C2', '+A1*1.5', 'de-DE')
    expect([await o.formula(0, 1, 0), await o.shown(0, 1, 0)]).toEqual(['=1.5+A1', '2.5'])
    expect([await o.formula(0, 1, 1), await o.shown(0, 1, 1)]).toEqual(['=SUM(1,5)', '6'])
    expect([await o.formula(0, 1, 2), await o.shown(0, 1, 2)]).toEqual(['=+A1*1.5', '1.5'])
    const xml = await (await o.save()).text('xl/worksheets/sheet1.xml')
    expect(xml).toContain('<f>1.5+A1</f>')
    expect(xml).toContain('<f>SUM(1,5)</f>')
  })

  test('a cell formatted as Text keeps what is typed; a cell showing percentages reads 50 as 50%', async () => {
    const o = await openBook(
      await book({
        sheets: [{ name: 'S', rows: rows([[{ s: 1 }, { s: 2, v: 0 }, { s: 2, v: 0 }]]) }],
        styles: `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`,
      }),
    )
    await typeIn(o, 'A1', '1,5', 'de-DE')
    await typeIn(o, 'B1', '12,5', 'de-DE')
    await typeIn(o, 'C1', '0,5', 'de-DE')
    expect([(await detail(o, 'A1')).type, (await detail(o, 'A1')).value]).toEqual(['string', '1,5'])
    expect([await o.shown(0, 0, 1), await o.shown(0, 0, 2)]).toEqual(['12.50%', '50.00%'])
  })

  test('a locale given to one edit changes no value already in the workbook', async () => {
    const src = await book({ sheets: [{ name: 'S', rows: rows([[1.5, 1234.5, '1,5', { f: 'A1*2', v: 3 }]]) }] })
    const o = await openBook(src)
    const before = await (await o.save()).text('xl/worksheets/sheet1.xml')
    await typeIn(o, 'A3', '2,5', 'de-DE')
    for (const [ref, v] of [['A1', '1.5'], ['B1', '1234.5'], ['C1', '1,5'], ['D1', '3']]) expect((await detail(o, ref)).value).toBe(v)
    const after = await (await o.save()).text('xl/worksheets/sheet1.xml')
    const row1 = (xml: string) => /<row r="1"[^>]*>.*?<\/row>/.exec(xml)![0]
    expect(row1(after)).toBe(row1(before))
  })
})

// ---------------------------------------------------------------- collaboration

describe('collaboration: the locale goes with the edit', () => {
  const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))
  const SAMPLE = `${import.meta.dir}/../playground/public/samples/sample.xlsx`
  const peer = async (hub: CollabHub, bytes: Uint8Array, name: string) => {
    const client = await DocumentClient.open(new Blob([bytes as BlobPart]), { worker: false })
    await client.call('xlsx.load', 0)
    await sleep(30)
    const collab = startCollab(client, { transport: new MemoryTransport(hub, () => 1), room: 'x', user: { name } })
    return { client, collab }
  }
  const settle = async (cs: Collab[]) => {
    for (let i = 0; i < 300; i++) {
      await sleep(10)
      await Promise.all(cs.map((c) => c.session.idle()))
      if (cs.every((c) => c.pending === 0 && c.status === 'live') && new Set(cs.map((c) => c.session.rev)).size === 1) return
    }
    throw new Error('did not settle')
  }
  const values = async (client: DocumentClient) => {
    const out: unknown[] = []
    for (const c of [0, 1, 2, 3, 4]) {
      const d = await client.call('xlsx.cell', 0, 40, c)
      out.push([d?.type, d?.value])
    }
    return out
  }

  test('typed in de-DE by one, read the same by the others, by a late joiner and by a server replaying the log', async () => {
    const bytes = new Uint8Array(await Bun.file(SAMPLE).arrayBuffer())
    const hub = new CollabHub()
    const ann = await peer(hub, bytes, 'Ann')
    const bob = await peer(hub, bytes, 'Bob')
    await settle([ann.collab, bob.collab])
    await ann.client.call('xlsx.set', 0, [{ r: 40, c: 0, input: '1,5' }, { r: 40, c: 1, input: '1.234,5' }, { r: 40, c: 2, input: '50 %' }], { locale: 'de-de' })
    // Bob types without a locale: 1,5 is text for him, everywhere
    await bob.client.call('xlsx.set', 0, [{ r: 40, c: 3, input: '1,5' }])
    await ann.client.call('xlsx.set', 0, [{ r: 40, c: 4, input: '2,5' }], { locale: 'de_DE' })
    await settle([ann.collab, bob.collab])
    const expected = [['number', '1.5'], ['number', '1234.5'], ['number', '0.5'], ['string', '1,5'], ['string', '2,5']]
    expect(await values(ann.client)).toEqual(expected)
    expect(await values(bob.client)).toEqual(expected)
    // the op carries the canonical tag; an op without a locale carries none
    const log = hub.rooms.get('x')!.log.flatMap((b) => b.ops) as { t: string; locale?: string }[]
    expect(log.filter((o) => o.t === 'set').map((o) => o.locale)).toEqual(['de-DE', undefined, undefined])

    const cid = await peer(hub, bytes, 'Cid')
    await settle([ann.collab, bob.collab, cid.collab])
    expect(await values(cid.client)).toEqual(expected)

    const server = await DocumentClient.open(new Blob([bytes as BlobPart]), { worker: false })
    await server.call('xlsx.load', 0)
    await sleep(30)
    await compact(server, { batches: hub.rooms.get('x')!.log })
    expect(await values(server)).toEqual(expected)
    for (const p of [ann, bob, cid]) p.collab.close()
  })
})
