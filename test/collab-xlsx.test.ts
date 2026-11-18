import { describe, expect, test } from 'bun:test'
import { getLocalHistory, subscribeLocalHistory } from '../src/collab/hosts'
import { DocumentClient } from '../src/core/worker/client'
import { CollabHub } from '../src/collab/hub'
import { startCollab, type Collab } from '../src/collab'
import { MemoryTransport } from '../src/collab/transport'
import { ZipArchive } from '../src/core/zip'
import { openSource } from '../src/core/source'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`
const file = (n: string) => Bun.file(S(n)) as unknown as Blob
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

async function open(hub: CollabHub, name: string, latency: () => number) {
  const client = await DocumentClient.open(file('sample.xlsx'), { worker: false })
  await client.call('xlsx.load', 0)
  await sleep(50)
  const collab = startCollab(client, { transport: new MemoryTransport(hub, latency), room: 'x', user: { name } })
  return { client, collab }
}

async function settle(cs: Collab[]) {
  for (let i = 0; i < 300; i++) {
    await sleep(10)
    await Promise.all(cs.map((c) => c.session.idle()))
    if (cs.every((c) => c.pending === 0 && c.status === 'live') && new Set(cs.map((c) => c.session.rev)).size === 1) return
  }
  throw new Error('did not settle')
}

/** shown text + style css of the first rows / cols of every sheet */
async function dump(client: DocumentClient) {
  const meta = client.info.xlsx!
  const out: unknown[] = []
  for (const s of meta.sheets) {
    await client.call('xlsx.load', s.index)
    await sleep(20)
    const b = await client.call('xlsx.block', s.index, 0, 40)
    out.push(s.name, b.rows.map((r) => [r.r, r.cells.map((c) => [c[0], c[1], c[2]])]))
  }
  // the shared style table must be identical (style indexes are exchanged by paint-format ops)
  const zip = await ZipArchive.open(await openSource(await client.call('save')))
  out.push(new TextDecoder().decode(await zip.bytes('xl/styles.xml')))
  return JSON.stringify(out)
}

describe('collab xlsx', () => {
  test('local history reports the session stack across rename, new edits and undo/redo', async () => {
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => 0)
    await settle([a.collab])
    let notifications = 0
    const off = subscribeLocalHistory(a.client, () => notifications++)
    await a.client.call('xlsx.set', 0, [{ r: 11, c: 9, input: 'before rename' }])
    await settle([a.collab])
    expect(getLocalHistory(a.client)).toEqual({ undo: 1, redo: 0 })
    await a.client.call('xlsx.bookOp', { op: 'rename', index: 0, name: 'Renamed' })
    await settle([a.collab])
    // Sheet-list operations invalidate the worker's shared undo records as well as its local history.
    expect(getLocalHistory(a.client)).toEqual({ undo: 0, redo: 0 })
    expect(await a.client.call('xlsx.undo')).toBeNull()
    await a.client.call('xlsx.set', 0, [{ r: 11, c: 9, input: 'after rename' }])
    await settle([a.collab])
    expect(getLocalHistory(a.client)).toEqual({ undo: 1, redo: 0 })
    await a.client.call('xlsx.undo')
    expect(getLocalHistory(a.client)).toEqual({ undo: 0, redo: 1 })
    expect((await a.client.call('xlsx.cell', 0, 11, 9))?.value).toBe('before rename')
    await a.client.call('xlsx.redo')
    expect(getLocalHistory(a.client)).toEqual({ undo: 1, redo: 0 })
    expect((await a.client.call('xlsx.cell', 0, 11, 9))?.value).toBe('after rename')
    expect(notifications).toBeGreaterThan(4)
    off(); a.collab.close(); a.client.destroy()
    expect(getLocalHistory(a.client)).toBeUndefined()
  })

  for (const start of [7, 11, 23, 42, 99]) test(`concurrent cell edits, formatting, structure and undo converge (seed ${start})`, async () => {
    let seed = start
    const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296)
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => Math.floor(rand() * 6))
    const b = await open(hub, 'Bob', () => Math.floor(rand() * 6))
    await settle([a.collab, b.collab])

    const ops: Promise<unknown>[] = []
    const run = (p: Promise<unknown>) => ops.push(p.catch((e) => console.error(e)))
    for (let i = 0; i < 70; i++) {
      const who = rand() < 0.5 ? a : b
      const x = rand()
      const r = Math.floor(rand() * 12)
      const c = Math.floor(rand() * 5)
      if (x < 0.45) run(who.client.call('xlsx.set', 0, [{ r, c, input: rand() < 0.3 ? `=A${r + 2}*2` : `v${i}` }]))
      else if (x < 0.65) run(who.client.call('xlsx.format', 0, { r0: r, c0: c, r1: r + 1, c1: c + 1 }, rand() < 0.5 ? { b: rand() < 0.5 } : { fill: rand() < 0.5 ? '#ffee00' : '#00ccff' }))
      else if (x < 0.72) run(who.client.call('xlsx.sheetOp', 0, { op: rand() < 0.6 ? 'insertRows' : 'deleteRows', at: r, n: 1 }))
      else if (x < 0.8) run(who.client.call('xlsx.undo'))
      else if (x < 0.85) run(who.client.call('xlsx.redo'))
      else if (x < 0.92) run(who.client.call('xlsx.sheetOp', 0, { op: 'colWidth', c0: c, c1: c, px: 40 + i }))
      else run(who.client.call('xlsx.sheetOp', 0, { op: 'fill', range: { r0: r, c0: c, r1: r + 2, c1: c }, dir: 'down' }))
      if (rand() < 0.4) await sleep(Math.floor(rand() * 5))
    }
    await Promise.all(ops)
    await settle([a.collab, b.collab])
    const da = await dump(a.client)
    const db = await dump(b.client)
    expect(da.match(/"v\d+"/g)?.length ?? 0).toBeGreaterThan(5)
    expect(da).toBe(db)

    // a late joiner replays the log to the same workbook
    const c = await open(hub, 'Cid', () => 0)
    await settle([a.collab, b.collab, c.collab])
    expect(await dump(c.client)).toBe(da)

    // sheet list operations go through the hub order
    await Promise.all([a.client.call('xlsx.bookOp', { op: 'add', name: 'Extra' }), b.client.call('xlsx.bookOp', { op: 'rename', index: 1, name: 'Renamed' })])
    await settle([a.collab, b.collab, c.collab])
    const names = (x: DocumentClient) => x.info.xlsx!.sheets.map((s) => s.name).join(',')
    expect(names(a.client)).toBe(names(b.client))
    expect(names(c.client)).toBe(names(a.client))
    expect(names(a.client)).toContain('Renamed')
    expect(names(a.client)).toContain('Extra')
    for (const x of [a, b, c]) {
      x.collab.close()
      x.client.destroy()
    }
  }, 30_000)

  for (const lag of [[0, 6], [6, 0], [3, 3]]) test(`text with line breaks (wrap style, grown row) against concurrent formatting and row heights converges (latency ${lag})`, async () => {
    const hub = new CollabHub()
    const a = await open(hub, 'Ann', () => lag[0])
    const b = await open(hub, 'Bob', () => lag[1])
    await settle([a.collab, b.collab])
    /** cells, style table and the row heights saved for the first sheet */
    const state = async (client: DocumentClient) => {
      const zip = await ZipArchive.open(await openSource(await client.call('save')))
      const xml = new TextDecoder().decode(await zip.bytes(client.info.xlsx!.sheets[0].path))
      const rows = [...xml.matchAll(/<row [^>]*>/g)].map((m) => `${/ r="(\d+)"/.exec(m[0])?.[1]}:${/ ht="([^"]+)"/.exec(m[0])?.[1] ?? ''}`)
      return (await dump(client)) + rows.join(' ')
    }
    await Promise.all([
      a.client.call('xlsx.set', 0, [{ r: 20, c: 6, input: 'one\ntwo\nthree' }]),
      b.client.call('xlsx.format', 0, { r0: 22, c0: 6, r1: 22, c1: 6 }, { i: true }),
      b.client.call('xlsx.sheetOp', 0, { op: 'rowHeight', r0: 20, r1: 20, px: 24 }),
      b.client.call('xlsx.set', 0, [{ r: 22, c: 7, input: 'x\ny' }]),
    ])
    await settle([a.collab, b.collab])
    const sa = await state(a.client)
    expect(sa).toContain(JSON.stringify('one\ntwo\nthree'))
    expect(sa).toBe(await state(b.client))
    await Promise.all([a.client.call('xlsx.undo'), b.client.call('xlsx.set', 0, [{ r: 20, c: 7, input: 'p\nq' }])])
    await settle([a.collab, b.collab])
    expect(await state(a.client)).toBe(await state(b.client))
    const c = await open(hub, 'Cid', () => 0)
    await settle([a.collab, b.collab, c.collab])
    expect(await state(c.client)).toBe(await state(a.client))
    for (const x of [a, b, c]) {
      x.collab.close()
      x.client.destroy()
    }
  }, 30_000)

  // A rebase rolls a local formatting op back, and the styles it derived with it. The index it freed is then given to
  // another style: what was derived *from* that index before must not be reused for it.
  test('a style index freed by a rebase and derived again formats like one never freed', async () => {
    const fmt = (id: string, r: number, c: number, patch: object) => ({ t: 'fmt', id, sheet: '', range: { r0: r, c0: c, r1: r, c1: c }, patch })
    const run = async (rebased: boolean) => {
      const client = await DocumentClient.open(file('sample.xlsx'), { worker: false })
      await client.call('xlsx.load', 0)
      await sleep(20)
      const sheet = client.info.xlsx!.sheets[0].path
      const apply = async (op: ReturnType<typeof fmt>) => (await client.direct('collab.xlsx.apply', { ...op, sheet } as never)).rec as number
      // D5 (currency) turns blue
      await apply(fmt('0', 4, 3, { fill: '#00ccff' }))
      if (rebased) {
        // D3 (currency): yellow derives a new style, blue on top of it is the style D5 already has
        const yellow = await apply(fmt('p1', 2, 3, { fill: '#ffee00' }))
        const blue = await apply(fmt('p2', 2, 3, { fill: '#00ccff' }))
        await client.direct('collab.xlsx.revert', blue)
        await client.direct('collab.xlsx.revert', yellow)
      }
      // F3 (percent): bold takes the freed index, blue is derived from it
      await apply(fmt('1', 2, 5, { b: true }))
      await apply(fmt('2', 2, 5, { fill: '#00ccff' }))
      const out = await dump(client)
      client.destroy()
      return out
    }
    const plain = await run(false)
    // F3 is still a percentage
    expect(JSON.parse(plain)[1][2][1][5][1]).toBe('8.4%')
    const got = await run(true)
    expect(JSON.parse(got)[1][2][1][5][1]).toBe('8.4%')
    expect(got).toBe(plain)
  })
})
