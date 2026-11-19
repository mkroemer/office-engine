/** Second round of the QA pass of 18/11/2026: two people in one workbook (P2-032, P2-033). */
import { describe, expect, test } from 'bun:test'
import { DocumentClient } from '../src/core/worker/client'
import { CollabHub } from '../src/collab/hub'
import { startCollab, type Collab } from '../src/collab'
import { MemoryTransport } from '../src/collab/transport'
import { book, rows } from './xlsx-builder'

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

async function open(hub: CollabHub, file: Blob, name: string, latency: number) {
  const client = await DocumentClient.open(file, { worker: false })
  for (const s of client.info.xlsx!.sheets) await client.call('xlsx.load', s.index)
  await sleep(40)
  const collab = startCollab(client, { transport: new MemoryTransport(hub, () => latency), room: 'x', user: { name } })
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

const cell = async (c: DocumentClient, sheet: number, r: number, col: number) => {
  const d = await c.call('xlsx.cell', sheet, r, col)
  const shown = (await c.call('xlsx.block', sheet, r, r + 1)).rows[0]?.cells.find((x) => x[0] === col)?.[1] ?? ''
  return [d?.formula ?? '', shown]
}

const two = async () => {
  const file = await book({ sheets: [{ name: 'S1', rows: rows(Array.from({ length: 12 }, (_, r) => [r + 1, `b${r + 1}`, 60])) }, { name: 'S2', rows: rows([[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12], [13, 14, 15]]) }] })
  const hub = new CollabHub()
  // (each sees the other's edits late: what they do at the same time is done on the same state)
  const a = await open(hub, file, 'Ann', 30)
  const b = await open(hub, file, 'Bob', 30)
  await settle([a.collab, b.collab])
  return { a, b, done: () => [a, b].forEach((x) => (x.collab.close(), x.client.destroy())) }
}

describe('P2-032: a formula typed while someone moves rows or renames a sheet reads the cells meant', () => {
  test('rows inserted above: the formula follows the cell it was typed for', async () => {
    const { a, b, done } = await two()
    await Promise.all([a.client.call('xlsx.sheetOp', 0, { op: 'insertRows', at: 2, n: 1 }), b.client.call('xlsx.set', 0, [{ r: 0, c: 12, input: '=A10*2' }])])
    await settle([a.collab, b.collab])
    // A10 (10) is A11 now
    for (const x of [a, b]) expect(await cell(x.client, 0, 0, 12)).toEqual(['=A11*2', '20'])
    done()
  }, 30_000)

  test('a sheet renamed: the formula names it as it is called', async () => {
    const { a, b, done } = await two()
    await Promise.all([a.client.call('xlsx.bookOp', { op: 'rename', index: 1, name: 'Second Sheet' }), b.client.call('xlsx.set', 0, [{ r: 0, c: 13, input: '=S2!C5*10' }])])
    await settle([a.collab, b.collab])
    for (const x of [a, b]) expect(await cell(x.client, 0, 0, 13)).toEqual(["='Second Sheet'!C5*10", '150'])
    done()
  }, 30_000)

  test('a column deleted: a formula typed beside the cell it reads moves with it, and still reads it', async () => {
    const { a, b, done } = await two()
    await Promise.all([
      a.client.call('xlsx.sheetOp', 0, { op: 'deleteCols', at: 0, n: 1 }),
      b.client.call('xlsx.set', 0, [{ r: 4, c: 1, input: 'B-edit' }, { r: 4, c: 2, input: '=B5&"!"' }]),
    ])
    await settle([a.collab, b.collab])
    for (const x of [a, b]) {
      expect(await cell(x.client, 0, 4, 0)).toEqual(['', 'B-edit'])
      expect(await cell(x.client, 0, 4, 1)).toEqual(['=A5&"!"', 'B-edit!'])
    }
    done()
  }, 30_000)

  test('rows inserted in another sheet: a formula typed here that reads it by name follows', async () => {
    const { a, b, done } = await two()
    await Promise.all([a.client.call('xlsx.sheetOp', 1, { op: 'insertRows', at: 0, n: 2 }), b.client.call('xlsx.set', 0, [{ r: 1, c: 12, input: '=S2!A3+A1' }])])
    await settle([a.collab, b.collab])
    for (const x of [a, b]) expect(await cell(x.client, 0, 1, 12)).toEqual(['=S2!A5+A1', '8'])
    done()
  }, 30_000)
})

describe('P2-033: Undo takes back one\'s own change, not what someone typed over it', () => {
  test('A types, B types over it, A undoes: B\'s text stays; a cell B did not touch goes back', async () => {
    const { a, b, done } = await two()
    await a.client.call('xlsx.set', 0, [{ r: 2, c: 2, input: 'A1st' }, { r: 3, c: 2, input: 'A-only' }])
    await settle([a.collab, b.collab])
    await b.client.call('xlsx.set', 0, [{ r: 2, c: 2, input: 'B2nd' }])
    await settle([a.collab, b.collab])
    await a.client.call('xlsx.undo')
    await settle([a.collab, b.collab])
    for (const x of [a, b]) {
      expect((await cell(x.client, 0, 2, 2))[1]).toBe('B2nd')
      expect((await cell(x.client, 0, 3, 2))[1]).toBe('60')
    }
    // and Redo brings back what A typed where it is still A's to bring back
    await a.client.call('xlsx.redo')
    await settle([a.collab, b.collab])
    for (const x of [a, b]) {
      expect((await cell(x.client, 0, 2, 2))[1]).toBe('B2nd')
      expect((await cell(x.client, 0, 3, 2))[1]).toBe('A-only')
    }
    done()
  }, 30_000)
})
