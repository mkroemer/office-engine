import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'
import type { XlsxEditResult } from '../src/core/xlsx/types'

const S = (n: string) => `${import.meta.dir}/../playground/public/samples/${n}`

describe('xlsx edit results carry the undo history', () => {
  test('an edit, an undo, and a sort that changes nothing records no step', async () => {
    const engine = new Engine(() => {})
    await engine.open(Bun.file(S('sample.xlsx')) as unknown as Blob, {})
    const call = (m: string, ...args: unknown[]) => engine.call(m as never, args as never) as Promise<XlsxEditResult | null>
    expect((await call('xlsx.set', 0, [{ r: 11, c: 9, input: 'x' }]))?.history).toEqual({ undo: 1, redo: 0 })
    expect((await call('xlsx.undo'))?.history).toEqual({ undo: 0, redo: 1 })
    // one cell is always in order: the sort returns a result but leaves the history as it was
    const sorted = await call('xlsx.sheetOp', 0, { op: 'sort', range: { r0: 50, c0: 20, r1: 50, c1: 20 }, col: 20 })
    expect(sorted).not.toBeNull()
    expect(sorted?.history).toEqual({ undo: 0, redo: 1 })
  })
})
