import { describe, expect, test } from 'bun:test'
import { Engine } from '../src/core/worker/engine'

const sample = async () => new Blob([await Bun.file(`${import.meta.dir}/../playground/public/samples/sample.xlsx`).arrayBuffer()])

async function open(src: Blob, sheet = 0) {
  const events: { event: string; data: any }[] = []
  const engine = new Engine((event, data) => events.push({ event, data }))
  const info = await engine.open(src, {})
  const load = async (i: number) => {
    await engine.call('xlsx.load', [i])
    while (!events.some((e) => e.event === 'xlsx:progress' && e.data.sheet === i && e.data.done)) await new Promise((r) => setTimeout(r, 10))
  }
  await load(sheet)
  const cells = async (i: number, r0 = 0, r1 = 80) => {
    const b: any = await engine.call('xlsx.block', [i, r0, r1])
    return new Map<string, string>(b.rows.flatMap((row: any) => row.cells.map((c: any) => [`${row.r}:${c[0]}`, c[1]])))
  }
  const progress = (i: number) => events.filter((e) => e.event === 'xlsx:progress' && e.data.sheet === i).map((e) => e.data)
  return { engine, info, load, cells, progress }
}

describe('N1-019 duplicate sheet', () => {
  test('a copy with the cells, formats, widths and merges; own-sheet references name the copy; edits made here are copied', async () => {
    const o = await open(await sample())
    const names = o.info.xlsx!.sheets.map((s) => s.name)
    const name = names[0]
    await o.engine.call('xlsx.set', [0, [{ r: 70, c: 0, input: 'edited' }, { r: 71, c: 0, input: `='${name}'!A71&"!"` }, { r: 72, c: 0, input: '=A71' }]])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'merge', range: { r0: 80, c0: 0, r1: 81, c1: 1 } }])
    await o.engine.call('xlsx.sheetOp', [0, { op: 'colWidth', c0: 7, c1: 7, px: 150 }])
    await o.engine.call('xlsx.format', [0, { r0: 70, c0: 0, r1: 70, c1: 0 }, { b: true }])
    const res: any = await o.engine.call('xlsx.bookOp', [{ op: 'duplicate', index: 0 }])
    expect(res.sheets.map((s: any) => s.name)).toEqual([name, `${name} (2)`, ...names.slice(1)])
    await o.load(1)
    const a = await o.cells(0, 0, 100)
    const b = await o.cells(1, 0, 100)
    expect(b.get('70:0')).toBe('edited')
    // every value of the original is in the copy, but for the cell that named the original
    for (const [k, v] of a) if (k !== '71:0') expect(b.get(k)).toBe(v)
    const own: any = await o.engine.call('xlsx.cell', [1, 71, 0])
    expect(own.formula).toBe(`='${name} (2)'!A71&"!"`)
    expect(((await o.engine.call('xlsx.cell', [1, 72, 0])) as any).formula).toBe('=A71')
    const p = o.progress(1)
    expect(p.find((x) => x.colInfo)?.colInfo.some((ci: any) => ci.min === 7 && ci.w === 150)).toBe(true)
    expect(p.filter((x) => x.done).at(-1).merges).toContainEqual({ r0: 80, c0: 0, r1: 81, c1: 1 })
    // the copy is its own sheet: editing it leaves the original alone
    await o.engine.call('xlsx.set', [1, [{ r: 70, c: 0, input: 'copy only' }]])
    expect((await o.cells(0, 70, 71)).get('70:0')).toBe('edited')
    // saved and reopened: both sheets are there
    const saved = (await o.engine.call('save', [])) as Blob
    const re = await open(saved, 1)
    expect(re.info.xlsx!.sheets.map((s) => s.name)[1]).toBe(`${name} (2)`)
    expect((await re.cells(1, 70, 71)).get('70:0')).toBe('copy only')
    // a second copy gets the next number
    const again: any = await o.engine.call('xlsx.bookOp', [{ op: 'duplicate', index: 0 }])
    expect(again.sheets[1].name).toBe(`${name} (3)`)
  }, 60000)
})
