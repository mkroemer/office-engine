import { expect, test } from 'bun:test'
import { Axis } from '../src/react/xlsx/axis'

/** Bound the old non-converging search as well, so a regression fails without hanging the test runner. */
class BoundedAxis extends Axis {
  lookups = 0
  override pos(i: number): number {
    if (++this.lookups > 64) throw new Error('Axis.index did not converge within 64 lookups')
    return super.pos(i)
  }
  lookup(px: number): number {
    this.lookups = 0
    return this.index(px)
  }
}

test('axis lookup preserves large CSV counts without signed 32-bit overflow', () => {
  const axis = new BoundedAxis(20)
  for (const count of [500_000_000, 2 ** 31, 4_000_000_000, 2 ** 32 + 1, Number.MAX_SAFE_INTEGER]) {
    axis.count = count
    expect(axis.lookup(0)).toBe(0)
    expect(axis.lookup(20)).toBe(1)
    expect(axis.lookup((count - 2) * 20)).toBe(count - 2)
    expect(axis.lookup(Infinity)).toBe(count - 1)
    expect(axis.count).toBe(count)
  }
})

test('large sparse axes find custom-sized and hidden rows at their exact boundaries', () => {
  const axis = new BoundedAxis(20)
  axis.count = 4_000_000_000
  axis.set(0, 0)
  axis.set(3_000_000_000, 40)
  axis.set(3_000_000_001, 0)
  expect(axis.lookup(1)).toBe(1)
  const start = axis.pos(3_000_000_000)
  expect(start).toBe(3_000_000_000 * 20 - 20)
  expect(axis.lookup(start)).toBe(3_000_000_000)
  expect(axis.lookup(start + 39)).toBe(3_000_000_000)
  expect(axis.lookup(start + 40)).toBe(3_000_000_002)
  expect(axis.total()).toBe(axis.count * 20 - 20)
})

test('empty axes and nonfinite inputs terminate', () => {
  const axis = new BoundedAxis(20)
  for (const count of [0, -1, NaN, Infinity, 1.5]) {
    axis.count = count
    expect(axis.lookup(20)).toBe(0)
  }
  axis.count = 4_000_000_000
  expect(axis.lookup(NaN)).toBe(0)
  expect(axis.lookup(-Infinity)).toBe(0)
})
