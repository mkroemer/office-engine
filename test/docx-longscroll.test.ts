import { describe, expect, test } from 'bun:test'
import { ScrollMap } from '../src/react/docx/longscroll'

// a box 800 px high that scrolls over 12,000,000 px of spacer, and a document of 26,400,800 px: 2.2 times as far
const VIEW = 800
const MAX_TOP = 12_000_000 - VIEW
const MAX_Y = 26_400_800 - VIEW

function capped() {
  const m = new ScrollMap()
  m.fit(MAX_Y, MAX_TOP, VIEW, 1)
  return m
}

/** the reader scrolls by `d`, `n` times; a box that cannot go on is settled, as the view does */
function steps(m: ScrollMap, d: number, n: number) {
  let settled = 0
  for (let i = 0; i < n; i++) {
    const top = Math.min(MAX_TOP, Math.max(0, m.top + d))
    if (m.scrolled(top) && m.settle() !== null) settled++
  }
  return settled
}

describe('a document no longer than its box scrolls over', () => {
  test('the scroll position is the place', () => {
    const m = new ScrollMap()
    m.fit(5000, 5000, VIEW, 1)
    expect(m.capped).toBe(false)
    expect(m.scrolled(1234)).toBe(false)
    expect(m.at()).toBe(1234)
    expect(m.scrolled(5000)).toBe(false)
    expect(m.at(4000)).toBe(4000)
    expect(m.to(300)).toBe(300)
    expect(m.to(9000)).toBe(5000)
    expect(m.settle()).toBeNull()
    expect(m.off).toBe(0)
  })
})

describe('a document longer than its box scrolls over', () => {
  test('a step moves the document as far as the box: the pages stay where they are in it', () => {
    const m = capped()
    expect(m.capped).toBe(true)
    m.scrolled(5_000_000)
    const { off } = m
    const y = m.at()
    for (const d of [300, 120, -40, VIEW, -VIEW, 2000]) {
      const before = m.at()
      expect(m.scrolled(m.top + d)).toBe(false)
      expect(m.at() - before).toBe(d)
    }
    expect(m.off).toBe(off)
    expect(m.at()).toBe(y + 300 + 120 - 40 + 2000)
  })

  test('a jump goes where the thumb says: both ends, and as far along the document as the box is along its range', () => {
    const m = capped()
    m.scrolled(MAX_TOP)
    expect(m.at()).toBe(MAX_Y)
    m.scrolled(0)
    expect(m.at()).toBe(0)
    m.scrolled(MAX_TOP / 2)
    expect(m.at()).toBeCloseTo(MAX_Y / 2, 0)
    // a box that ends a little before its spacer says (`end`): still the end of the document
    m.scrolled(100)
    m.scrolled(MAX_TOP - 3, true)
    expect(m.at()).toBe(MAX_Y)
  })

  test('a long scroll that ends on a page that is drawn is the browser’s (the caret, the focus): the pages stay where they are', () => {
    const m = capped()
    m.scrolled(5_000_000)
    const y = m.at()
    const { off } = m
    // a page up, to what took the focus there
    expect(m.scrolled(5_000_000 - 167_000, false, () => true)).toBe(false)
    expect(m.off).toBe(off)
    expect(m.at()).toBe(y - 167_000)
    // the same scroll to where nothing is drawn: the thumb
    m.scrolled(5_000_000, false, () => true)
    m.scrolled(5_000_000 - 167_000, false, () => false)
    expect(m.at()).toBe(Math.round(m.place(5_000_000 - 167_000)))
    // a step does not ask
    m.scrolled(m.top + 300, false, () => {
      throw new Error('asked')
    })
  })

  test('near both ends the document scrolls pixel for pixel', () => {
    const m = capped()
    expect(m.place(500)).toBe(500)
    expect(m.home(500)).toBe(500)
    expect(m.place(999_999)).toBe(999_999)
    expect(m.place(MAX_TOP - 500)).toBe(MAX_Y - 500)
    expect(m.home(MAX_Y - 500)).toBe(MAX_TOP - 500)
    // from the top, by the wheel: no offset, as in a short document
    steps(m, 100, 50)
    expect(m.at()).toBe(5000)
    expect(m.off).toBe(0)
    expect(m.settle()).toBeNull()
  })

  test('where the box belongs for a place, and the place a jump shows, are each other’s inverse and never go back', () => {
    const m = capped()
    let last = -1
    for (let top = 0; top <= MAX_TOP; top += 37_337) {
      const y = m.place(top)
      expect(y).toBeGreaterThan(last)
      expect(y).toBeLessThanOrEqual(MAX_Y)
      expect(m.home(y)).toBeCloseTo(top, 3)
      last = y
    }
    expect(m.place(MAX_TOP)).toBe(MAX_Y)
    expect(m.home(MAX_Y)).toBe(MAX_TOP)
  })

  test('a box that rests is put where its thumb belongs, and shows the same place', () => {
    const m = capped()
    m.scrolled(6_000_000)
    expect(m.settle()).toBeNull()
    steps(m, 300, 100)
    const y = m.at()
    const top = m.settle()
    expect(top).toBe(Math.round(m.home(y)))
    // 30,000 px of the document are less than half as many of the box's range
    expect(top! - 6_000_000).toBeGreaterThan(12_000)
    expect(top! - 6_000_000).toBeLessThan(12_600)
    expect(m.at()).toBe(y)
    expect(m.top).toBe(top!)
    expect(m.settle()).toBeNull()
    // the scroll event of that position is a step of nothing
    expect(m.scrolled(top!)).toBe(false)
    expect(m.at()).toBe(y)
  })

  test('steps that reach an end of the box before the end of the document go on: the last page is reached by the wheel', () => {
    const m = capped()
    // 3,000,000 px before the end of the box's range, where the document has more than twice that left
    m.scrolled(MAX_TOP - 3_000_000)
    const left = MAX_Y - m.at()
    expect(left).toBeGreaterThan(5_000_000)
    const settled = steps(m, 500, Math.ceil(left / 500) + 5)
    expect(m.at()).toBe(MAX_Y)
    expect(m.top).toBe(MAX_TOP)
    // settled a few times on the way, not at every step
    expect(settled).toBeGreaterThan(0)
    expect(settled).toBeLessThan(12)
    // and back up to the first page
    m.scrolled(3_000_000)
    const down = m.at()
    const up = steps(m, -500, Math.ceil(down / 500) + 5)
    expect(m.at()).toBe(0)
    expect(m.top).toBe(0)
    expect(up).toBeLessThan(12)
  })

  test('a box settled at an end where the document goes on has room again', () => {
    const m = capped()
    m.scrolled(2_000_000)
    // as steps left it: the box 1500 px from its start, the document 101,500
    m.top = 1500
    m.off = 100_000
    expect(m.scrolled(0)).toBe(true)
    expect(m.at()).toBe(100_000)
    expect(m.settle()).toBe(100_000)
    expect(m.off).toBe(0)
    expect(m.at()).toBe(100_000)
  })

  test('a place asked for: a step when it is close and in reach, else where the box belongs for it', () => {
    const m = capped()
    m.scrolled(6_000_000)
    const y = m.at()
    const { off } = m
    // close by: the pages stay where they are
    expect(m.to(y + 400)).toBe(6_000_400)
    expect(m.off).toBe(off)
    expect(m.at()).toBe(y + 400)
    // far: the box where the thumb belongs
    expect(m.to(20_000_000)).toBe(Math.round(m.home(20_000_000)))
    expect(m.at()).toBe(20_000_000)
    // the scroll event of a position that was set is no jump
    const top = m.top
    m.scrolled(top)
    expect(m.at()).toBe(20_000_000)
    // out of the document
    m.to(MAX_Y + 5000)
    expect(m.at()).toBe(MAX_Y)
    expect(m.top).toBe(MAX_TOP)
    m.to(-10)
    expect(m.at()).toBe(0)
    expect(m.top).toBe(0)
    // close, but out of the box's reach: no step
    m.off = 500
    m.top = 0
    expect(m.to(300)).toBe(300)
    expect(m.off).toBe(0)
  })

  test('the position the element took is the one the place is counted from', () => {
    const m = capped()
    m.to(13_000_000.4)
    const y = m.at()
    // a fraction of a pixel: the pages stay
    const { off } = m
    m.landed(m.top + 0.5)
    expect(m.off).toBe(off)
    // the element stopped short: the place asked for is still the one shown
    m.landed(m.top - 40)
    expect(m.at()).toBeCloseTo(y + 0.5, 3)
  })

  test('another zoom: the same place of the taller or shorter pages', () => {
    const m = capped()
    m.scrolled(6_000_000)
    const y = m.at()
    expect(m.refit()).toBeNull()
    m.fit(MAX_Y * 2, MAX_TOP, VIEW, 2)
    const top = m.refit()
    expect(top).not.toBeNull()
    expect(m.at()).toBeCloseTo(y * 2, 0)
    expect(m.refit()).toBeNull()
    // down to a size at which the document fits the box's range: the place by the scroll position alone
    m.fit(MAX_Y / 4, MAX_Y / 4, VIEW, 0.25)
    expect(m.capped).toBe(false)
    expect(m.refit()).toBeCloseTo(y / 4, 0)
    expect(m.off).toBe(0)
    expect(m.at()).toBeCloseTo(y / 4, 0)
  })

  test('a document that grows past the cap while it is read stays where it is', () => {
    const m = new ScrollMap()
    m.fit(8_000_000, 8_000_000, VIEW, 1)
    m.scrolled(7_000_000)
    // more pages laid out: the same box position, the same place
    m.fit(MAX_Y, MAX_TOP, VIEW, 1)
    expect(m.capped).toBe(true)
    expect(m.refit()).toBeNull()
    expect(m.at()).toBe(7_000_000)
    // resting, the box goes to where its thumb belongs for that place
    const top = m.settle()
    expect(top).toBe(Math.round(m.home(7_000_000)))
    expect(m.at()).toBe(7_000_000)
  })

  test('pages removed: a place past the new end is put back in the document', () => {
    const m = capped()
    m.scrolled(MAX_TOP)
    m.fit(20_000_000, MAX_TOP, VIEW, 1)
    expect(m.at()).toBe(20_000_000)
    expect(m.settle()).toBe(MAX_TOP)
    expect(m.top + m.off).toBe(20_000_000)
    // shorter than the box's range: no longer capped, the place it showed (now the end)
    m.fit(9_000_000, 9_000_000, VIEW, 1)
    expect(m.refit()).toBe(9_000_000)
    expect(m.off).toBe(0)
  })
})
