import { afterEach, describe, expect, test } from 'bun:test'
import { convertCfb } from '../src/core/legacy'
import { mayBePlain, plainFormat } from '../src/core/plain'
import type { SlideEl, SlideModel } from '../src/core/pptx/types'
import { VISIO_LIMITS } from '../src/core/visio/convert'
import { apply, placement, roundCorners, sectionPath, subPaths, type Seg } from '../src/core/visio/geometry'
import { parseVisioXml, Resolved, Sheet, Styles, vkid, vkids, type Geometry } from '../src/core/visio/sheet'
import { readText } from '../src/core/visio/text'
import { drawingPaint, parseVisioTheme, ThemeLook } from '../src/core/visio/theme'
import { Engine } from '../src/core/worker/engine'
import { parseXml } from '../src/core/xml'
import { writeCfb } from './cfb-writer'
import { cells, geom, line, relRect, rows, sampleVisio, shape, STYLE, themeXml, VisioBuilder } from './visio-builder'

type Shape = Extract<SlideEl, { t: 'shape' }>

async function open(bytes: Uint8Array<ArrayBuffer>, name = 'drawing.vsdx') {
  const engine = new Engine(() => {})
  const info = await engine.open(new Blob([bytes]), { name })
  const slides: SlideModel[] = []
  for (let i = 0; i < (info.pptx?.count ?? 0); i++) slides.push((await engine.call('pptx.slide', [i])) as SlideModel)
  return { engine, info, slides }
}
const shapes = (s: SlideModel) => s.els.filter((e): e is Shape => e.t === 'shape')
const textOf = (e: SlideEl) => (e.t === 'shape' && e.text ? e.text.paras.map((p) => p.runs.map((r) => (r.br ? '\n' : r.text)).join('')).join('\n') : '')
const texts = (s: SlideModel) => s.els.map(textOf).filter(Boolean)
const byText = (s: SlideModel, t: string) => shapes(s).find((e) => textOf(e) === t)!
/** the drawn shape just under a text (the shape the text belongs to) */
const under = (s: SlideModel, t: string) => s.els[s.els.indexOf(byText(s, t)) - 1] as Shape
const px = (inches: number) => inches * 96
const near = (a: number, b: number, eps = 0.05) => expect(Math.abs(a - b)).toBeLessThan(eps)
const box = (e: SlideEl, x: number, y: number, w: number, h: number) => {
  near(e.box.x, px(x))
  near(e.box.y, px(y))
  near(e.box.w, px(w))
  near(e.box.h, px(h))
}

/** a drawing of one page holding `shapesXml` */
const onePage = (shapesXml: string, edit?: (b: VisioBuilder) => void, spec: Parameters<VisioBuilder['page']>[2] = {}) => {
  const b = new VisioBuilder()
  edit?.(b)
  b.page('Page-1', shapesXml, spec)
  return b.zip()
}

describe('visio: the sample drawing', () => {
  test('a slide per page, the background page under each, read-only', async () => {
    const { info, slides } = await open(await sampleVisio())
    expect(info.kind).toBe('pptx')
    expect(info.from).toBe('visio')
    // 11 × 8.5 in; the background page is not a slide of its own
    expect(info.pptx).toMatchObject({ w: px(11), h: px(8.5), count: 2 })
    for (const s of slides) {
      expect(s.bg).toEqual({ t: 'solid', c: '#ffffff' })
      expect(s.els[0]).toMatchObject({ t: 'shape', fill: { t: 'solid', c: '#1e3a8a' } })
      box(s.els[0], 0, 0, 11, 1.1)
      const title = s.els[1] as Shape
      expect(textOf(title)).toBe('Order handling')
      expect(title.text!.paras[0]).toMatchObject({ algn: 'left', runs: [{ b: true, sz: 22, color: '#ffffff' }] })
      expect(title.text!.anchor).toBe('ctr')
      expect(textOf(s.els[2])).toBe('Northwind Field Lab · process map')
      expect((s.els[2] as Shape).text!.paras[0].algn).toBe('right')
    }
  })

  test('instances of masters: the master’s look, the instance’s place, size, colors and text', async () => {
    const { slides } = await open(await sampleVisio())
    const s = slides[0]
    expect(texts(s)).toEqual(['Order handling', 'Northwind Field Lab · process map', 'Start', 'Receive order', 'In stock?', 'Pick, pack and ship', 'Back-order', 'Notify customer', 'Done', 'yes', 'no'])

    const receive = under(s, 'Receive order')
    expect(receive.fill).toEqual({ t: 'solid', c: '#dbeafe' })
    expect(receive.stroke).toMatchObject({ c: '#1d4ed8' })
    near(receive.stroke!.w, px(1.25 / 72))
    box(receive, 3.65 - 0.875, 8.5 - 5.6 - 0.375, 1.75, 0.75)
    // the Rounding cell: each corner is a curve
    expect(receive.path!.match(/C/g)).toHaveLength(4)
    expect(byText(s, 'Receive order').text!.paras[0].runs[0]).toMatchObject({ sz: 11, color: '#1f2937' })

    // wider than the master: its geometry is in fractions of the box
    box(under(s, 'Pick, pack and ship'), 8.9 - 1.1, 8.5 - 5.6 - 0.375, 2.2, 0.75)
    // its own colors over the master's
    expect(under(s, 'Back-order')).toMatchObject({ fill: { c: '#fee2e2' }, stroke: { c: '#b91c1c' } })
    // a diamond: four lines; a stadium: two lines and two half circles (ArcTo), bowing outwards
    expect(under(s, 'In stock?').path).toBe('M0.00 52.80 L81.60 105.60 L163.20 52.80 L81.60 0.00 L0.00 52.80 Z')
    const start = under(s, 'Start')
    box(start, 1.4 - 0.75, 8.5 - 5.6 - 0.3, 1.5, 0.6)
    expect(start.path!.match(/C/g)).toHaveLength(4)
    expect(byText(s, 'Start').text!.paras[0].runs[0]).toMatchObject({ b: true, color: '#14532d' })
  })

  test('connectors: 1-D shapes along their line, line ends, dashes, labels over a background', async () => {
    const { slides } = await open(await sampleVisio())
    const s = slides[0]
    const links = shapes(s).filter((e) => e.stroke?.tail)
    expect(links).toHaveLength(7)
    for (const l of links) expect(l).toMatchObject({ fill: { t: 'none' }, stroke: { tail: 'triangle' } })
    expect(links.every((l) => l.stroke!.head === undefined)).toBe(true)
    // from the terminator's edge to the process's
    box(links[0], 2.15, 8.5 - 5.6, 0.625, 0)
    // "no": down, dashed, red
    const no = links[3]
    expect(no.stroke).toMatchObject({ c: '#b91c1c' })
    expect(no.stroke!.dash).toBeTruthy()
    box(no, 6.1, 8.5 - 5.05, 0, 5.05 - 3.875)
    expect(no.path).toBe('M0.00 0.00 L0.00 112.80')
    expect(links[0].stroke!.dash).toBeUndefined()

    // the label: at the middle of its connector, level (the connector's turn is taken back by TxtAngle)
    const label = byText(s, 'no')
    expect(label.box.rot).toBeUndefined()
    near(label.box.x + label.box.w / 2, px(6.1))
    near(label.box.y + label.box.h / 2, px(8.5 - (5.05 + 3.875) / 2))
    expect(label.text!.paras[0].runs[0]).toMatchObject({ hl: '#ffffff', sz: 9, color: '#475569' })
  })

  test('shapes drawn by hand: gradient, turn, sections, group, picture, curve, text formats, theme, layers', async () => {
    const { engine, slides } = await open(await sampleVisio())
    const s = slides[1]
    // on a hidden layer
    expect(texts(s)).not.toContain('Hidden layer')

    // the first stop at the top (the angle is clockwise from left-to-right)
    const ellipse = under(s, 'Ellipse')
    expect(ellipse.fill).toEqual({ t: 'grad', css: 'linear-gradient(180deg, #ccfbf1 0%, #14b8a6 100%)' })
    expect(ellipse.path!.match(/C/g)).toHaveLength(4)
    box(ellipse, 0.5, 8.5 - 6.45 - 0.65, 2.2, 1.3)

    // turned: the path is drawn turned, the text box turns (Visio counterclockwise, the slide clockwise)
    const turned = under(s, 'Turned 20°')
    expect(turned.stroke!.dash).toBeTruthy()
    expect(turned.box.rot).toBeUndefined()
    near(turned.box.w, px(1.9 * Math.cos(Math.PI / 9) + 0.9 * Math.sin(Math.PI / 9)))
    const turnedText = byText(s, 'Turned 20°')
    near(turnedText.box.rot!, 340)
    near(turnedText.box.w, px(1.9))

    // two sections: both outlined, only the body filled
    const i = s.els.indexOf(byText(s, 'Orders'))
    expect(s.els[i - 2]).toMatchObject({ fill: { t: 'solid', c: '#e0f2fe' } })
    expect((s.els[i - 2] as Shape).stroke).toBeUndefined()
    expect(s.els[i - 1]).toMatchObject({ fill: { t: 'none' }, stroke: { c: '#0369a1' } })
    expect((s.els[i - 1] as Shape).path!.match(/M/g)).toHaveLength(2)

    // the group's own frame and title behind its members; the mirrored member's path is mirrored, its text is not
    const order = texts(s)
    expect(order.indexOf('Group')).toBeLessThan(order.indexOf('A'))
    expect(under(s, 'A').path).toBe('M0.00 76.80 L124.80 38.40 L0.00 0.00 L0.00 76.80 Z')
    expect(under(s, 'B').path).toBe('M124.80 76.80 L0.00 38.40 L124.80 0.00 L124.80 76.80 Z')
    box(under(s, 'B'), 0.5 + 2.65 - 0.65, 8.5 - (4.3 - 0.9 + 0.7 + 0.4), 1.3, 0.8)
    expect(byText(s, 'B').box).not.toHaveProperty('flipH')
    expect(byText(s, 'B').box.rot).toBeUndefined()

    const pic = s.els.find((e) => e.t === 'pic')!
    box(pic, 5.8 - 1.2, 8.5 - 4.3 - 0.8, 2.4, 1.6)
    expect(pic).toMatchObject({ target: 'ppt/media/image1.png' })
    const png = new Uint8Array(await ((await engine.call('part', ['ppt/media/image1.png'])) as Blob).arrayBuffer())
    expect([...png.subarray(1, 4)]).toEqual([0x50, 0x4e, 0x47])

    // a NURBS curve: lines along it, from its first point to its last; a line end at each end
    const curve = shapes(s).find((e) => e.stroke?.c === '#be123c')!
    expect(curve.stroke).toMatchObject({ head: 'oval', tail: 'triangle' })
    near(curve.stroke!.w, px(2 / 72))
    expect(curve.path!.startsWith('M0.00 96.00 L')).toBe(true)
    expect(curve.path!.endsWith('L249.60 0.00')).toBe(true)

    const notes = byText(s, 'Notes\nShapes keep their fills, lines and text\nThe page is Page-2, drawn 2 June 2026\nMasters, groups and themes are followed')
    expect(notes.text!.anchor).toBe('t')
    expect(notes.text!.paras[0]).toMatchObject({ algn: 'left', runs: [{ text: 'Notes', b: true, sz: 14, color: '#92400e' }] })
    expect(notes.text!.paras.map((p) => p.bullet?.text)).toEqual([undefined, '•', '•', '•'])
    expect(notes.text!.paras[1].runs).toMatchObject([{ text: 'Shapes keep their ' }, { text: 'fills, lines', b: true, i: true, color: '#b91c1c' }, { text: ' and text' }])
    expect(notes.text!.paras[2].runs[1]).toMatchObject({ text: 'Page-2', u: true })
    near(notes.text!.ins[0], px(0.15))

    // every look from the theme: the variation's first color, the thin grey line (a shade in linear light), white bold text in the theme's font
    const themed = under(s, 'Themed')
    expect(themed).toMatchObject({ fill: { t: 'solid', c: '#5b9bd5' }, stroke: { c: '#c8c8c8' } })
    const run = byText(s, 'Themed').text!.paras[0].runs[0]
    expect(run).toMatchObject({ color: '#feffff', b: true, sz: 12 })
    expect(run.font).toContain('Verdana')
  })

  test('saving gives a PowerPoint file', async () => {
    const { engine } = await open(await sampleVisio())
    const pptx = (await engine.call('save', [])) as Blob
    const again = await open(new Uint8Array(await pptx.arrayBuffer()), 'drawing.pptx')
    expect(again.info.from).toBeUndefined()
    expect(again.info.pptx?.count).toBe(2)
    expect(texts(again.slides[0])).toContain('In stock?')
  })
})

// ---------------------------------------------------------------- geometry

const g = (list: [string, Record<string, number | string>][], flags: Record<string, number> = {}): Geometry => ({
  ix: 0,
  cells: new Map(Object.entries(flags).map(([k, v]) => [k, String(v)])),
  rows: list.map(([t, c], i) => ({ ix: i + 1, t, cells: new Map(Object.entries(c).map(([k, v]) => [k, String(v)])) })),
})
const ends = (segs: Seg[]) => segs.flatMap((s) => (s.t === 'Z' ? [] : [[+s.x.toFixed(6), +s.y.toFixed(6)]]))

describe('visio: geometry rows as paths', () => {
  test('lines, relative to the box or not; a path that does not start with a move starts at the origin', () => {
    expect(sectionPath(g([['MoveTo', { X: 1, Y: 2 }], ['LineTo', { X: 3, Y: 2 }], ['RelLineTo', { X: 0.5, Y: 1 }], ['RelMoveTo', { X: 1, Y: 1 }]]), 4, 6)).toEqual([
      { t: 'M', x: 1, y: 2 },
      { t: 'L', x: 3, y: 2 },
      { t: 'L', x: 2, y: 6 },
      { t: 'M', x: 4, y: 6 },
    ])
    expect(sectionPath(g([['LineTo', { X: 1, Y: 1 }]]), 1, 1)).toEqual([{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 1, y: 1 }])
    // guides and rows of other kinds are not drawn; a row without numbers neither
    expect(sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['InfiniteLine', { X: 1, Y: 1, A: 2, B: 2 }], ['LineTo', { X: 'x', Y: 1 }]]), 1, 1)).toEqual([{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 0, y: 1 }])
  })

  test('ArcTo: a circular arc that bows A to the right of its way (to the left when negative)', () => {
    // a half circle from (0,0) to (2,0): two quarter turns, through (1,-1)
    const down = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['ArcTo', { X: 2, Y: 0, A: 1 }]]), 2, 2)
    expect(ends(down)).toEqual([[0, 0], [1, -1], [2, 0]])
    const up = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['ArcTo', { X: 2, Y: 0, A: -1 }]]), 2, 2)
    expect(ends(up)).toEqual([[0, 0], [1, 1], [2, 0]])
    // a shallow arc: one curve; its middle is A from the chord
    const [, c] = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['ArcTo', { X: 2, Y: 0, A: 0.25 }]]), 2, 2)
    if (c.t !== 'C') throw new Error('not a curve')
    near((0 + 3 * c.y1 + 3 * c.y2 + 0) / 8, -0.25, 1e-4)
    // no bow: a line
    expect(sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['ArcTo', { X: 2, Y: 0, A: 0 }]]), 2, 2)[1]).toEqual({ t: 'L', x: 2, y: 0 })
  })

  test('EllipticalArcTo: through its control point, whatever way the ellipse is turned', () => {
    // half of an ellipse 2 wide, 1 high, centered at (1,0): through its lowest point
    const flat = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['EllipticalArcTo', { X: 2, Y: 0, A: 1, B: -0.5, C: 0, D: 2 }]]), 2, 2)
    expect(ends(flat)).toEqual([[0, 0], [1, -0.5], [2, 0]])
    // the same ellipse standing (its major axis turned a quarter): from (0,0) up to (0,2) through (0.5,1)
    const tall = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['EllipticalArcTo', { X: 0, Y: 2, A: 0.5, B: 1, C: Math.PI / 2, D: 2 }]]), 2, 2)
    expect(ends(tall)).toEqual([[0, 0], [0.5, 1], [0, 2]])
    // in fractions of the box
    const rel = sectionPath(g([['RelMoveTo', { X: 0, Y: 0 }], ['RelEllipticalArcTo', { X: 1, Y: 0, A: 0.5, B: -0.25, C: 0, D: 2 }]]), 2, 2)
    expect(ends(rel)).toEqual([[0, 0], [1, -0.5], [2, 0]])
    // the three points in line: a line
    expect(sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['EllipticalArcTo', { X: 2, Y: 0, A: 1, B: 0, C: 0, D: 2 }]]), 2, 2)).toEqual([{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 2, y: 0 }])
  })

  test('Ellipse: a closed path of four curves through the ends of its axes', () => {
    const e = sectionPath(g([['Ellipse', { X: 2, Y: 1, A: 4, B: 1, C: 2, D: 2 }]]), 4, 2)
    expect(e.map((s) => s.t).join('')).toBe('MCCCCZ')
    expect(ends(e)).toEqual([[4, 1], [2, 2], [0, 1], [2, 0], [4, 1]])
    expect(subPaths(e)).toMatchObject([{ closed: true }])
  })

  test('Béziers in fractions of the box; a quadratic one as a cubic', () => {
    expect(sectionPath(g([['RelMoveTo', { X: 0, Y: 0 }], ['RelCubBezTo', { X: 1, Y: 1, A: 0.25, B: 0, C: 0.75, D: 1 }]]), 4, 2)[1]).toEqual({ t: 'C', x1: 1, y1: 0, x2: 3, y2: 2, x: 4, y: 2 })
    const q = sectionPath(g([['RelMoveTo', { X: 0, Y: 0 }], ['RelQuadBezTo', { X: 1, Y: 0, A: 0.5, B: 1 }]]), 3, 3)[1]
    if (q.t !== 'C') throw new Error('not a curve')
    // the same curve: its middle is at a quarter of the way from the chord to twice the control point's height
    near((3 * q.y1 + 3 * q.y2) / 8, 1.5, 1e-9)
    expect([q.x, q.y]).toEqual([3, 0])
  })

  test('PolylineTo: its points (fractions of the box when the type says 0), then the row’s end', () => {
    expect(ends(sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['PolylineTo', { X: 0, Y: 2, A: 'POLYLINE(0, 0, 1,0, 1,1)' }]]), 4, 2))).toEqual([[0, 0], [4, 0], [4, 2], [0, 2]])
    expect(ends(sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['PolylineTo', { X: 0, Y: 2, A: 'POLYLINE(1, 1, 1,0, 1,1)' }]]), 4, 2))).toEqual([[0, 0], [1, 0], [1, 1], [0, 2]])
    // a formula that is not one: the end alone
    expect(ends(sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['PolylineTo', { X: 1, Y: 1, A: 'nothing' }]]), 4, 2))).toEqual([[0, 0], [1, 1]])
  })

  test('NURBSTo: the curve of its control points, knots and weights', () => {
    // four control points, knots 0,0,0,0,1…: a cubic Bézier from (0,0) to (3,0) with controls (1,2) and (2,2)
    const n = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['NURBSTo', { X: 3, Y: 0, A: 0, B: 1, C: 0, D: 1, E: 'NURBS(1, 3, 1, 1, 1,2,0,1, 2,2,0,1)' }]]), 3, 2)
    expect(n).toHaveLength(33)
    const mid = n[16] as { x: number; y: number }
    near(mid.x, 1.5, 1e-9)
    near(mid.y, 1.5, 1e-9)
    expect(n[32]).toEqual({ t: 'L', x: 3, y: 0 })
    // control points in fractions of the box
    const rel = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['NURBSTo', { X: 3, Y: 0, A: 0, B: 1, C: 0, D: 1, E: 'NURBS(1, 3, 0, 0, 0.3333333333333333,1,0,1, 0.6666666666666666,1,0,1)' }]]), 3, 2)
    near((rel[16] as { y: number }).y, 1.5, 1e-9)
    // a heavier control point pulls the curve to it
    const heavy = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['NURBSTo', { X: 3, Y: 0, A: 0, B: 1, C: 0, D: 1, E: 'NURBS(1, 3, 1, 1, 1,2,0,5, 2,2,0,5)' }]]), 3, 2)
    expect((heavy[16] as { y: number }).y).toBeGreaterThan(1.8)
    // nothing to draw from: a line to the end
    expect(sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['NURBSTo', { X: 3, Y: 0, E: 'NURBS()' }]]), 3, 2).at(-1)).toEqual({ t: 'L', x: 3, y: 0 })
  })

  test('SplineStart / SplineKnot: one curve from where the path is to the last knot', () => {
    const s = sectionPath(g([['MoveTo', { X: 0, Y: 0 }], ['SplineStart', { X: 1, Y: 2, A: 0, B: 0, C: 1, D: 3 }], ['SplineKnot', { X: 2, Y: 2, A: 0 }], ['SplineKnot', { X: 3, Y: 0, A: 0 }], ['LineTo', { X: 4, Y: 0 }]]), 4, 2)
    expect(s[0]).toEqual({ t: 'M', x: 0, y: 0 })
    near((s[16] as { x: number }).x, 1.5, 1e-9)
    near((s[16] as { y: number }).y, 1.5, 1e-9)
    expect(s.at(-2)).toEqual({ t: 'L', x: 3, y: 0 })
    expect(s.at(-1)).toEqual({ t: 'L', x: 4, y: 0 })
  })

  test('a part of a path is closed when it ends where it starts', () => {
    const square: Seg[] = [{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 1, y: 0 }, { t: 'L', x: 1, y: 1 }, { t: 'L', x: 0, y: 1 }, { t: 'L', x: 0, y: 1e-9 }]
    expect(subPaths(square)).toMatchObject([{ closed: true }])
    expect(subPaths(square.slice(0, 4))).toMatchObject([{ closed: false }])
    // there and back along one line is not a closed shape; a move alone is nothing
    expect(subPaths([{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 1, y: 0 }, { t: 'M', x: 5, y: 5 }])).toMatchObject([{ closed: false }])
    expect(subPaths([...square, { t: 'M', x: 3, y: 3 }, { t: 'L', x: 4, y: 4 }]).map((p) => p.closed)).toEqual([true, false])
  })

  test('rounding: corners between two lines become curves, no deeper than half a side', () => {
    const [square] = subPaths([{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 2, y: 0 }, { t: 'L', x: 2, y: 2 }, { t: 'L', x: 0, y: 2 }, { t: 'L', x: 0, y: 0 }])
    const r = roundCorners(square, 0.5)
    expect(r.closed).toBe(true)
    expect(r.segs.map((s) => s.t).join('')).toBe('MLCLCLCLC')
    expect(r.segs[0]).toEqual({ t: 'M', x: 0.5, y: 0 })
    expect(r.segs[1]).toEqual({ t: 'L', x: 1.5, y: 0 })
    expect(r.segs.at(-1)).toMatchObject({ x: 0.5, y: 0 })
    // a radius larger than the shape: half of each side
    expect(roundCorners(square, 9).segs[0]).toEqual({ t: 'M', x: 1, y: 0 })
    // an open path keeps its ends; no radius, no change
    const [open] = subPaths([{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 2, y: 0 }, { t: 'L', x: 2, y: 2 }])
    expect(roundCorners(open, 0.5).segs.map((s) => s.t).join('')).toBe('MLCL')
    expect(roundCorners(open, 0)).toBe(open)
  })

  test('placement: the local pin lands on the pin, after the flips and the turn', () => {
    const p = (m: ReturnType<typeof placement>, x: number, y: number) => apply(m, x, y).map((v) => +v.toFixed(9))
    expect(p(placement(5, 5, 1, 1, 0, false, false), 0, 0)).toEqual([4, 4])
    // a quarter turn counterclockwise about the pin
    expect(p(placement(5, 5, 1, 1, Math.PI / 2, false, false), 2, 1)).toEqual([5, 6])
    expect(p(placement(5, 5, 1, 1, 0, true, false), 0, 0)).toEqual([6, 4])
    expect(p(placement(5, 5, 1, 1, 0, false, true), 0, 0)).toEqual([4, 6])
    // flipped, then turned
    expect(p(placement(5, 5, 1, 1, Math.PI / 2, true, false), 0, 1)).toEqual([5, 6])
  })
})

// ---------------------------------------------------------------- sheets

describe('visio: what a shape inherits', () => {
  const sheetOf = (xml: string) => new Sheet(parseVisioXml(xml))

  test('cells: the shape, its master, then the styles of each kind and their parents', () => {
    const styles = new Styles()
    styles.add('0', parseVisioXml(`<StyleSheet ID='0'>${cells({ LineWeight: 0.01, FillForegnd: 1, LeftMargin: 0, Rounding: 0 })}</StyleSheet>`))
    styles.add('1', parseVisioXml(`<StyleSheet ID='1' LineStyle='0' FillStyle='0' TextStyle='0'>${cells({ LineWeight: 0.02, LeftMargin: 0.1 })}</StyleSheet>`))
    styles.add('2', parseVisioXml(`<StyleSheet ID='2' LineStyle='0' FillStyle='0' TextStyle='0'>${cells({ FillForegnd: '#ff0000', LineWeight: 0.5 })}</StyleSheet>`))
    const own = sheetOf(`<Shape>${cells({ PinX: 3 })}</Shape>`)
    const master = sheetOf(`<Shape>${cells({ PinX: 1, Width: 2, Rounding: 0.3 })}</Shape>`)
    // the line cells from style 1, the fill cells from style 2: its line weight is not looked at
    const r = new Resolved([own, master], styles, { line: '1', fill: '2', text: '1' })
    expect(r.num('PinX')).toBe(3)
    expect(r.num('Width')).toBe(2)
    expect(r.num('Rounding')).toBe(0.3)
    expect(r.num('LineWeight')).toBe(0.02)
    expect(r.cell('FillForegnd')).toBe('#ff0000')
    expect(r.num('LeftMargin')).toBe(0.1)
    expect(r.cell('Nothing')).toBeUndefined()
    expect(r.num('Nothing', 7)).toBe(7)
    // a themed value is not a number
    expect(new Resolved([sheetOf(`<Shape>${cells({ LineWeight: 'Themed' })}</Shape>`)], styles, { line: '1', fill: '1', text: '1' }).num('LineWeight', 9)).toBe(9)
    // styles that name each other end
    styles.add('8', parseVisioXml("<StyleSheet ID='8' LineStyle='9'/>"))
    styles.add('9', parseVisioXml("<StyleSheet ID='9' LineStyle='8'/>"))
    expect(styles.chain('8', 'line')).toHaveLength(2)
    expect(styles.chain('none', 'line')).toEqual([])
  })

  test('rows: by their number over the master’s; the styles’ one row stands for every row', () => {
    const styles = new Styles()
    styles.add('0', parseVisioXml(`<StyleSheet ID='0'>${rows('Character', [{ Font: 'Calibri', Size: 0.1, Color: 0 }])}</StyleSheet>`))
    const own = sheetOf(`<Shape>${rows('Character', [{ Size: 0.2 }, { Color: '#00ff00' }])}</Shape>`)
    const master = sheetOf(`<Shape>${rows('Character', [{ Color: '#0000ff' }])}</Shape>`)
    const r = new Resolved([own, master], styles, { line: '0', fill: '0', text: '0' })
    expect(r.rowKeys('Character')).toEqual(['0', '1'])
    expect(r.rowCell('Character', '0', 'Size')).toBe('0.2')
    expect(r.rowCell('Character', '0', 'Color')).toBe('#0000ff')
    expect(r.rowCell('Character', '1', 'Color')).toBe('#00ff00')
    expect(r.rowCell('Character', '1', 'Size')).toBe('0.1')
    expect(r.rowCell('Character', '5', 'Font')).toBe('Calibri')
  })

  test('geometry: rows merged cell by cell with the master’s; deleted rows and sections are gone', () => {
    const master = sheetOf(
      `<Shape>${geom(0, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: 1, Y: 0 }], ['LineTo', { X: 1, Y: 1 }]], { NoFill: 1 })}${geom(1, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: 9, Y: 9 }]])}${geom(2, [['Ellipse', { X: 1, Y: 1 }]])}</Shape>`,
    )
    // row 2: only X is the instance's; row 3 deleted; section 1 deleted; a section of its own
    const own = sheetOf(`<Shape>${geom(0, [['MoveTo', {}], ['LineTo', { X: 5 }], { del: true }])}<Section N='Geometry' IX='1' Del='1'/>${geom(3, [['MoveTo', { X: 7, Y: 7 }]])}</Shape>`)
    const sections = new Resolved([own, master], new Styles(), { line: undefined, fill: undefined, text: undefined }).geometry()
    expect(sections.map((s) => s.ix)).toEqual([0, 2, 3])
    expect(sections[0].cells.get('NoFill')).toBe('1')
    expect(sections[0].rows.map((r) => [r.t, Object.fromEntries(r.cells)])).toEqual([
      ['MoveTo', { X: '0', Y: '0' }],
      ['LineTo', { X: '5', Y: '0' }],
    ])
    expect(sections[1].rows[0].t).toBe('Ellipse')
  })

  test('an instance draws its master’s members: those it lists, or all of them when it lists none', async () => {
    const build = (members?: string[]) =>
      onePage(shape({ id: 1, type: 'Group', master: 1, cells: { PinX: 4, PinY: 4 }, shapes: members }), (b) => {
        b.master(
          'Pair',
          shape({
            id: 5, type: 'Group', style: STYLE.NORMAL, box: [1, 1, 2, 1], cells: { FillPattern: 0, LinePattern: 0 },
            shapes: [
              shape({ id: 6, style: STYLE.NORMAL, box: [0.5, 0.5, 1, 1], cells: { FillForegnd: '#ff0000' }, sections: relRect(), text: 'left\n' }),
              shape({ id: 7, style: STYLE.NORMAL, box: [1.5, 0.5, 1, 1], cells: { FillForegnd: '#0000ff' }, sections: relRect(), text: 'right\n' }),
            ],
          }),
        )
      })
    // no member listed: both, as the master has them
    const all = (await open(await build())).slides[0]
    expect(texts(all)).toEqual(['left', 'right'])
    expect(under(all, 'left').fill).toEqual({ t: 'solid', c: '#ff0000' })
    // the group's pin (its center) at (4,4): the left member's box starts a group width / 2 to the left
    box(under(all, 'left'), 3, 8.5 - 4.5, 1, 1)
    box(under(all, 'right'), 4, 8.5 - 4.5, 1, 1)
    // one listed, with its own fill and text: the other was deleted from the instance
    const one = (await open(await build([shape({ id: 2, masterShape: 7, cells: { FillForegnd: '#00ff00' }, text: 'only\n' })]))).slides[0]
    expect(texts(one)).toEqual(['only'])
    expect(under(one, 'only').fill).toEqual({ t: 'solid', c: '#00ff00' })
    box(under(one, 'only'), 4, 8.5 - 4.5, 1, 1)
  })

  test('not drawn: guides, sections marked NoShow, the own drawing of a group that hides it, text marked hidden', async () => {
    const { slides } = await open(
      await onePage(
        shape({ id: 1, type: 'Guide', style: STYLE.NORMAL, box: [1, 1, 1, 1], sections: relRect(), text: 'guide\n' }) +
          shape({ id: 2, style: STYLE.NORMAL, box: [3, 3, 1, 1], sections: relRect(0, { NoShow: 1 }) + geom(1, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: 1, Y: 1 }]]), text: 'shown\n' }) +
          shape({ id: 3, type: 'Group', style: STYLE.NORMAL, box: [6, 3, 2, 2], cells: { DisplayMode: 0 }, sections: relRect(), text: 'group\n', shapes: [shape({ id: 4, style: STYLE.NORMAL, box: [1, 1, 1, 1], sections: relRect(), text: 'member\n' })] }) +
          shape({ id: 5, style: STYLE.NORMAL, box: [9, 3, 1, 1], cells: { HideText: 1, FillPattern: 0 }, sections: relRect(), text: 'hidden\n' }),
      ),
    )
    expect(texts(slides[0])).toEqual(['shown', 'member'])
    // the diagonal only: the NoShow rectangle is not there
    expect(under(slides[0], 'shown').path).toBe('M0.00 96.00 L96.00 0.00')
    // the hidden text's shape is still outlined
    expect(shapes(slides[0]).at(-1)).toMatchObject({ fill: { t: 'none' }, stroke: { c: '#000000' } })
    expect(shapes(slides[0]).at(-1)!.text).toBeUndefined()
  })
})

// ---------------------------------------------------------------- looks

describe('visio: fills, lines and text', () => {
  test('colors: #rrggbb, a number of the document’s palette or of the built-in one; see-through fills', async () => {
    const { slides } = await open(
      await onePage(
        shape({ id: 1, style: STYLE.NORMAL, box: [1, 1, 1, 1], cells: { FillForegnd: 2, LineColor: 4 }, sections: relRect(), text: 'builtin\n' }) +
          shape({ id: 2, style: STYLE.NORMAL, box: [3, 1, 1, 1], cells: { FillForegnd: 24, FillForegndTrans: 0.5, LineColor: 99 }, sections: relRect(), text: 'palette\n' }) +
          shape({ id: 3, style: STYLE.NORMAL, box: [5, 1, 1, 1], cells: { FillForegnd: 'RGB(1,2,3)', LinePattern: 0 }, sections: relRect(), text: 'rgb\n' }) +
          shape({ id: 4, style: STYLE.NORMAL, box: [7, 1, 1, 1], cells: { FillPattern: 0, LineColor: '#123456', LineColorTrans: 0.25, LineWeight: 0 }, sections: relRect(), text: 'none\n' }),
        (b) => (b.colors = "<ColorEntry IX='24' RGB='#7F7F7F'/>"),
      ),
    )
    const s = slides[0]
    expect(under(s, 'builtin')).toMatchObject({ fill: { t: 'solid', c: '#ff0000' }, stroke: { c: '#0000ff' } })
    // an unknown number: black
    expect(under(s, 'palette')).toMatchObject({ fill: { t: 'solid', c: 'rgba(127,127,127,0.5)' }, stroke: { c: '#000000' } })
    expect(under(s, 'rgb').fill).toEqual({ t: 'solid', c: '#010203' })
    expect(under(s, 'rgb').stroke).toBeUndefined()
    // no fill; a weight of 0 is still a line, the thinnest
    expect(under(s, 'none')).toMatchObject({ fill: { t: 'none' }, stroke: { c: 'rgba(18,52,86,0.75)', w: 0.5 } })
  })

  test('fill patterns: the gradients of older drawings, hatches as a mix; gradient sections', async () => {
    const two = { FillForegnd: '#000000', FillBkgnd: '#ffffff' }
    const { slides } = await open(
      await onePage(
        shape({ id: 1, style: STYLE.NORMAL, box: [1, 1, 1, 1], cells: { ...two, FillPattern: 25 }, sections: relRect(), text: 'across\n' }) +
          shape({ id: 2, style: STYLE.NORMAL, box: [3, 1, 1, 1], cells: { ...two, FillPattern: 29 }, sections: relRect(), text: 'axial\n' }) +
          shape({ id: 3, style: STYLE.NORMAL, box: [5, 1, 1, 1], cells: { ...two, FillPattern: 40 }, sections: relRect(), text: 'radial\n' }) +
          shape({ id: 4, style: STYLE.NORMAL, box: [7, 1, 1, 1], cells: { ...two, FillPattern: 6 }, sections: relRect(), text: 'hatch\n' }) +
          shape({
            id: 5, style: STYLE.NORMAL, box: [9, 1, 1, 1], cells: { FillGradientEnabled: 1, FillGradientDir: 0, FillGradientAngle: 0 },
            sections: rows('FillGradient', [{ GradientStopColor: '#ff0000', GradientStopColorTrans: 0, GradientStopPosition: 0 }, { GradientStopColor: '#00ff00', GradientStopColorTrans: 0.5, GradientStopPosition: 0.4 }, { GradientStopColor: '#0000ff', GradientStopColorTrans: 0, GradientStopPosition: 1 }]) + relRect(),
            text: 'stops\n',
          }) +
          shape({ id: 6, style: STYLE.NORMAL, box: [9, 3, 1, 1], cells: { FillGradientEnabled: 1, FillGradientDir: 10 }, sections: rows('FillGradient', [{ GradientStopColor: '#ff0000', GradientStopPosition: 0 }, { GradientStopColor: '#0000ff', GradientStopPosition: 1 }]) + relRect(), text: 'from center\n' }) +
          // enabled without stops: the plain fill
          shape({ id: 7, style: STYLE.NORMAL, box: [7, 3, 1, 1], cells: { FillGradientEnabled: 1, FillForegnd: '#abcdef' }, sections: relRect(), text: 'no stops\n' }),
      ),
    )
    const s = slides[0]
    expect(under(s, 'across').fill).toEqual({ t: 'grad', css: 'linear-gradient(90deg, #000000 0%, #ffffff 100%)' })
    expect(under(s, 'axial').fill).toEqual({ t: 'grad', css: 'linear-gradient(180deg, #ffffff 0%, #000000 50%, #ffffff 100%)' })
    expect(under(s, 'radial').fill).toEqual({ t: 'grad', css: 'radial-gradient(circle, #000000 0%, #ffffff 100%)' })
    expect(under(s, 'hatch').fill).toEqual({ t: 'solid', c: '#a6a6a6' })
    expect(under(s, 'stops').fill).toEqual({ t: 'grad', css: 'linear-gradient(90deg, #ff0000 0%, rgba(0,255,0,0.5) 40%, #0000ff 100%)' })
    expect(under(s, 'from center').fill).toEqual({ t: 'grad', css: 'radial-gradient(circle, #ff0000 0%, #0000ff 100%)' })
    expect(under(s, 'no stops').fill).toEqual({ t: 'solid', c: '#abcdef' })
  })

  test('lines: patterns as dashes, line ends only on open paths, their kinds and sizes', async () => {
    const { engine, slides } = await open(
      await onePage(
        line(1, 1, 1, 3, 1, { cells: { LinePattern: 3, BeginArrow: 13, EndArrow: 1, BeginArrowSize: 0, EndArrowSize: 6 } }) +
          line(2, 1, 2, 3, 2, { cells: { LinePattern: 4, EndArrow: 22 } }) +
          line(3, 1, 3, 3, 3, { cells: { LinePattern: 16, EndArrow: 5 } }) +
          // a closed shape has no ends to put them on
          shape({ id: 4, style: STYLE.NORMAL, box: [6, 2, 1, 1], cells: { EndArrow: 4, BeginArrow: 4, LinePattern: 99 }, sections: relRect(), text: 'closed\n' }),
      ),
    )
    const [a, b, c] = shapes(slides[0])
    expect(a.stroke).toMatchObject({ head: 'triangle', tail: 'arrow' })
    expect(b.stroke).toMatchObject({ tail: 'diamond' })
    expect(b.stroke!.head).toBeUndefined()
    expect(c.stroke).toMatchObject({ tail: 'stealth' })
    // three different dashes
    expect(new Set([a.stroke!.dash, b.stroke!.dash, c.stroke!.dash]).size).toBe(3)
    const closed = under(slides[0], 'closed')
    expect(closed.stroke!.head).toBeUndefined()
    expect(closed.stroke!.tail).toBeUndefined()
    expect(closed.stroke!.dash).toBeTruthy()
    // the sizes are written for PowerPoint
    const xml = await ((await engine.call('part', ['ppt/slides/slide1.xml'])) as Blob).text()
    expect(xml).toContain('<a:headEnd type="triangle" w="sm" len="sm"/><a:tailEnd type="arrow" w="lg" len="lg"/>')
    expect(xml).toContain('<a:prstDash val="sysDot"/>')
  })

  test('text: its own block in the shape (place, size, turn), margins, alignment, spacing, cases, positions', async () => {
    const { engine, slides } = await open(
      await onePage(
        // the text under the shape, wider than it, turned a quarter
        shape({
          id: 1, style: STYLE.NORMAL, box: [2, 6, 1, 1],
          cells: { TxtPinX: 0.5, TxtPinY: -0.5, TxtWidth: 2, TxtHeight: 0.4, TxtLocPinX: 1, TxtLocPinY: 0.2, TxtAngle: Math.PI / 2, LeftMargin: 0.1, TopMargin: 0.02, RightMargin: 0.2, BottomMargin: 0.03, VerticalAlign: 2 },
          sections: rows('Character', [{ Size: 10 / 72, Case: 1, Pos: 1, Strikethru: 1 }, { Size: 8 / 72, Pos: 2, Style: 8, DblUnderline: 1, ColorTrans: 0.5, Color: '#ff0000', Font: 'Georgia' }]) + rows('Paragraph', [{ HorzAlign: 2, SpLine: -2.4, SpBefore: 6 / 72, SpAfter: 3 / 72, IndLeft: 0.25, IndFirst: 0.125 }, { HorzAlign: 3, SpLine: 0.25 }]) + relRect(),
          text: "<cp IX='0'/><pp IX='0'/>one two\n<pp IX='1'/><cp IX='1'/>three\n",
        }) +
          // no text at all: no text box
          shape({ id: 2, style: STYLE.NORMAL, box: [6, 6, 1, 1], sections: relRect(), text: '\n' }),
      ),
    )
    const t = shapes(slides[0]).find((e) => e.text)!
    expect(shapes(slides[0]).filter((e) => e.text)).toHaveLength(1)
    // centered at the text pin: (2 - 0.5 + 0.5, 6 - 0.5 - 0.5) on the page
    near(t.box.x + t.box.w / 2, px(2))
    near(t.box.y + t.box.h / 2, px(8.5 - 5))
    near(t.box.w, px(2))
    near(t.box.h, px(0.4))
    near(t.box.rot!, 270)
    expect(t.text!.anchor).toBe('b')
    expect(t.text!.ins.map((v) => +(v / 96).toFixed(3))).toEqual([0.1, 0.02, 0.2, 0.03])
    const [p1, p2] = t.text!.paras
    expect(p1.algn).toBe('right')
    // the line break stays in the paragraph
    expect(p1.runs.map((r) => (r.br ? '<br>' : r.text))).toEqual(['one', '<br>', 'two'])
    expect(p1.runs[0]).toMatchObject({ sz: 10, caps: true, strike: true, baseline: 30 })
    // 240 % of the font size is twice single spacing
    expect(p1.lnPct).toBe(2)
    near(p1.spcBef, 8)
    near(p1.spcAft, 4)
    near(p1.marL, px(0.25))
    near(p1.indent, px(0.125))
    expect(p2.algn).toBe('justify')
    near(p2.lnPx!, px(0.25))
    expect(p2.runs[0]).toMatchObject({ text: 'three', sz: 8, u: true, baseline: -25, color: 'rgba(255,0,0,0.5)' })
    expect(p2.runs[0].font).toContain('Georgia')
    const xml = await ((await engine.call('part', ['ppt/slides/slide1.xml'])) as Blob).text()
    expect(xml).toContain('cap="small"')
    expect(xml).toContain('u="dbl"')
  })

  test('text blocks too narrow: laid across a line; widened for a word that barely fits', async () => {
    const { slides } = await open(
      await onePage(
        // a vertical line: the block is as wide as the line is thick
        shape({ id: 1, style: STYLE.NORMAL, cells: { PinX: 2, PinY: 4, Width: 0, Height: 2, LocPinX: 0, LocPinY: 1 }, sections: geom(0, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: 0, Y: 2 }]]), text: 'along the line\n' }) +
          shape({ id: 2, style: STYLE.NORMAL, box: [6, 4, 0.8, 0.5], sections: rows('Character', [{ Size: 12 / 72 }]) + relRect(), text: 'Applicants\n' }) +
          shape({ id: 3, style: STYLE.NORMAL, box: [9, 4, 0.8, 0.5], sections: rows('Character', [{ Size: 12 / 72 }]) + relRect(), text: 'a b c d e f g\n' }),
      ),
    )
    const along = byText(slides[0], 'along the line')
    expect(along.text!.wrap).toBe(false)
    expect(along.box.w).toBeGreaterThan(px(0.9))
    near(along.box.x + along.box.w / 2, px(2))
    // ten letters at 12 pt need more than 0.8 in: the box grows about its center, not past half as wide again
    const word = byText(slides[0], 'Applicants')
    expect(word.text!.wrap).toBe(true)
    expect(word.box.w).toBeGreaterThan(px(0.9))
    expect(word.box.w).toBeLessThanOrEqual(px(1.2) + 0.01)
    near(word.box.x + word.box.w / 2, px(6))
    // short words wrap where the drawing wraps them
    near(byText(slides[0], 'a b c d e f g').box.w, px(0.8))
  })

  test('the text of a shape is read with its markers: character and paragraph rows, fields, line breaks', () => {
    const el = vkid(parseVisioXml("<Shape><Text><cp IX='2'/><pp IX='1'/>A <fld IX='0'>12</fld> B\r\n<cp IX='3'/>C D\n\n<tp IX='0'/>E\n</Text></Shape>"), 'Text')
    expect(readText(el)).toEqual([
      { pp: '1', runs: [{ text: 'A 12 B', cp: '2' }] },
      { pp: '1', runs: [{ text: 'C', cp: '3' }, { text: '', cp: '3', br: true }, { text: 'D', cp: '3' }] },
      { pp: '1', runs: [] },
      { pp: '1', runs: [{ text: 'E', cp: '3' }] },
    ])
    expect(readText(vkid(parseVisioXml('<Shape><Text>\n</Text></Shape>'), 'Text'))).toEqual([])
    expect(readText(undefined)).toEqual([])
    // other parts of a shape keep their text out of the way of the elements
    expect(vkids(parseVisioXml("<Shape><Cell N='A' V='1'/> <Cell N='B' V='2'/></Shape>"), 'Cell')).toHaveLength(2)
  })

  test('pictures: the part of a larger picture a shape shows; a picture of the master; a missing one is left out', async () => {
    const b = new VisioBuilder()
    b.image('a.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47]))
    b.image('m.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1]))
    const stamp = b.master('Stamp', shape({ id: 5, type: 'Foreign', style: STYLE.NORMAL, box: [1, 1, 1, 1], cells: { FillPattern: 0, LinePattern: 0 }, foreign: 'rId1' }), { images: ['m.png'] })
    b.page(
      'Page-1',
      // the picture is 4 × 2, the shape shows the 2 × 1 at its middle; flipped
      shape({ id: 1, type: 'Foreign', style: STYLE.NORMAL, box: [3, 3, 2, 1], cells: { ImgOffsetX: -1, ImgOffsetY: -0.5, ImgWidth: 4, ImgHeight: 2, FlipX: 1, Angle: Math.PI / 2, FillPattern: 0, LinePattern: 0 }, foreign: 'rId1' }) +
        shape({ id: 2, master: stamp, cells: { PinX: 7, PinY: 3 } }) +
        shape({ id: 3, type: 'Foreign', style: STYLE.NORMAL, box: [9, 3, 1, 1], cells: { FillPattern: 0, LinePattern: 0 }, foreign: 'rId9' }),
      { images: ['a.png'] },
    )
    const { engine, slides } = await open(await b.zip())
    const pics = slides[0].els.filter((e) => e.t === 'pic')
    expect(pics).toHaveLength(2)
    expect(pics[0]).toMatchObject({ target: 'ppt/media/image1.png', crop: [0.25, 0.25, 0.25, 0.25], box: { flipH: true } })
    near(pics[0].box.w, px(2))
    near(pics[0].box.x + pics[0].box.w / 2, px(3))
    // flipped, then a quarter turn counterclockwise: on the slide, three quarters clockwise
    near(pics[0].box.rot!, 270)
    // the master's picture comes from the master's part
    expect(pics[1]).toMatchObject({ target: 'ppt/media/image2.png' })
    box(pics[1], 6.5, 8.5 - 3.5, 1, 1)
    expect(((await engine.call('part', ['ppt/media/image2.png'])) as Blob).size).toBe(5)
  })
})

// ---------------------------------------------------------------- themes

describe('visio: themes', () => {
  const theme = parseVisioTheme(themeXml(33))
  const el = (xml: string) => parseXml(`<x xmlns:a="a">${xml}</x>`).c[0]

  test('the theme part: colors, the background, variations, styles, Visio’s own lists', () => {
    expect(theme.id).toBe(33)
    expect(theme.colors).toMatchObject({ dk1: '#000000', lt1: '#feffff', accent1: '#5b9bd5', accent6: '#70ad47', bkgnd: '#ffffff' })
    expect(theme.variations).toEqual([['#5b9bd5', '#759fcc', '#41719c', '#bdd0e9', '#ed7d31', '#70ad47', '#fec000']])
    expect(theme.font).toBe('Verdana')
    expect([theme.fills.length, theme.lines.length, theme.connectorLines.length, theme.fonts.length, theme.connectorFonts.length]).toEqual([3, 3, 1, 3, 1])
    expect(theme.varStyles[0]).toEqual([{ fill: 2, line: 2, font: 2 }, { fill: 1, line: 1, font: 1 }, { fill: 3, line: 3, font: 3 }, { fill: 2, line: 1, font: 2 }])
    expect(theme.connectorLineEx[0]).toMatchObject({ end: 4, pattern: 2 })
  })

  test('DrawingML colors: tints and shades mix in linear light, as Office shows them', () => {
    // values Visio itself writes for these in a drawing with this theme
    expect(drawingPaint(el('<a:schemeClr val="lt1"><a:shade val="58000"/></a:schemeClr>'), theme, undefined)).toEqual({ hex: '#c8c8c8', trans: 0 })
    expect(drawingPaint(el('<a:schemeClr val="phClr"><a:shade val="75000"/></a:schemeClr>'), theme, '#5b9bd5')).toEqual({ hex: '#4f88bb', trans: 0 })
    expect(drawingPaint(el('<a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr>'), theme, '#000000')!.hex).toBe('#bcbcbc')
    expect(drawingPaint(el('<a:srgbClr val="FF0000"><a:alpha val="40000"/></a:srgbClr>'), theme, undefined)).toEqual({ hex: '#ff0000', trans: 0.6 })
    expect(drawingPaint(el('<a:schemeClr val="tx1"/>'), theme, undefined)!.hex).toBe('#000000')
    expect(drawingPaint(el('<a:sysClr val="window"/>'), theme, undefined)!.hex).toBe('#ffffff')
    // no placeholder color, an unknown name, something that is not a color
    expect(drawingPaint(el('<a:schemeClr val="phClr"/>'), theme, undefined)).toBeUndefined()
    expect(drawingPaint(el('<a:schemeClr val="constructor"/>'), theme, undefined)).toBeUndefined()
    expect(drawingPaint(el('<a:srgbClr val="fff url(x)"/>'), theme, undefined)).toBeUndefined()
  })

  test('what a shape’s quick style cells pick: a color, then the style it is put through', () => {
    const look = new ThemeLook(theme, theme, 0, 0)
    expect([0, 1, 2, 7, 8, 100, 104].map((i) => look.color(i))).toEqual(['#000000', '#feffff', '#5b9bd5', '#70ad47', '#ffffff', '#5b9bd5', '#ed7d31'])
    expect(look.color(55)).toBeUndefined()
    // style 3: the gradient, from a tint of the color to the color
    expect(look.fill(3, 3)).toMatchObject({ pattern: 1, angle: 90, stops: [{ pos: 0 }, { pos: 1, paint: { hex: '#ed7d31' } }] })
    // style 1 is white whatever the color; no style: the color itself
    expect(look.fill(1, 3)!.color.hex).toBe('#ffffff')
    expect(look.fill(0, 3)).toEqual({ color: { hex: '#ed7d31', trans: 0 }, pattern: 1 })
    // 100-103: the styles of the page's style variation
    expect(look.fill(102, 100)!.stops).toHaveLength(2)
    expect(look.line(102, 100, false)).toMatchObject({ color: { hex: '#4f88bb' }, pattern: 2 })
    near(look.line(102, 100, false)!.weight!, 1.5 / 72, 1e-9)
    // connectors have styles of their own: a dashed line with an end
    expect(look.line(1, 2, true)).toMatchObject({ color: { hex: '#5b9bd5' }, pattern: 2, ex: { end: 4 } })
    near(look.line(1, 2, true)!.weight!, 1 / 72, 1e-9)
    expect(look.text(2, 2, false)).toEqual({ color: { hex: '#feffff', trans: 0 }, style: 1 })
    expect(look.text(1, 2, true)).toMatchObject({ color: { hex: '#4f88bb' } })
    expect(look.background).toBe('#ffffff')
  })

  test('a themed connector takes the connector styles; themed text over its own fill stays readable', async () => {
    const themedText = rows('Character', [{ Size: 10 / 72 }])
    const { slides } = await open(
      await onePage(
        line(1, 1, 1, 3, 1, { style: STYLE.THEMED, cells: { QuickStyleLineMatrix: 1, QuickStyleLineColor: 3, EndArrow: ['Themed', 'THEMEVAL()'] } }) +
          // font style 1 gives the color itself: the same as the fill, so black or white takes its place
          shape({ id: 2, style: STYLE.THEMED, box: [6, 2, 2, 1], cells: { QuickStyleFillMatrix: 2, QuickStyleFontMatrix: 1, QuickStyleVariation: 8 }, sections: themedText + relRect(), text: 'dark fill\n' }) +
          shape({ id: 3, style: STYLE.THEMED, box: [6, 4, 2, 1], cells: { QuickStyleFillMatrix: 2, QuickStyleFillColor: 1, QuickStyleFontMatrix: 1, QuickStyleFontColor: 1, QuickStyleVariation: 8 }, sections: themedText + relRect(), text: 'light fill\n' }) +
          // without the variation the colors are the theme's, readable or not
          shape({ id: 4, style: STYLE.THEMED, box: [6, 6, 2, 1], cells: { QuickStyleFillMatrix: 2, QuickStyleFontMatrix: 1 }, sections: themedText + relRect(), text: 'as themed\n' }) +
          // line style 1 gives the color itself, the fill's: with the line variation the outline is a darker fill color
          shape({ id: 5, style: STYLE.THEMED, box: [9, 2, 1, 1], cells: { QuickStyleFillMatrix: 2, QuickStyleLineMatrix: 1, QuickStyleVariation: 2 }, sections: themedText + relRect(), text: 'outlined\n' }) +
          shape({ id: 6, style: STYLE.THEMED, box: [9, 4, 1, 1], cells: { QuickStyleFillMatrix: 2, QuickStyleLineMatrix: 1 }, sections: themedText + relRect(), text: 'same line\n' }),
        (b) => (b.theme = themeXml(40)),
        { cells: { ThemeIndex: 40 } },
      ),
    )
    const s = slides[0]
    expect(shapes(s)[0].stroke).toMatchObject({ c: '#ed7d31', tail: 'triangle' })
    near(shapes(s)[0].stroke!.w, px(1 / 72))
    expect(shapes(s)[0].stroke!.dash).toBeTruthy()
    expect(under(s, 'dark fill').fill).toEqual({ t: 'solid', c: '#5b9bd5' })
    expect(byText(s, 'dark fill').text!.paras[0].runs[0].color).toBe('#ffffff')
    expect(under(s, 'light fill').fill).toEqual({ t: 'solid', c: '#feffff' })
    expect(byText(s, 'light fill').text!.paras[0].runs[0].color).toBe('#000000')
    expect(byText(s, 'as themed').text!.paras[0].runs[0].color).toBe('#5b9bd5')
    expect(under(s, 'same line')).toMatchObject({ fill: { c: '#5b9bd5' }, stroke: { c: '#5b9bd5' } })
    expect(under(s, 'outlined')).toMatchObject({ fill: { c: '#5b9bd5' }, stroke: { c: '#2d72b1' } })
  })

  test('a drawing whose theme part is missing: themed looks are black on white', async () => {
    const { slides } = await open(await onePage(shape({ id: 1, style: STYLE.THEMED, box: [2, 2, 2, 1], sections: relRect(), text: 'plain\n' })))
    expect(under(slides[0], 'plain')).toMatchObject({ fill: { t: 'solid', c: '#ffffff' }, stroke: { c: '#000000' } })
    expect(byText(slides[0], 'plain').text!.paras[0].runs[0]).toMatchObject({ color: '#000000' })
  })
})

// ---------------------------------------------------------------- pages

describe('visio: pages', () => {
  const square = (text: string) => shape({ id: 1, style: STYLE.NORMAL, box: [10, 10, 10, 10], cells: { LineWeight: 1 / 72 }, sections: rows('Character', [{ Size: 12 / 72 }]) + relRect(), text: `${text}\n` })

  test('a scaled drawing: sizes in drawing units, on paper at the page scale; line weights and fonts are paper sizes', async () => {
    // 1 in on paper is 10 ft: the page is 110 × 85 units, 11 × 8.5 in
    const { info, slides } = await open(await onePage(square('room'), undefined, { w: 110, h: 85, cells: { PageScale: 1, DrawingScale: 10 } }))
    expect(info.pptx).toMatchObject({ w: px(11), h: px(8.5) })
    box(under(slides[0], 'room'), 0.5, 8.5 - 1.5, 1, 1)
    near(under(slides[0], 'room').stroke!.w, px(1 / 72))
    expect(byText(slides[0], 'room').text!.paras[0].runs[0].sz).toBe(12)
  })

  test('pages of different sizes: one slide size that holds them all, each page a sheet at its center', async () => {
    const b = new VisioBuilder()
    b.page('Wide', square('wide'), { w: 20, h: 10 })
    b.page('Tall', square('tall'), { w: 10, h: 20 })
    const { info, slides } = await open(await b.zip())
    expect(info.pptx).toMatchObject({ w: px(20), h: px(20), count: 2 })
    for (const s of slides) expect(s.bg).toEqual({ t: 'solid', c: '#f0f0f0' })
    expect(slides[0].els[0]).toMatchObject({ fill: { t: 'solid', c: '#ffffff' }, stroke: { c: '#c8c8c8' } })
    box(slides[0].els[0], 0, 5, 20, 10)
    box(slides[1].els[0], 5, 0, 10, 20)
    // (5..15, 5..15) on each page
    box(under(slides[0], 'wide'), 5, 5 + 10 - 15, 10, 10)
    box(under(slides[1], 'tall'), 5 + 5, 20 - 15, 10, 10)
  })

  test('a page larger than a slide can be: everything smaller, line weights and fonts too', async () => {
    const { info, slides } = await open(await onePage(square('big'), undefined, { w: 112, h: 56 }))
    expect(info.pptx).toMatchObject({ w: px(56), h: px(28) })
    box(under(slides[0], 'big'), 2.5, 28 - 7.5, 5, 5)
    near(under(slides[0], 'big').stroke!.w, px(0.5 / 72))
    expect(byText(slides[0], 'big').text!.paras[0].runs[0].sz).toBe(6)
  })

  test('background pages: under the pages that name them, one under another; shown when there is nothing else', async () => {
    const b = new VisioBuilder()
    const mark = (t: string) => shape({ id: 1, style: STYLE.NORMAL, box: [2, 2, 1, 1], sections: relRect(), text: `${t}\n` })
    const deep = b.page('Deep', mark('deep'), { background: true })
    const back = b.page('Back', mark('back'), { background: true, backPage: deep })
    b.page('Front', mark('front'), { backPage: back })
    b.page('Alone', mark('alone'))
    // a page naming itself as its background is drawn once
    b.page('Loop', mark('loop'), { backPage: 16 })
    const { slides } = await open(await b.zip())
    expect(slides.map(texts)).toEqual([['deep', 'back', 'front'], ['alone'], ['loop']])

    const only = new VisioBuilder()
    only.page('Back', mark('back'), { background: true })
    expect((await open(await only.zip())).slides.map(texts)).toEqual([['back']])
  })

  test('layers: a shape is hidden when every layer it is on is', async () => {
    const on = (id: number, layers: string, t: string) => shape({ id, style: STYLE.NORMAL, box: [id, 2, 0.5, 0.5], cells: { LayerMember: layers }, sections: relRect(), text: `${t}\n` })
    const { slides } = await open(await onePage(on(1, '0', 'hidden') + on(2, '1', 'shown') + on(3, '0;1', 'on both') + on(4, '0;2', 'on two hidden') + on(5, '', 'on none'), undefined, { layers: [false, true, false] }))
    expect(texts(slides[0])).toEqual(['shown', 'on both', 'on none'])
  })

  test('a stencil (masters, no pages) shows its masters; a file with neither says so', async () => {
    const b = new VisioBuilder()
    b.master('Box', shape({ id: 5, style: STYLE.NORMAL, box: [1, 1, 1, 1], sections: relRect(), text: 'box\n' }), { w: 2, h: 2 })
    b.master('Dot', shape({ id: 5, style: STYLE.NORMAL, box: [1, 1, 1, 1], sections: relRect(), text: 'dot\n' }), { w: 2, h: 2 })
    const { info, slides } = await open(await b.zip('stencil'), 'shapes.vssx')
    expect(info).toMatchObject({ kind: 'pptx', from: 'visio', pptx: { w: px(2), h: px(2), count: 2 } })
    expect(slides.map(texts)).toEqual([['box'], ['dot']])
    await expect(open(await new VisioBuilder().zip())).rejects.toThrow('The Visio file has no pages')
  })

  test('a page whose part is missing is an empty slide; a shape naming a master that is not there draws what it has', async () => {
    const b = new VisioBuilder()
    b.page('Page-1', shape({ id: 1, master: 99, style: STYLE.NORMAL, box: [2, 2, 1, 1], sections: relRect(), text: 'orphan\n' }) + shape({ id: 2, masterShape: 5, style: STYLE.NORMAL, box: [5, 2, 1, 1], sections: relRect(), text: 'member\n' }))
    b.page('Page-2', '')
    const { slides } = await open(await b.zip())
    expect(slides.map(texts)).toEqual([['orphan', 'member'], []])
  })
})

// ---------------------------------------------------------------- files

describe('visio: files', () => {
  const head = new Uint8Array([0x50, 0x4b, 3, 4])

  test('the package is told by its document part, whatever its name; the extensions are never read as text', async () => {
    for (const name of ['a.vsdx', 'a.vsdm', 'a.vstx', 'a.vssx', 'A.VSDX', 'noname']) expect((await open(await sampleVisio(), name)).info.from).toBe('visio')
    for (const ext of ['vsdx', 'vsdm', 'vstx', 'vstm', 'vssx', 'vssm']) {
      expect(mayBePlain(`a.${ext}`)).toBe(false)
      expect(plainFormat(`a.${ext}`, undefined, head, true)).toBeNull()
    }
  })

  test('a Visio 2003-2010 drawing (.vsd) says what to do', async () => {
    await expect(convertCfb(writeCfb({ VisioDocument: new Uint8Array(100) }))).rejects.toThrow('Visio 2003-2010 drawings (.vsd) are not supported: save the drawing as .vsdx')
    const engine = new Engine(() => {})
    await expect(engine.open(new Blob([writeCfb({ VisioDocument: new Uint8Array(5000) })]), { name: 'old.vsd' })).rejects.toThrow(/save the drawing as \.vsdx/)
  })
})

describe('visio: limits', () => {
  const saved = { ...VISIO_LIMITS }
  afterEach(() => Object.assign(VISIO_LIMITS, saved))

  /** 40 instances of a master of 50 shapes: 2,040 shapes drawn from a few kilobytes */
  const many = () =>
    onePage(Array.from({ length: 40 }, (_, i) => shape({ id: i + 1, type: 'Group', master: 1, cells: { PinX: 1 + (i % 10), PinY: 1 + Math.floor(i / 10) } })).join(''), (b) => {
      b.master('Many', shape({ id: 5, type: 'Group', style: STYLE.NORMAL, box: [1, 1, 1, 1], cells: { FillPattern: 0, LinePattern: 0 }, shapes: Array.from({ length: 50 }, (_, k) => shape({ id: 10 + k, style: STYLE.NORMAL, box: [0.5, 0.5, 0.2, 0.2], sections: relRect() })) }))
    })

  test('a drawing that draws its masters over and over opens while it is within the limits', async () => {
    const { slides } = await open(await many())
    expect(shapes(slides[0])).toHaveLength(2000)
  })

  test('more shapes, or more slide than the limits: refused, not built', async () => {
    VISIO_LIMITS.shapes = 1000
    await expect(open(await many())).rejects.toThrow('The Visio drawing is too large to show')
    Object.assign(VISIO_LIMITS, saved)
    VISIO_LIMITS.output = 100_000
    await expect(open(await many())).rejects.toThrow('The Visio drawing is too large to show')
  })

  test('shapes nested without end, and numbers that are not, do not stop the rest from showing', async () => {
    let deep = shape({ id: 9999, style: STYLE.NORMAL, box: [0.5, 0.5, 1, 1], sections: relRect(), text: 'bottom\n' })
    for (let i = 0; i < 200; i++) deep = shape({ id: 1000 + i, type: 'Group', style: STYLE.NORMAL, box: [1, 1, 2, 2], cells: { FillPattern: 0, LinePattern: 0 }, shapes: [deep] })
    const { slides } = await open(
      await onePage(
        deep +
          shape({ id: 1, style: STYLE.NORMAL, box: [3, 3, 1, 1], cells: { Angle: 'NaN', PinX: '1e999' }, sections: relRect(), text: 'nowhere\n' }) +
          shape({ id: 2, style: STYLE.NORMAL, box: [5, 3, 1, 1], sections: geom(0, [['MoveTo', { X: 0, Y: 0 }], ['LineTo', { X: '1e400', Y: 1 }], ['ArcTo', { X: 1, Y: 1, A: '1e308' }], ['NURBSTo', { X: 1, Y: 0, E: `NURBS(1, 99999, 1, 1${', 0.5,0.5,0,1'.repeat(20000)})` }], ['LineTo', { X: 0, Y: 0 }]]), text: 'still here\n' }),
      ),
    )
    // groups nest as deep as the XML does (P3-039); a cell that is not a number counts as missing (the pin at 0, no turn)
    expect(texts(slides[0])).toEqual(['bottom', 'nowhere', 'still here'])
    box(under(slides[0], 'nowhere'), -0.5, 8.5 - 3.5, 1, 1)
    expect(under(slides[0], 'still here').path).toMatch(/^M[\d. ]+L/)
    expect(under(slides[0], 'still here').path).not.toMatch(/NaN|Infinity/)
  })
})
