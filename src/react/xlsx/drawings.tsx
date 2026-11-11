import { useEffect, useState, type CSSProperties, type ReactNode } from 'react'
import type { DocumentClient } from '../../core/worker/client'
import type { SheetDrawing } from '../../core/xlsx/drawings'
import { Icon } from '../edit/Toolbar'
import type { Axis } from './axis'

/** a sheet's drawing with its picture's object URL, where the browser can draw it */
export interface ShownDrawing extends SheetDrawing {
  url?: string
}

/** The pictures and charts of the sheet (`xlsx.drawings`), read again when another sheet is shown */
export function useSheetDrawings(client: DocumentClient, sheet: number): ShownDrawing[] {
  const [list, setList] = useState<ShownDrawing[]>([])
  useEffect(() => {
    let live = true
    setList([])
    client
      .call('xlsx.drawings', sheet)
      .then((drawings) => Promise.all(drawings.map(async (d) => ({ ...d, url: d.target ? await client.mediaUrl(d.target).catch(() => undefined) : undefined }))))
      .then((shown) => live && setList(shown))
      .catch((e: Error) => console.warn('[office-engine] drawings not read:', e.message))
    return () => {
      live = false
    }
  }, [client, sheet])
  return list
}

/**
 * Where a drawing lies on the sheet, px from its top left: anchored to cells it moves and grows with their rows and
 * columns (a hidden one is 0 wide); given by a size it keeps that size.
 */
export function drawingBox(d: SheetDrawing, cols: Axis, rows: Axis): { x: number; y: number; w: number; h: number } | null {
  let x: number
  let y: number
  if (d.from) {
    x = cols.pos(d.from.c) + d.from.dx
    y = rows.pos(d.from.r) + d.from.dy
  } else if (d.at) {
    x = d.at.x
    y = d.at.y
  } else return null
  const w = d.to ? cols.pos(d.to.c) + d.to.dx - x : (d.size?.w ?? 0)
  const h = d.to ? rows.pos(d.to.r) + d.to.dy - y : (d.size?.h ?? 0)
  return w >= 2 && h >= 2 ? { x, y, w, h } : null
}

/** How many rows and columns the sheet needs to reach every drawing (a drawing given by a size taken over default-sized cells) */
export function drawingExtent(list: SheetDrawing[], cols: Axis, rows: Axis): { rows: number; cols: number } {
  let r = 0
  let c = 0
  for (const d of list) {
    const end = d.to ?? (d.from && d.size ? { c: d.from.c + Math.ceil(d.size.w / cols.def), r: d.from.r + Math.ceil(d.size.h / rows.def) } : undefined)
    const fromAt = !end && d.at && d.size ? { c: Math.ceil((d.at.x + d.size.w) / cols.def), r: Math.ceil((d.at.y + d.size.h) / rows.def) } : undefined
    const e = end ?? fromAt
    if (e) {
      r = Math.max(r, e.r + 1)
      c = Math.max(c, e.c + 1)
    }
  }
  return { rows: r, cols: c }
}

const CHART_NAMES: Record<string, string> = {
  column: 'Column chart', bar: 'Bar chart', line: 'Line chart', pie: 'Pie chart', doughnut: 'Doughnut chart', area: 'Area chart',
  scatter: 'Scatter chart', bubble: 'Bubble chart', radar: 'Radar chart', stock: 'Stock chart', surface: 'Surface chart',
}

/** A drawing as shown: the picture itself, or a box of its size saying what is there (a chart, a picture the browser cannot draw) */
export function drawingEl(key: string, d: ShownDrawing, style: CSSProperties): ReactNode {
  if (d.url) return <img key={key} className="de-xl-pic" src={d.url} alt={d.title} style={style} draggable={false} />
  const what = d.kind === 'chart' ? CHART_NAMES[d.chart ?? ''] ?? 'Chart' : 'Picture'
  return (
    <div key={key} className="de-xl-drawing" style={style} title={[what, d.title].filter(Boolean).join(': ')}>
      <Icon name={d.kind === 'chart' ? 'chart' : 'image'} size={24} />
      <strong>{what}</strong>
      {d.title && <span>{d.title}</span>}
    </div>
  )
}
