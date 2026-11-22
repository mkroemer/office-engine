import type { ChartModel } from '../../core/xlsx/chart'
import { renderChart } from '../../core/xlsx/chart-svg'

/** each chart's picture, drawn for the size it was last shown at ('' when it could not be drawn at that size) */
const chartUrls = new WeakMap<ChartModel, { key: string; url: string }>()

/**
 * The chart drawn at `w` × `h` px as an SVG picture's object URL (in an img: nothing in it runs), or '' when drawing it
 * failed: charts are drawn in a view's render, so one a crafted file breaks shows its placeholder box, not a blank view.
 */
export function chartUrl(m: ChartModel, w: number, h: number): string {
  const key = `${Math.round(w)}x${Math.round(h)}`
  const had = chartUrls.get(m)
  if (had?.key === key) return had.url
  if (had?.url) URL.revokeObjectURL(had.url)
  let url = ''
  try {
    url = URL.createObjectURL(new Blob([renderChart(m, Math.max(1, Math.round(w)), Math.max(1, Math.round(h)))], { type: 'image/svg+xml' }))
  } catch (e) {
    console.warn('[office-engine] chart not drawn:', (e as Error).message)
  }
  chartUrls.set(m, { key, url })
  return url
}

/** lets go of the pictures of charts no longer shown */
export function releaseChart(m: ChartModel | undefined) {
  const had = m && chartUrls.get(m)
  if (!had) return
  if (had.url) URL.revokeObjectURL(had.url)
  chartUrls.delete(m)
}
