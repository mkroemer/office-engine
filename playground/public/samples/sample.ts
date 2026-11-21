/**
 * A source file opens as code: colored by its language (from its name, its MIME type or the `language` prop),
 * with line numbers, edited as text and saved as it is.
 */
import { readFile } from 'node:fs/promises'

interface Reading {
  site: string
  value: number
  done?: boolean
}

const SITES = ['North', 'South'] as const

export async function load(path: string): Promise<Reading[]> {
  const text = await readFile(path, 'utf8')
  return text
    .split(/\r?\n/)
    .filter((line) => line.trim() !== '')
    .map((line) => {
      const [site, value] = line.split(',')
      return { site, value: Number(value), done: SITES.includes(site as (typeof SITES)[number]) }
    })
}

// totals by site, as a template literal
export const summary = (rows: Reading[]) => `${rows.length} readings, ${rows.reduce((n, r) => n + r.value, 0)} in all`
