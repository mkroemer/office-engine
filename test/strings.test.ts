import { describe, expect, test } from 'bun:test'
import { de, en, formatBytes, formatCount, LOCALES, resolveStrings, type Strings } from '../src/react/strings'

type Entry = [path: string, value: unknown]

/** every leaf of a table, with its dotted key */
function leaves(table: object, prefix = ''): Entry[] {
  return Object.entries(table).flatMap(([k, v]): Entry[] => (v && typeof v === 'object' ? leaves(v, `${prefix}${k}.`) : [[prefix + k, v]]))
}

/** a sample call of a text function: numbers for counts, names otherwise */
const sample = (fn: (...args: never[]) => unknown, n: number) => (fn as (...a: unknown[]) => unknown)(...Array.from({ length: fn.length }, () => n), 'x')

describe('UI strings', () => {
  const enLeaves = leaves(en)

  test('every locale has exactly the keys of en, with the same kind of value', () => {
    for (const [name, table] of Object.entries(LOCALES)) {
      const got = new Map(leaves(table))
      expect([...got.keys()].sort(), name).toEqual(enLeaves.map(([k]) => k).sort())
      for (const [k, v] of enLeaves) expect(typeof got.get(k), `${name}.${k}`).toBe(typeof v)
    }
  })

  test('no text is empty, also for 1 and for many', () => {
    for (const [name, table] of Object.entries(LOCALES)) {
      for (const [k, v] of leaves(table)) {
        const texts = typeof v === 'function' ? [sample(v as never, 1), sample(v as never, 3)] : [v]
        for (const t of texts) {
          expect(typeof t, `${name}.${k}`).toBe('string')
          expect((t as string).trim().length, `${name}.${k}`).toBeGreaterThan(0)
        }
      }
    }
  })

  test('German is Swiss: no sharp s', () => {
    for (const [k, v] of leaves(de)) {
      const texts = typeof v === 'function' ? [sample(v as never, 1), sample(v as never, 3)] : [v]
      for (const t of texts) expect(String(t).includes('ß'), `de.${k}`).toBe(false)
    }
  })

  test('an override merges key by key over the locale', () => {
    const s = resolveStrings('de', { file: { menu: 'Ablage' }, zoom: { in: 'Grösser' } })
    expect(s.file.menu).toBe('Ablage')
    expect(s.file.print).toBe(de.file.print)
    expect(s.zoom.in).toBe('Grösser')
    expect(s.zoom.out).toBe(de.zoom.out)
    expect(s.details.core.title).toBe(de.details.core.title)
    // the shipped tables are left alone
    expect(de.file.menu).toBe('Datei')
    expect(resolveStrings('fr' as never)).toBe(en as Strings)
  })

  test('numbers keep the English output and follow the locale', () => {
    expect(formatBytes(en, 512)).toBe('512 B')
    expect(formatBytes(en, 2048)).toBe('2 KB')
    expect(formatBytes(en, 1536 * 1024 * 1024)).toBe('1536.0 MB')
    expect(formatCount(en, 1234567)).toBe('1,234,567')
    expect(formatCount(de, 1234567)).not.toBe('1,234,567')
  })
})
