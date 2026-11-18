import { describe, expect, test } from 'bun:test'
import { en, formatBytes, formatCount, formatNumber, resolveStrings } from '../src/react/strings'
import { de } from '../src/locales/de'

type Entry = [path: string, value: unknown]

/** every leaf of a table, with its dotted key */
function leaves(table: object, prefix = ''): Entry[] {
  return Object.entries(table).flatMap(([k, v]): Entry[] => (v && typeof v === 'object' ? leaves(v, `${prefix}${k}.`) : [[prefix + k, v]]))
}

/** a sample call of a text function: numbers for counts, names otherwise */
const sample = (fn: (...args: never[]) => unknown, n: number) => (fn as (...a: unknown[]) => unknown)(...Array.from({ length: fn.length }, () => n))

describe('UI strings', () => {
  const enLeaves = leaves(en)

  test('every locale has exactly the keys of en, with the same kind of value', () => {
    for (const [name, table] of Object.entries({ en, de })) {
      const got = new Map(leaves(table))
      expect([...got.keys()].sort(), name).toEqual(enLeaves.map(([k]) => k).sort())
      for (const [k, v] of enLeaves) if (k !== 'locale') expect(typeof got.get(k), `${name}.${k}`).toBe(typeof v)
    }
  })

  test('no text is empty, also for 1 and for many', () => {
    for (const [name, table] of Object.entries({ en, de })) {
      for (const [k, v] of leaves(table)) {
        if (k === 'locale') continue
        const texts = typeof v === 'function' ? [sample(v as never, 1), sample(v as never, 3)] : [v]
        for (const t of texts) {
          expect(typeof t, `${name}.${k}`).toBe('string')
          expect((t as string).trim().length, `${name}.${k}`).toBeGreaterThan(0)
        }
      }
    }
  })

  test('English is the only default; German is supplied explicitly with standard spelling', () => {
    expect(en.locale).toBeUndefined()
    expect(resolveStrings('de-DE').common.cancel).toBe(en.common.cancel)
    expect(resolveStrings(undefined, de)).toEqual(de)
    expect(de.common.close).toBe('Schließen')
    expect(de.details.fileSize).toBe('Dateigröße')
    expect(de.headerFooter.kinds.footer).toBe('Fußzeile')
  })

  test('an override merges key by key over the locale', () => {
    const s = resolveStrings(undefined, { ...de, file: { ...de.file, menu: 'Ablage' }, zoom: { ...de.zoom, in: 'Größer' } })
    expect(s.file.menu).toBe('Ablage')
    expect(s.file.print).toBe(de.file.print)
    expect(s.zoom.in).toBe('Größer')
    expect(s.zoom.out).toBe(de.zoom.out)
    expect(s.details.core.title).toBe(de.details.core.title)
    // the shipped tables are left alone
    expect(de.file.menu).toBe('Datei')
    const british = resolveStrings(undefined, { locale: 'en-GB', edit: { textColor: 'Text colour' } })
    expect(british.locale).toBe('en-GB')
    expect(british.edit.textColor).toBe('Text colour')
    expect(british.edit.bold).toBe(en.edit.bold)
    expect(en.edit.textColor).toBe('Text color')
    expect(de.locale).toBe('de-DE')
    expect(de.common.close).toBe('Schließen')
    for (const locale of ['fr', 'de-AT', 'constructor', 'toString', '__proto__']) {
      expect(resolveStrings(locale).common.cancel, locale).toBe(en.common.cancel)
      expect(resolveStrings(locale, { file: { menu: 'Custom' } }).common.cancel, locale).toBe('Cancel')
    }
  })

  test('numbers follow explicit English and German formatting', () => {
    const english = resolveStrings('en-US')
    expect(formatBytes(english, 512)).toBe('512 B')
    expect(formatBytes(english, 2048)).toBe('2 KB')
    expect(formatBytes(english, 1536 * 1024 * 1024)).toBe('1536.0 MB')
    expect(formatCount(en, 1234567)).toBe((1234567).toLocaleString())
    expect(de.locale).toBe('de-DE')
    expect(formatCount(de, 1234567)).toBe('1.234.567')
    expect(formatNumber(de, 1234.5, { useGrouping: true })).toBe('1.234,5')
    expect(formatBytes(de, 1536 * 1024 * 1024)).toBe('1536,0 MB')
  })
})

test('omitting locale preserves browser formatting; malformed tags fall back without throwing', () => {
  expect(en.locale).toBeUndefined()
  expect(formatCount(resolveStrings(), 1234567)).toBe((1234567).toLocaleString())
  expect(resolveStrings('en-GB').locale).toBe('en-GB')
  expect(resolveStrings('en-GB').common.cancel).toBe('Cancel')
  for (const locale of ['de_DE', '', 'constructor', '__proto__']) {
    const s = resolveStrings(locale)
    expect(s.locale).toBeUndefined()
    expect(() => formatCount(s, 1234)).not.toThrow()
    expect(() => new Date().toLocaleString(s.locale)).not.toThrow()
  }
  expect(resolveStrings(undefined, { locale: 'de_DE' }).locale).toBeUndefined()
})

test('count callbacks accept numbers and choose singular/plural before formatting', () => {
  expect(de.docx.status(1, 1, 1)).toBe('Seite 1 von 1 · 1 Wort')
  expect(de.docx.status(1, 2, 1234, de.locale)).toBe('Seite 1 von 2 · 1.234 Wörter')
  expect(de.xlsx.status(1, 1)).toBe('1 Zeile × 1 Spalte')
  expect(de.xlsx.usedRangeValue(2, 3)).toBe('2 Zeilen × 3 Spalten')
  expect(en.xlsx.status(1, 1)).toBe('1 row × 1 col')
})
