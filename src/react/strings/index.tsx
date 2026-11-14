import { createContext, useContext, useMemo, type ReactNode } from 'react'
import { en, type Strings } from './en'
import { de } from './de'

export type { Strings } from './en'
export { en } from './en'
export { de } from './de'

/** Locales shipped with the engine. */
export type Locale = 'en' | 'de'

export const LOCALES: Record<Locale, Strings> = { en, de }

type DeepPartial<T> = { [K in keyof T]?: T[K] extends string | ((...args: never[]) => unknown) ? T[K] : DeepPartial<T[K]> }

/** A partial table: groups merge key by key over the locale's texts. */
export type StringsOverride = DeepPartial<Strings>

const isGroup = (v: unknown): v is Record<string, unknown> => !!v && typeof v === 'object'

function merge(base: Record<string, unknown>, over: Record<string, unknown>): Record<string, unknown> {
  const out = { ...base }
  for (const [k, v] of Object.entries(over)) {
    if (v === undefined) continue
    const cur = out[k]
    out[k] = isGroup(cur) && isGroup(v) ? merge(cur, v) : v
  }
  return out
}

/** The texts for `locale` with `override` merged over them (unknown locales fall back to English). */
export function resolveStrings(locale: Locale | string = 'en', override?: StringsOverride): Strings {
  const base = LOCALES[locale as Locale] ?? en
  return override ? (merge(base, override as Record<string, unknown>) as Strings) : base
}

const formats = new Map<string, Intl.NumberFormat>()

/** A number in the texts' locale, without digit grouping unless `opts` asks for it. */
export function formatNumber(s: Strings, n: number, opts?: Intl.NumberFormatOptions): string {
  const key = s.locale + JSON.stringify(opts ?? {})
  let f = formats.get(key)
  if (!f) formats.set(key, (f = new Intl.NumberFormat(s.locale, { useGrouping: false, ...opts })))
  return f.format(n)
}

/** A count (pages, words, rows) with digit grouping. */
export const formatCount = (s: Strings, n: number) => formatNumber(s, n, { useGrouping: true })

/** A byte count as B / KB / MB. */
export function formatBytes(s: Strings, n: number): string {
  if (n < 1024) return `${formatNumber(s, n)} B`
  if (n < 1024 * 1024) return `${formatNumber(s, n / 1024, { maximumFractionDigits: 0 })} KB`
  return `${formatNumber(s, n / 1024 / 1024, { minimumFractionDigits: 1, maximumFractionDigits: 1 })} MB`
}

const StringsContext = createContext<Strings>(en)

/** Provides the UI texts to the views below it (`DocViewer` provides its own from `locale` / `strings`). */
export function StringsProvider({ locale, strings, children }: { locale?: Locale | string; strings?: StringsOverride; children?: ReactNode }) {
  const value = useMemo(() => resolveStrings(locale, strings), [locale, strings])
  return <StringsContext.Provider value={value}>{children}</StringsContext.Provider>
}

/** The UI texts of the nearest `StringsProvider` (English without one). */
export const useStrings = () => useContext(StringsContext)
