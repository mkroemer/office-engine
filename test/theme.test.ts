import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { THEME_TOKENS, themeMode, themeStyle } from '../src/react/theme'

const css = readFileSync(`${import.meta.dir}/../src/react/styles.css`, 'utf8')
const block = (selector: string) => {
  const at = css.indexOf(`${selector} {\n`)
  return css.slice(at, css.indexOf('\n}', at))
}
const defined = (text: string) => [...text.matchAll(/^\s*--de-([a-z0-9-]+):/gm)].map((m) => m[1])
const COLOR = /#[0-9a-f]{3,8}\b|rgba?\(|hsla?\(/i

describe('theme tokens', () => {
  test('THEME_TOKENS lists exactly the tokens the stylesheet defines on .de-root', () => {
    expect([...defined(block('.de-root'))].sort()).toEqual([...THEME_TOKENS].sort())
  })

  test('the dark theme redefines only known tokens', () => {
    const dark = defined(block('.de-root.de-dark'))
    expect(dark.length).toBeGreaterThan(40)
    for (const t of dark) expect(THEME_TOKENS as readonly string[]).toContain(t)
  })

  test('every token the stylesheet uses is defined', () => {
    const used = new Set([...css.matchAll(/var\(--de-([a-z0-9-]+)/g)].map((m) => m[1]))
    const known = new Set<string>(THEME_TOKENS)
    expect([...used].filter((t) => !known.has(t))).toEqual([])
  })

  test('no literal color outside the token definitions and var() fallbacks', () => {
    const stray = css
      .split('\n')
      .map((line, i) => ({ line: line.replace(/var\(--de-[a-z0-9-]+,[^)]*\)?\)/g, ''), n: i + 1 }))
      .filter(({ line }) => COLOR.test(line) && !/^\s*--de-[a-z0-9-]+:/.test(line))
    expect(stray).toEqual([])
  })
})

describe('themeStyle', () => {
  const theme = {
    mode: 'auto' as const,
    tokens: { accent: '#7c3aed', font: 'Inter, sans-serif', radiusLg: '4px' },
    light: { bg: '#fafafa' },
    dark: { bg: '#111', accent: '#a78bfa' },
  }

  test('maps camelCase names to the custom properties', () => {
    expect(themeStyle(theme, false)).toEqual({ '--de-accent': '#7c3aed', '--de-font': 'Inter, sans-serif', '--de-radius-lg': '4px', '--de-bg': '#fafafa' } as never)
  })

  test('the mode set goes over the shared tokens', () => {
    expect(themeStyle(theme, true)).toEqual({ '--de-accent': '#a78bfa', '--de-font': 'Inter, sans-serif', '--de-radius-lg': '4px', '--de-bg': '#111' } as never)
  })

  test('a plain mode sets no properties', () => {
    expect(themeStyle('dark', true)).toEqual({})
    expect(themeMode('dark')).toBe('dark')
    expect(themeMode(theme)).toBe('auto')
    expect(themeMode({ tokens: {} })).toBe('light')
    expect(themeMode(undefined)).toBe('light')
  })
})
