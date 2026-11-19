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

describe('color-mix compatibility', () => {
  test('mixed source selection tokens are gated by feature support', () => {
    const gated = css.match(/@supports \(background: color-mix\(in srgb, red, transparent\)\) \{\n([\s\S]*?)\n\}/)
    expect(gated).not.toBeNull()
    expect(gated![1]).toContain('.de-root { --de-src-sel: color-mix(in srgb, var(--de-accent) 25%, transparent); }')
    expect(gated![1]).toContain('.de-root.de-dark { --de-src-sel: color-mix(in srgb, var(--de-accent) 35%, transparent); }')
    const withoutSupport = css.replace(gated![0], '')
    // Duplicate custom properties alone would still accept the unsupported function as raw tokens.
    expect(withoutSupport).not.toMatch(/--de-[\w-]+:\s*color-mix\(/)
    expect(block('.de-root')).toContain('--de-src-sel: var(--de-accent-soft);')
    expect(css).toContain('.de-src-input::selection { background: var(--de-src-sel); }')
  })

  test('each mixed property has an immediately preceding non-mixed fallback', () => {
    let checked = 0
    for (const body of css.split(/[{}]/)) {
      const declarations = body.split(';').map((s) => s.trim()).filter(Boolean)
      for (const [i, declaration] of declarations.entries()) {
        const mixed = declaration.match(/^([a-z-]+):\s*.*color-mix\(/)
        if (!mixed || mixed[1].startsWith('--')) continue
        const fallback = declarations[i - 1] ?? ''
        expect(fallback.startsWith(`${mixed[1]}:`)).toBe(true)
        expect(fallback).not.toContain('color-mix(')
        expect(fallback).toContain('var(--de-')
        checked++
      }
    }
    expect(checked).toBeGreaterThan(0)
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
