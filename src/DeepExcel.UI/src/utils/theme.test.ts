import { describe, expect, it } from 'vitest'
import { nextThemePref, parseHostTheme, parseThemePref, resolveTheme } from './theme'

describe('panel theme', () => {
  it('follows the host unless the user pinned one', () => {
    expect(resolveTheme('auto', 'dark', false)).toBe('dark')
    expect(resolveTheme('auto', 'light', true)).toBe('light')
    expect(resolveTheme('light', 'dark', true)).toBe('light')
    expect(resolveTheme('dark', 'light', false)).toBe('dark')
  })

  it('falls back to the system color scheme when the host says system', () => {
    expect(resolveTheme('auto', 'system', true)).toBe('dark')
    expect(resolveTheme('auto', 'system', false)).toBe('light')
  })

  it('cycles auto → light → dark → auto', () => {
    expect(nextThemePref('auto')).toBe('light')
    expect(nextThemePref('light')).toBe('dark')
    expect(nextThemePref('dark')).toBe('auto')
  })

  it('treats anything unexpected as auto / system', () => {
    expect(parseThemePref(null)).toBe('auto')
    expect(parseThemePref('dark')).toBe('dark')
    expect(parseHostTheme(undefined)).toBe('system')
    expect(parseHostTheme('Black')).toBe('system')
  })
})
