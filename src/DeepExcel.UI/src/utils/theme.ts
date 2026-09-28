// 面板深浅色：默认跟随宿主（Excel 回 Office 主题，WPS 回 system），用户可以在「更多」菜单里固定。
// 宿主回 system 时按 prefers-color-scheme（WebView2 跟随 Windows 的应用主题）。
// 纯函数，有测试：theme.test.ts

export type ThemePref = 'auto' | 'light' | 'dark'
export type HostTheme = 'light' | 'dark' | 'system'
export type Theme = 'light' | 'dark'

export const THEME_PREF_KEY = 'deepexcel.theme'

export const THEME_PREF_TEXT: Record<ThemePref, string> = {
  auto: '外观：跟随 Office',
  light: '外观：浅色',
  dark: '外观：深色',
}

export function resolveTheme(pref: ThemePref, host: HostTheme, systemDark: boolean): Theme {
  if (pref !== 'auto') return pref
  if (host !== 'system') return host
  return systemDark ? 'dark' : 'light'
}

export function nextThemePref(pref: ThemePref): ThemePref {
  return pref === 'auto' ? 'light' : pref === 'light' ? 'dark' : 'auto'
}

export function parseThemePref(raw: unknown): ThemePref {
  return raw === 'light' || raw === 'dark' ? raw : 'auto'
}

export function parseHostTheme(raw: unknown): HostTheme {
  return raw === 'light' || raw === 'dark' ? raw : 'system'
}
