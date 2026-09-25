import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { PERMISSION_MODE_ORDER, PERMISSION_MODE_TEXT, nextPermissionMode } from './InputArea'

describe('permission modes', () => {
  it('cycles default → accept_writes → plan → default', () => {
    expect(nextPermissionMode('default')).toBe('accept_writes')
    expect(nextPermissionMode('accept_writes')).toBe('plan')
    expect(nextPermissionMode('plan')).toBe('default')
  })

  it('matches the modes the sidecar understands', () => {
    const source = readFileSync(resolve(__dirname, '../../../DeepExcel.Sidecar/permission_modes.py'), 'utf8')
    const modes = source.match(/^MODES = \(([^)]*)\)/m)![1]
    for (const m of PERMISSION_MODE_ORDER) {
      expect(source).toContain(`= "${m}"`)
      expect(PERMISSION_MODE_TEXT[m].label.length).toBeGreaterThan(0)
    }
    expect(modes.split(',').filter(s => s.trim()).length).toBe(PERMISSION_MODE_ORDER.length)
  })

  it('is never persisted by the panel', () => {
    // 「本次会话自动应用写入」不能记住：App 里的模式只放在 state
    const app = readFileSync(resolve(__dirname, '../App.tsx'), 'utf8')
    expect(app).not.toMatch(/localStorage[^\n]*[Pp]ermission/)
    expect(app).not.toMatch(/[Pp]ermission[Mm]ode[^\n]*localStorage/)
  })
})
