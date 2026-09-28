import { describe, expect, it } from 'vitest'
import { escOwnedByOverlay, isEditable } from './keys'

// 不引 jsdom：只要 tagName / type / querySelector 这几个字段
const el = (tagName: string, extra: Record<string, unknown> = {}) => ({ tagName, ...extra }) as unknown as EventTarget

describe('keyboard guards', () => {
  it('treats text fields as the owner of Enter / Esc', () => {
    expect(isEditable(el('TEXTAREA'))).toBe(true)
    expect(isEditable(el('INPUT', { type: 'text' }))).toBe(true)
    expect(isEditable(el('INPUT', { type: 'checkbox' }))).toBe(false)
    expect(isEditable(el('BUTTON'))).toBe(false)
    expect(isEditable(el('DIV', { isContentEditable: true }))).toBe(true)
    expect(isEditable(null)).toBe(false)
  })

  it('lets an open menu or dialog take Esc', () => {
    const seen: string[] = []
    const root = (hit: boolean) => ({ querySelector: (q: string) => { seen.push(q); return hit ? {} : null } }) as unknown as ParentNode
    expect(escOwnedByOverlay(root(false))).toBe(false)
    expect(escOwnedByOverlay(root(true))).toBe(true)
    expect(seen[0]).toContain('.header-menu-list')
    expect(seen[0]).toContain('.model-picker-pop')
  })
})
