import { describe, it, expect } from 'vitest'
import { pickSelectedProvider, shouldResetForm } from './providerSync'

/**
 * 用户报的：选了别的厂商 → 获取模型 → 添加，界面跳回 Anthropic。
 * 原因是 config 每变一次都把选中厂商重置成 currentProvider，而「添加模型」会回写 config。
 */

const config = {
  currentProvider: 'anthropic',
  providers: { anthropic: {}, deepseek: {}, zhipu: {} },
}

describe('pickSelectedProvider', () => {
  it('keeps the vendor the user is looking at when config refreshes after adding a model', () => {
    expect(pickSelectedProvider('deepseek', config, false)).toBe('deepseek')
  })

  it('follows the saved current vendor when the dialog has just been opened', () => {
    expect(pickSelectedProvider('deepseek', config, true)).toBe('anthropic')
  })

  it('follows the saved current vendor on first load, when nothing is selected yet', () => {
    expect(pickSelectedProvider('', config, false)).toBe('anthropic')
  })

  it('falls back when the selected vendor no longer exists in the config', () => {
    expect(pickSelectedProvider('gone', config, false)).toBe('anthropic')
  })

  it('does not treat inherited object keys as vendors', () => {
    expect(pickSelectedProvider('toString', config, false)).toBe('anthropic')
  })
})

describe('shouldResetForm', () => {
  it('does not wipe the typed key and base URL on a same-vendor refresh', () => {
    expect(shouldResetForm('deepseek', 'deepseek')).toBe(false)
  })

  it('resets when the user switches vendor', () => {
    expect(shouldResetForm('deepseek', 'zhipu')).toBe(true)
  })

  it('resets after the dialog is reopened (caller clears the form owner)', () => {
    expect(shouldResetForm('', 'deepseek')).toBe(true)
  })
})
