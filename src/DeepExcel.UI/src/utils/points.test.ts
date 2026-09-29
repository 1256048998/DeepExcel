import { describe, expect, it } from 'vitest'
import { pointsForModel, pointsLabel } from './points'

const weights = { deepseek: 1, 'claude-sonnet': 12, 'claude-sonnet-5': 8, 'claude-opus': 58 }

describe('points per task', () => {
  it('matches the longest model-name prefix, case-insensitively', () => {
    expect(pointsForModel('deepseek-v4-pro', weights)).toBe(1)
    expect(pointsForModel('DeepSeek-V4-Flash', weights)).toBe(1)
    expect(pointsForModel('claude-sonnet-4-5', weights)).toBe(12)
    expect(pointsForModel('claude-sonnet-5', weights)).toBe(8)
    expect(pointsForModel('claude-opus-5', weights)).toBe(58)
  })

  it('uses the server fallback for an unpriced model, and shows nothing rather than guessing', () => {
    expect(pointsForModel('gpt-9', weights, 58)).toBe(58)
    expect(pointsForModel('gpt-9', weights)).toBeNull()
    expect(pointsForModel('deepseek-v4-pro', null)).toBeNull()
    expect(pointsForModel('', weights)).toBeNull()
  })

  it('labels per task', () => {
    expect(pointsLabel(12)).toBe('12 点/任务')
  })
})
