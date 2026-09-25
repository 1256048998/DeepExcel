import { describe, expect, it } from 'vitest'
import { STICK_THRESHOLD_PX, isNearBottom } from './useStickToBottom'

describe('isNearBottom', () => {
  const at = (gap: number) => ({ scrollHeight: 2000, clientHeight: 600, scrollTop: 2000 - 600 - gap })

  it('counts anything within the threshold as the bottom', () => {
    expect(isNearBottom(at(0))).toBe(true)
    expect(isNearBottom(at(STICK_THRESHOLD_PX))).toBe(true)
  })

  it('stops following once the user scrolls further up', () => {
    expect(isNearBottom(at(STICK_THRESHOLD_PX + 1))).toBe(false)
    expect(isNearBottom(at(1200))).toBe(false)
  })

  it('treats content shorter than the viewport as at the bottom', () => {
    expect(isNearBottom({ scrollHeight: 300, clientHeight: 600, scrollTop: 0 })).toBe(true)
  })
})
