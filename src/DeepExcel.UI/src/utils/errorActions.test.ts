import { describe, expect, it } from 'vitest'
import { errorActions, retryTarget } from './errorActions'

describe('errorActions', () => {
  it('offers retry only when the error is retryable and there is something to resend', () => {
    expect(errorActions({ code: 'network', retryable: true }, true)).toEqual(['retry'])
    expect(errorActions({ code: 'network', retryable: true }, false)).toEqual([])
    expect(errorActions({ code: 'auth', retryable: false }, true)).toEqual(['settings'])
  })

  it('points each class of error at the place that fixes it', () => {
    expect(errorActions({ code: 'model_not_found' }, true)).toEqual(['settings'])
    expect(errorActions({ code: 'quota' }, true)).toEqual(['account'])
    expect(errorActions({ code: 'task_limit' }, true)).toEqual(['account'])
    expect(errorActions({ code: 'context_too_long' }, true)).toEqual(['new_chat'])
    expect(errorActions({ code: 'unknown', retryable: true, detail: 'Traceback…' }, true)).toEqual(['retry', 'copy_detail'])
  })
})

describe('retryTarget', () => {
  const msgs = [
    { role: 'user', content: '第一问' },
    { role: 'assistant', content: '答' },
    { role: 'user', content: '按日期排序' },
    { role: 'assistant', content: '' },
  ]
  it('resends the latest user message before the error card', () => {
    expect(retryTarget(msgs, 4)).toBe('按日期排序')
    expect(retryTarget(msgs, 2)).toBe('第一问')
    expect(retryTarget(msgs, 0)).toBeNull()
  })
})
