import { describe, expect, it } from 'vitest'
import type { Message } from '../types'
import { composeAnswer, isAnswered, markClarifyAnswered, toQuestions } from './clarify'

describe('toQuestions', () => {
  it('reads the multi-question payload', () => {
    const qs = toQuestions('1. 口径？\n2. 指标？', [], [
      { question: '按什么汇总？', header: '汇总口径', options: [{ label: '按月', description: '每月一行' }, '按部门'] },
      { question: '要哪些指标？', header: '指标', multi_select: true, options: ['销售额', '毛利', ''] },
      { question: '' },
      null,
    ])
    expect(qs).toHaveLength(2)
    expect(qs[0].options).toEqual([{ label: '按月', description: '每月一行' }, { label: '按部门' }])
    expect(qs[1]).toMatchObject({ multiSelect: true, options: [{ label: '销售额' }, { label: '毛利' }] })
  })

  it('falls back to the single-question fields (older sidecar, restored history)', () => {
    expect(toQuestions('求和还是计数？', ['SUM', 'COUNTA'])).toEqual([
      { question: '求和还是计数？', header: '', multiSelect: false, options: [{ label: 'SUM' }, { label: 'COUNTA' }] },
    ])
    expect(toQuestions('Q', ['A'], [])[0].question).toBe('Q')
  })
})

describe('composeAnswer', () => {
  const single = toQuestions('求和还是计数？', ['SUM', 'COUNTA'])
  const multi = toQuestions('', [], [
    { question: '按什么汇总？', header: '汇总口径', options: ['按月'] },
    { question: '要哪些指标？', options: ['销售额', '毛利'], multi_select: true },
  ])

  it('uses the bare answer for one question', () => {
    expect(composeAnswer(single, [{ picked: ['COUNTA'] }])).toBe('COUNTA')
    expect(composeAnswer(single, [{ picked: [], other: ' 都要 ' }])).toBe('都要')
  })

  it('writes one line per question for several', () => {
    expect(composeAnswer(multi, [{ picked: ['按月'] }, { picked: ['销售额', '毛利'], other: '客单价' }]))
      .toBe('汇总口径：按月\n要哪些指标？：销售额、毛利、客单价')
  })

  it('knows whether a question is answered', () => {
    expect(isAnswered(undefined)).toBe(false)
    expect(isAnswered({ picked: [], other: '  ' })).toBe(false)
    expect(isAnswered({ picked: [], other: 'x' })).toBe(true)
  })
})

describe('markClarifyAnswered', () => {
  const msgs: Message[] = [
    { role: 'assistant', type: 'clarify', content: 'Q1', answered: 'A' },
    { role: 'assistant', type: 'clarify', content: 'Q2' },
    { role: 'user', content: 'hi' },
  ]
  it('marks the latest unanswered card when the user types the answer', () => {
    const out = markClarifyAnswered(msgs, '按月')
    expect(out[1].answered).toBe('按月')
    expect(out[0].answered).toBe('A')
  })
  it('leaves the list alone when nothing is waiting', () => {
    const done = markClarifyAnswered(msgs, 'x')
    expect(markClarifyAnswered(done, 'y')).toBe(done)
  })
})
