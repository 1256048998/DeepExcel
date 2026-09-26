// 提问卡（clarify_intent）：宿主消息 → 题目列表；用户的选择 → 发回给模型的一段话。
// 纯函数，有测试：clarify.test.ts

import type { ClarifyQuestion, Message } from '../types'

/** 新侧车发 questions（多题、选项带说明）；老消息和历史记录只有 question + options。 */
export function toQuestions(question?: string, options?: string[], questions?: unknown): ClarifyQuestion[] {
  if (Array.isArray(questions)) {
    const out: ClarifyQuestion[] = []
    for (const raw of questions) {
      if (!raw || typeof raw !== 'object') continue
      const q = raw as Record<string, unknown>
      const text = typeof q.question === 'string' ? q.question.trim() : ''
      if (!text) continue
      const opts = Array.isArray(q.options) ? q.options : []
      out.push({
        question: text,
        header: typeof q.header === 'string' ? q.header : '',
        multiSelect: q.multi_select === true,
        options: opts
          .map(o => (typeof o === 'string' ? { label: o } : o && typeof o === 'object'
            ? { label: String((o as any).label ?? ''), description: (o as any).description ? String((o as any).description) : '' }
            : { label: '' }))
          .filter(o => o.label.trim()),
      })
    }
    if (out.length > 0) return out
  }
  return [{ question: (question ?? '').trim(), header: '', multiSelect: false, options: (options ?? []).map(label => ({ label })) }]
}

export type QuestionAnswer = { picked: string[]; other?: string }

/** 一道题答了没有：选了选项，或者「其他」里写了字。 */
export function isAnswered(a?: QuestionAnswer): boolean {
  return !!a && (a.picked.length > 0 || !!a.other?.trim())
}

function answerText(a: QuestionAnswer): string {
  const parts = [...a.picked]
  if (a.other?.trim()) parts.push(a.other.trim())
  return parts.join('、')
}

/** 单题直接是答案本身；多题每行「短标签（没有就用问题）：答案」。 */
export function composeAnswer(questions: ClarifyQuestion[], answers: QuestionAnswer[]): string {
  if (questions.length <= 1) return answers[0] ? answerText(answers[0]) : ''
  return questions
    .map((q, i) => `${q.header?.trim() || q.question}：${answers[i] ? answerText(answers[i]) : '（未回答）'}`)
    .join('\n')
}

/** 用户直接在输入框里打字回答：把最近一张还没回答的提问卡标为已答。 */
export function markClarifyAnswered(messages: Message[], answer: string, index?: number): Message[] {
  let target = index ?? -1
  if (target < 0) {
    for (let i = messages.length - 1; i >= 0; i--) {
      if (messages[i].type === 'clarify' && !messages[i].answered) { target = i; break }
    }
  }
  if (target < 0 || messages[target]?.type !== 'clarify' || messages[target].answered) return messages
  const next = messages.slice()
  next[target] = { ...messages[target], answered: answer }
  return next
}
