import { useState } from 'react'
import type { ClarifyQuestion } from '../types'
import { composeAnswer, isAnswered } from '../utils/clarify'
import type { QuestionAnswer } from '../utils/clarify'

interface Props {
  questions: ClarifyQuestion[]
  /** 已经回答过（点选提交或在输入框里直接打字）：卡片收成摘要 */
  answered?: string
  onSubmit?: (answer: string) => void
}

/**
 * 提问卡（clarify_intent）：一题或多题，单选 / 多选，每题都能选「其他」自己填，一次提交。
 * 只有一题且单选时，点选项就直接提交，不用再点一次「提交」。
 */
export function AskCard({ questions, answered, onSubmit }: Props) {
  const [answers, setAnswers] = useState<QuestionAnswer[]>(() => questions.map(() => ({ picked: [] })))
  const [otherOpen, setOtherOpen] = useState<boolean[]>(() => questions.map(() => false))
  const quick = questions.length === 1 && !questions[0].multiSelect
  const multi = questions.length > 1
  const done = answered !== undefined
  const ready = answers.every(isAnswered)

  const submit = (next: QuestionAnswer[] = answers) => {
    if (done || !onSubmit || !next.every(isAnswered)) return
    onSubmit(composeAnswer(questions, next))
  }

  const pick = (qi: number, label: string) => {
    if (done) return
    const q = questions[qi]
    const next = answers.slice()
    const cur = next[qi]
    if (q.multiSelect) {
      next[qi] = { ...cur, picked: cur.picked.includes(label) ? cur.picked.filter(l => l !== label) : [...cur.picked, label] }
    } else {
      next[qi] = { picked: [label] }
      setOtherOpen(prev => prev.map((v, i) => (i === qi ? false : v)))
    }
    setAnswers(next)
    if (quick) submit(next)
  }

  const setOther = (qi: number, text: string) => {
    setAnswers(prev => prev.map((a, i) => (i === qi
      ? { picked: questions[qi].multiSelect ? a.picked : [], other: text } : a)))
  }

  if (done) {
    return (
      <div className="ask-card answered">
        <div className="ask-card-kicker">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <polyline points="20 6 9 17 4 12" />
          </svg>
          已回答
        </div>
        {questions.length === 1 ? (
          <div className="ask-summary-row">
            <span className="ask-summary-q">{questions[0].question}</span>
            <span className="ask-summary-a">{answered}</span>
          </div>
        ) : (
          answered.split('\n').map((line, i) => {
            const sep = line.indexOf('：')
            return (
              <div key={i} className="ask-summary-row">
                <span className="ask-summary-q">{sep > 0 ? line.slice(0, sep) : ''}</span>
                <span className="ask-summary-a">{sep > 0 ? line.slice(sep + 1) : line}</span>
              </div>
            )
          })
        )}
      </div>
    )
  }

  return (
    <div className="ask-card" role="group" aria-label="需要你确认">
      <div className="ask-card-kicker">
        需要你确认{multi ? ` · ${questions.length} 个问题` : ''}
      </div>
      {questions.map((q, qi) => {
        const a = answers[qi]
        return (
          <div key={qi} className="ask-question">
            <div className="ask-question-title">
              {multi && <span className="ask-index">{qi + 1}</span>}
              {q.header && <span className="ask-header">{q.header}</span>}
              <span>{q.question}</span>
            </div>
            {q.multiSelect && <div className="ask-hint">可多选</div>}
            <div className="ask-options">
              {q.options.map(opt => {
                const on = a.picked.includes(opt.label)
                return (
                  <button
                    key={opt.label}
                    type="button"
                    className={`ask-option${on ? ' on' : ''}`}
                    role={q.multiSelect ? 'checkbox' : 'radio'}
                    aria-checked={on}
                    onClick={() => pick(qi, opt.label)}
                  >
                    <span className={`ask-mark ${q.multiSelect ? 'box' : 'dot'}`} aria-hidden="true" />
                    <span className="ask-option-text">
                      <span className="ask-option-label">{opt.label}</span>
                      {opt.description && <span className="ask-option-desc">{opt.description}</span>}
                    </span>
                  </button>
                )
              })}
              {otherOpen[qi] ? (
                <input
                  className="ask-other-input"
                  autoFocus
                  placeholder="写下你的想法，Enter 提交"
                  value={a.other ?? ''}
                  onChange={e => setOther(qi, e.target.value)}
                  onKeyDown={e => {
                    if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
                      e.preventDefault()
                      submit()
                    }
                  }}
                />
              ) : (
                <button type="button" className="ask-option other"
                  onClick={() => setOtherOpen(prev => prev.map((v, i) => (i === qi ? true : v)))}>
                  <span className="ask-mark plus" aria-hidden="true">+</span>
                  <span className="ask-option-text"><span className="ask-option-label">其他…</span></span>
                </button>
              )}
            </div>
          </div>
        )
      })}
      {(!quick || otherOpen[0]) && (
        <div className="ask-actions">
          <span className="ask-progress">
            {multi ? `已答 ${answers.filter(isAnswered).length}/${questions.length}` : ''}
          </span>
          <button type="button" className="ask-submit" disabled={!ready || !onSubmit} onClick={() => submit()}>
            提交
          </button>
        </div>
      )}
    </div>
  )
}
