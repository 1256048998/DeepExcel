import { useMemo } from 'react'
import { detectChoices } from '../utils/choices'

interface Props {
  content: string
  streaming?: boolean
  onSelect: (choice: string) => void
}

/**
 * 检测 AI 流式响应中的选项模式，渲染为可点击卡片。
 *
 * 支持的模式：
 *   方案 A: ...        方案 B: ...
 *   A. ...             B. ...
 *   A、...             B、...
 *   选项 1: ...        选项 2: ...
 *   (A) ...            (B) ...
 *   【A】...           【B】...
 *
 * 普通编号列表只有在前一行让用户做选择时才算（见 utils/choices.ts）。
 * 仅当检测到 ≥2 个选项时显示卡片。
 * 流式输出时也实时更新（每次 content 变化重新解析）。
 */
export function StreamingChoices({ content, streaming, onSelect }: Props) {
  const choices = useMemo(() => detectChoices(content), [content])

  if (choices.length < 2) return null

  return (
    <div className="streaming-choices">
      <div className="streaming-choices-label">
        <svg width="11" height="11" viewBox="0 0 16 16" fill="none">
          <path
            d="M5 3l5 5-5 5"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
        <span>选择方案</span>
      </div>
      <div className="streaming-choices-list">
        {choices.map((c, i) => (
          <button
            key={i}
            className="streaming-choice-card"
            onClick={() => onSelect(c.value)}
            type="button"
          >
            <span className="choice-tag">{c.label}</span>
            <span className="choice-desc">{c.value}</span>
          </button>
        ))}
      </div>
      {streaming && (
        <div className="streaming-choices-hint">AI 仍在输出，可立即选择或等待</div>
      )}
    </div>
  )
}
