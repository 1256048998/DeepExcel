import { useMemo } from 'react'

interface Props {
  content: string
  streaming?: boolean
  onSelect: (choice: string) => void
}

interface Choice {
  /** 选项标签，如 "A" / "方案 A" / "1" */
  label: string
  /** 选项描述文本（点击后发送的完整内容） */
  value: string
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

/**
 * 从消息内容中检测选项。
 * 返回值：[{label, value}, ...]，空数组表示未检测到。
 */
function detectChoices(content: string): Choice[] {
  if (!content || content.length < 4) return []

  const lines = content.split('\n')
  const choices: Choice[] = []

  // 主模式：行首匹配 "方案 A: ..." / "A. ..." / "选项 1: ..." 等
  // 支持 A-Z / 1-9 / 一二三四五 作为编号
  const patterns: RegExp[] = [
    // 方案 A: ... / 方案A：... / 方案 A、...
    /^(方案|选项|Plan|Option)\s*([A-Z一-龥])\s*[.、:：]\s*(.+)$/i,
    // A. ... / A、... / A: ... / A）... / (A) ... / 【A】...
    /^([A-Z])\s*[.、:：）)]\s*(.+)$/,
    /^\(([A-Z])\)\s*(.+)$/,
    /^【([A-Z一-龥])】\s*(.+)$/,
    // 1. ... / 1、... / 1) ... / (1) ...
    /^(\d+)\s*[.、:：）)]\s*(.+)$/,
    /^\((\d+)\)\s*(.+)$/,
    // 一、... / 二、...
    /^([一二三四五六七八九十])\s*[、.]\s*(.+)$/,
  ]

  for (const line of lines) {
    const trimmed = line.trim()
    if (!trimmed) continue

    for (const re of patterns) {
      const m = trimmed.match(re)
      if (m) {
        // m[1] 可能是前缀(方案/选项)，m[2] 是编号，m[3] 是内容
        // 或 m[1] 是编号，m[2] 是内容
        let label: string
        let value: string

        if (m.length === 4 && m[1]) {
          // 带前缀：方案 A: xxx → label="方案 A", value="xxx"
          label = `${m[1]} ${m[2]}`
          value = m[3].trim()
        } else {
          // 无前缀：A. xxx → label="A", value="xxx"
          label = m[1]
          value = m[2].trim()
        }

        // 过滤过短的伪匹配（如单字符）
        if (value.length < 1) continue
        // 过滤明显不是选项的情况：内容里包含 "？" 但很短的提问
        // 限制 value 长度，避免把整段文字当选项
        if (value.length > 120) continue

        choices.push({ label, value })
        break
      }
    }
  }

  // 去重：相同 label 只保留第一个
  const seen = new Set<string>()
  const unique: Choice[] = []
  for (const c of choices) {
    if (seen.has(c.label)) continue
    seen.add(c.label)
    unique.push(c)
  }

  return unique
}
