import { useEffect, useRef } from 'react'
import type { Message } from '../types'
import { MarkdownRenderer } from './MarkdownRenderer'
import { CopyButton } from './CopyButton'
import { StreamingChoices } from './StreamingChoices'

interface Props {
  messages: Message[]
  loading: boolean
  onToggleToolGroup?: (idx: number) => void
  onClarifyAnswer?: (answer: string) => void
  onChoiceSelect?: (choice: string) => void
  // ★ 保存用户消息为提示词模板
  onSaveAsPrompt?: (content: string) => void
}

// 检测内容是否为 Markdown 格式
function isMarkdown(content: string): boolean {
  if (!content || content.length < 2) return false
  // Markdown 常见特征：标题、列表、代码块、粗体斜体、链接、表格等
  const patterns = [
    /^#{1,6}\s/m,           // 标题 # ## ###
    /\*\*[^*]+\*\*/,         // 粗体 **text**
    /\*[^*]+\*/,            // 斜体 *text*
    /`{1,3}[^`]/m,          // 行内代码 `code` 或代码块 ```
    /^\s*[-*+]\s/m,         // 无序列表 - * +
    /^\s*\d+\.\s/m,         // 有序列表 1. 2.
    /\[.+\]\(.+\)/,         // 链接 [text](url)
    /\|.+\|/m,              // 表格 | col |
    /^---+$/m,              // 分割线 ---
    /^\s*>\s/m,             // 引用 >
    /- \[ \] /,             // 任务列表 - [ ]
  ]
  return patterns.some(p => p.test(content))
}

export function MessageList({ messages, loading, onToggleToolGroup, onClarifyAnswer, onChoiceSelect, onSaveAsPrompt }: Props) {
  const endRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    endRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  return (
    <div className="messages">
      {messages.map((msg, idx) => (
        <MessageItem
          key={idx}
          message={msg}
          index={idx}
          onToggleToolGroup={onToggleToolGroup}
          onClarifyAnswer={onClarifyAnswer}
          onChoiceSelect={onChoiceSelect}
          onSaveAsPrompt={onSaveAsPrompt}
        />
      ))}
      {loading && (
        <div className="message assistant loading">
          <span className="dot"></span>
          <span className="dot"></span>
          <span className="dot"></span>
        </div>
      )}
      <div ref={endRef} />
    </div>
  )
}

function MessageItem({
  message,
  index,
  onToggleToolGroup,
  onClarifyAnswer,
  onChoiceSelect,
  onSaveAsPrompt
}: {
  message: Message
  index: number
  onToggleToolGroup?: (idx: number) => void
  onClarifyAnswer?: (answer: string) => void
  onChoiceSelect?: (choice: string) => void
  onSaveAsPrompt?: (content: string) => void
}) {
  // 工具调用组（已合并）：折叠卡片样式
  if (message.role === 'tool' && message.toolGroup) {
    const tools = message.toolGroup
    const expanded = message.expanded ?? false
    const summary = summarizeTools(tools)
    return (
      <div className={`message tool tool-group ${expanded ? 'expanded' : 'collapsed'}`}>
        <button
          className="tool-group-header"
          onClick={() => onToggleToolGroup?.(index)}
          aria-expanded={expanded}
        >
          <span className="chevron">{expanded ? '▾' : '▸'}</span>
          <span className="tool-group-label">
            已调用 {tools.length} 个工具
          </span>
          <span className="tool-group-summary">{summary}</span>
        </button>
        {expanded && (
          <div className="tool-group-list">
            {tools.map((name, i) => (
              <div key={i} className="tool-group-item">
                <span className="tool-icon">
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"></path>
                  </svg>
                </span>
                <span className="tool-name">{name}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    )
  }

  // 普通工具消息（未被合并的孤立项）
  if (message.role === 'tool') {
    return (
      <div className="message tool">
        <div className="message-header">
          <span className="tool-icon">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"></path>
            </svg>
          </span>
          <span className="tool-name">{message.toolName}</span>
          <CopyButton content={message.result || ''} className="tool-copy-btn" />
        </div>
        {message.result && <div className="tool-result">{message.result}</div>}
      </div>
    )
  }

  // Clarify 消息：在气泡上方显示选项按钮
  if (message.type === 'clarify' && message.options && message.options.length > 0) {
    return (
      <div className="message assistant clarify-message">
        <div className="clarify-options">
          {message.options.map((opt, i) => (
            <button
              key={i}
              className="clarify-option-btn"
              onClick={() => onClarifyAnswer?.(opt)}
            >
              {opt}
            </button>
          ))}
        </div>
        <div className="message-body">
          {isMarkdown(message.content) ? (
            <MarkdownRenderer content={message.content} />
          ) : (
            <span>{message.content}</span>
          )}
          {message.streaming && <span className="cursor">▊</span>}
        </div>
      </div>
    )
  }

  // 普通用户/助手消息
  const useMarkdown = message.role === 'assistant' && isMarkdown(message.content)

  return (
    <div className={`message ${message.role}`}>
      <div className="message-header">
        {message.role === 'assistant' && <span className="role-label">助手</span>}
        {message.role === 'user' && <span className="role-label">你</span>}
        <CopyButton content={message.content} className="msg-copy-btn" />
        {/* ★ 用户消息悬停时显示"保存为提示词"按钮 */}
        {message.role === 'user' && onSaveAsPrompt && (
          <button
            className="save-prompt-btn"
            onClick={() => onSaveAsPrompt(message.content)}
            title="保存为提示词"
            type="button"
          >
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M19 21l-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z" />
            </svg>
          </button>
        )}
      </div>
      <div className="message-content">
        {useMarkdown ? (
          <MarkdownRenderer content={message.content} />
        ) : (
          <span>{message.content}</span>
        )}
        {message.streaming && <span className="cursor">▊</span>}
      </div>
      {/* ★ 流式选项卡片：assistant 消息自动检测"方案 A/B/C"模式并渲染可点击卡片 */}
      {message.role === 'assistant' && onChoiceSelect && (
        <StreamingChoices
          content={message.content}
          streaming={message.streaming}
          onSelect={onChoiceSelect}
        />
      )}
    </div>
  )
}

// 把工具名列表压缩为摘要
function summarizeTools(tools: string[]): string {
  const counts = new Map<string, number>()
  for (const t of tools) {
    counts.set(t, (counts.get(t) ?? 0) + 1)
  }
  const parts: string[] = []
  for (const [name, n] of counts) {
    parts.push(n > 1 ? `${name} ×${n}` : name)
  }
  return parts.join(', ')
}
