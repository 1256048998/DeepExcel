import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { Message, PermissionMode, ToolStep, ToolStepChanges } from '../types'
import { formatDuration } from '../utils/uiEvents'
import { useStickToBottom } from '../utils/useStickToBottom'
import { MarkdownRenderer } from './MarkdownRenderer'
import { CopyButton } from './CopyButton'
import { StreamingChoices } from './StreamingChoices'
import { StarterCard } from './StarterCard'
import { LogoMark } from './Logo'
import type { StarterView } from './StarterCard'

interface Props {
  messages: Message[]
  loading: boolean
  statusText?: string | null
  onToggleToolGroup?: (idx: number) => void
  onClarifyAnswer?: (answer: string) => void
  onChoiceSelect?: (choice: string) => void
  // ★ 保存用户消息为提示词模板
  onSaveAsPrompt?: (content: string) => void
  // 「回到这一步之前」：任务进行中不可用
  rewind?: RewindControl
  // 方案卡片：批准（按哪种模式执行）或继续修改
  onPlanDecision?: (index: number, decision: PermissionMode | 'dismissed') => void
  // 首次使用卡片：点推荐问题直接发送；空工作簿可插入示例数据
  starter?: StarterView
  onStarterPick?: (prompt: string) => void
  onInsertSample?: () => void
}

export type RewindControl = {
  onRewind: (step: ToolStep) => void
  disabled: boolean
  pendingId: string | null
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

export function MessageList({ messages, loading, statusText, onToggleToolGroup, onClarifyAnswer, onChoiceSelect, onSaveAsPrompt, rewind, onPlanDecision, starter, onStarterPick, onInsertSample }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)
  const contentRef = useRef<HTMLDivElement>(null)

  // 回调每次 App 渲染都是新函数；包一层稳定引用，消息项才能 memo（流式输出时只重绘最后一条）
  const latest = useRef({ onToggleToolGroup, onClarifyAnswer, onChoiceSelect, onSaveAsPrompt, onRewind: rewind?.onRewind, onPlanDecision })
  latest.current = { onToggleToolGroup, onClarifyAnswer, onChoiceSelect, onSaveAsPrompt, onRewind: rewind?.onRewind, onPlanDecision }
  const toggle = useCallback((i: number) => latest.current.onToggleToolGroup?.(i), [])
  const clarify = useCallback((a: string) => latest.current.onClarifyAnswer?.(a), [])
  const choose = useCallback((c: string) => latest.current.onChoiceSelect?.(c), [])
  const savePrompt = useCallback((c: string) => latest.current.onSaveAsPrompt?.(c), [])
  const onRewind = useCallback((step: ToolStep) => latest.current.onRewind?.(step), [])
  const decidePlan = useCallback(
    (i: number, d: PermissionMode | 'dismissed') => latest.current.onPlanDecision?.(i, d), [])
  const hasRewind = !!rewind
  const rewindDisabled = rewind?.disabled ?? true
  const rewindPending = rewind?.pendingId ?? null
  const stableRewind = useMemo<RewindControl | undefined>(
    () => (hasRewind ? { onRewind, disabled: rewindDisabled, pendingId: rewindPending } : undefined),
    [hasRewind, onRewind, rewindDisabled, rewindPending],
  )

  // 用户刚发出一条消息：不管之前翻到哪里，都回到底部
  const lastLength = useRef(messages.length)
  const forcePin = messages.length > lastLength.current && messages[messages.length - 1]?.role === 'user'
  useEffect(() => { lastLength.current = messages.length }, [messages.length])

  const contentKey = useMemo(() => ({}), [messages, loading, statusText])
  const { pinned, jumpToBottom } = useStickToBottom(containerRef, contentRef, contentKey, forcePin)

  return (
    <div className="messages-wrap">
      <div className="messages" ref={containerRef}>
        <div className="messages-content" ref={contentRef}>
          {messages.map((msg, idx) => msg.type === 'starter' ? (
            <div key={idx} className="message assistant starter-message">
              <div className="starter-hero">
                <LogoMark size={28} />
                <div className="message-content"><span>{msg.content}</span></div>
              </div>
              {starter && (
                <StarterCard
                  starter={starter}
                  busy={loading}
                  onPick={prompt => onStarterPick?.(prompt)}
                  onInsertSample={() => onInsertSample?.()}
                />
              )}
            </div>
          ) : (
            <MessageItem
              key={idx}
              message={msg}
              index={idx}
              onToggleToolGroup={onToggleToolGroup ? toggle : undefined}
              onClarifyAnswer={onClarifyAnswer ? clarify : undefined}
              onChoiceSelect={onChoiceSelect ? choose : undefined}
              onSaveAsPrompt={onSaveAsPrompt ? savePrompt : undefined}
              rewind={stableRewind}
              onPlanDecision={onPlanDecision ? decidePlan : undefined}
              busy={loading}
            />
          ))}
          {loading && (
            <div className="message assistant loading">
              <span className="dot"></span>
              <span className="dot"></span>
              <span className="dot"></span>
              {statusText && <span className="loading-status">{statusText}</span>}
            </div>
          )}
        </div>
      </div>
      {!pinned && (
        <button type="button" className="jump-to-bottom" onClick={jumpToBottom} title="回到底部">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <line x1="12" y1="5" x2="12" y2="19" />
            <polyline points="19 12 12 19 5 12" />
          </svg>
          回到底部
        </button>
      )}
    </div>
  )
}

const MessageItem = memo(function MessageItem({
  message,
  index,
  onToggleToolGroup,
  onClarifyAnswer,
  onChoiceSelect,
  onSaveAsPrompt,
  rewind,
  onPlanDecision,
  busy
}: {
  message: Message
  index: number
  onToggleToolGroup?: (idx: number) => void
  onClarifyAnswer?: (answer: string) => void
  onChoiceSelect?: (choice: string) => void
  onSaveAsPrompt?: (content: string) => void
  rewind?: RewindControl
  onPlanDecision?: (index: number, decision: PermissionMode | 'dismissed') => void
  busy?: boolean
}) {
  // 方案卡片（present_plan）：涉及的表和区域、每一步、风险点，下面是批准按钮
  if (message.type === 'plan_proposal' && message.plan) {
    return <PlanCard message={message} index={index} busy={!!busy} onDecision={onPlanDecision} />
  }

  // 思考过程：生成时展开，结束后收成「思考了 N 秒」
  if (message.type === 'thinking') {
    return <ThinkingCard message={message} />
  }

  // 本次会话实时收到的工具步骤：每步一行叙事（● 读取 A1:D20，下面一行写结果：20 行 × 4 列）
  if (message.role === 'tool' && message.toolSteps && message.toolSteps.length > 0) {
    return (
      <ToolSteps
        steps={message.toolSteps}
        expanded={message.expanded ?? false}
        onToggle={() => onToggleToolGroup?.(index)}
        rewind={rewind}
      />
    )
  }

  // 错误卡片：分类后的原因 + 下一步，而不是英文异常原文
  if (message.type === 'error' && message.error) {
    return (
      <div className="message error-card" role="alert">
        <div className="error-card-title">{message.error.message}</div>
        {message.error.hint && <div className="error-card-hint">{message.error.hint}</div>}
      </div>
    )
  }

  // 终态行：只在失败、中断、轮次用尽或多步任务后出现
  if (message.type === 'run_summary') {
    return (
      <div className={`message run-summary outcome-${message.outcome ?? 'success'}`}>
        {message.content}
      </div>
    )
  }

  // 工具调用组（从历史恢复，只有工具名）：折叠卡片样式
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
          <Chevron open={expanded} />
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

  // ★ 压缩提示卡：autocompact 触发时插入的轻量提示
  if (message.type === 'compacted' || message.type === 'notice') {
    return (
      <div className="message compacted-hint">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: 6, verticalAlign: 'middle' }}>
          <path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
          <polyline points="21 3 21 8 16 8" />
        </svg>
        {message.content}
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
      {/* 复制 / 存为提示词：鼠标移上来才出现，不单独占一行 */}
      <div className="message-actions">
        <CopyButton content={message.content} className="msg-copy-btn" />
        {/* ★ 用户消息悬停时显示"保存为提示词"按钮 */}
        {message.role === 'user' && onSaveAsPrompt && (
          <button
            className="save-prompt-btn"
            onClick={() => onSaveAsPrompt(message.content)}
            title="保存为提示词/技能"
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
      {message.queued && (
        <div className={`queued-label queued-${message.queued}`}>
          {message.queued === 'pending' ? '等待送达' : message.queued === 'delivered' ? '已交给 AI' : '本轮结束后处理'}
        </div>
      )}
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
})

// 连续多步时默认只显示最后几步，避免一次任务把整屏推走；点击标题展开全部
const VISIBLE_STEPS_WHEN_COLLAPSED = 3

function ToolSteps({ steps, expanded, onToggle, rewind }: {
  steps: ToolStep[]
  expanded: boolean
  onToggle: () => void
  rewind?: RewindControl
}) {
  const failed = steps.filter(step => step.status === 'error').length
  const running = steps.some(step => step.status === 'running')
  const collapsible = steps.length > VISIBLE_STEPS_WHEN_COLLAPSED
  const visible = !collapsible || expanded ? steps : steps.slice(-VISIBLE_STEPS_WHEN_COLLAPSED)
  const hidden = steps.length - visible.length
  return (
    <div className={`message tool tool-steps${running ? ' running' : ''}`}>
      {collapsible && (
        <button className="tool-steps-header" onClick={onToggle} aria-expanded={expanded} type="button">
          <span>{expanded ? '收起步骤' : `显示全部 ${steps.length} 步`}</span>
          {failed > 0 && <span className="tool-steps-failed">{failed} 步失败</span>}
          {!expanded && hidden > 0 && <span className="tool-steps-hidden">已折叠 {hidden} 步</span>}
          <Chevron open={expanded} />
        </button>
      )}
      {visible.map(step => <ToolStepLine key={step.id} step={step} rewind={rewind} />)}
    </div>
  )
}

// 生成中的代码只显示最后几行，跟着往下走；写完后折叠，需要时点开看全文
const LIVE_CODE_LINES = 12

function lastLines(text: string, n: number): string {
  const lines = text.split('\n')
  return lines.slice(-n).join('\n')
}

const REWIND_TOOLTIP = '撤销这一步和它之后的所有修改（包括其他表上的）。回退前的状态会另存一份，可在「历史版本」里找回。'

function ToolStepLine({ step, rewind }: { step: ToolStep; rewind?: RewindControl }) {
  const busy = step.status === 'running' || step.status === 'generating'
  // 一秒以内的步骤不写耗时：用户关心的是慢在哪一步，不是每步几毫秒
  const duration = !busy && (step.durationMs ?? 0) >= 1000 ? formatDuration(step.durationMs) : ''
  return (
    <div className={`tool-step status-${step.status}${step.check && !step.check.ok ? ' has-warning' : ''}`}>
      <div className="tool-step-line">
        <span className="tool-step-dot" aria-hidden="true" />
        <span className="tool-step-label" title={step.name}>{step.label}</span>
        {duration && <span className="tool-step-duration">{duration}</span>}
      </div>
      {step.code && step.status === 'generating' && (
        <pre className="tool-step-code live">{lastLines(step.code, LIVE_CODE_LINES)}</pre>
      )}
      {step.code && step.status !== 'generating' && (
        <details className="tool-step-code-toggle">
          <summary>查看代码<Chevron /></summary>
          <pre className="tool-step-code">{step.code}</pre>
        </details>
      )}
      {step.status === 'ok' && step.summary && (
        <div className="tool-step-result">{step.summary}</div>
      )}
      {step.status === 'ok' && step.changes && <ChangesTable changes={step.changes} />}
      {step.status === 'ok' && step.check && !step.check.ok && (
        <div className="tool-step-result check-failed">{step.check.summary}</div>
      )}
      {/* 常驻在这一步下面：「放心让 AI 改表」的底气要看得见，不藏在悬停里 */}
      {rewind && step.status === 'ok' && step.checkpointId && (
        <button
          type="button"
          className={`tool-step-rewind${rewind.pendingId === step.checkpointId ? ' pending' : ''}`}
          title={REWIND_TOOLTIP}
          disabled={rewind.disabled || rewind.pendingId !== null}
          onClick={() => rewind.onRewind(step)}
        >
          <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <polyline points="9 14 4 9 9 4" />
            <path d="M20 20v-7a4 4 0 0 0-4-4H4" />
          </svg>
          {rewind.pendingId === step.checkpointId ? '回退中…' : '回到这一步之前'}
        </button>
      )}
      {step.status === 'error' && step.error && (
        <div className="tool-step-result error">
          {step.error.message}
          {step.error.hint && <div className="tool-step-hint">{step.error.hint}</div>}
        </div>
      )}
    </div>
  )
}

function ThinkingCard({ message }: { message: Message }) {
  const active = !!message.thinkingActive
  // 没手动点过就跟着状态走：思考中展开，结束收起；点过以后听用户的
  const [userOpen, setUserOpen] = useState<boolean | null>(null)
  const open = userOpen ?? active
  const bodyRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const el = bodyRef.current
    if (el && active) el.scrollTop = el.scrollHeight
  }, [message.content, active])
  if (!active && !message.content.trim()) return null
  const seconds = message.thinkingMs ? Math.max(1, Math.round(message.thinkingMs / 1000)) : 0
  return (
    <div className={`message thinking-card${active ? ' active' : ''}`}>
      <button type="button" className="thinking-header" onClick={() => setUserOpen(!open)} aria-expanded={open}>
        <svg className="thinking-icon" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"
          strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="M12 3c.4 3.6 1.9 5.1 5.5 5.5-3.6.4-5.1 1.9-5.5 5.5-.4-3.6-1.9-5.1-5.5-5.5C10.1 8.1 11.6 6.6 12 3z" />
          <path d="M18.5 14c.2 1.8 1 2.6 2.8 2.8-1.8.2-2.6 1-2.8 2.8-.2-1.8-1-2.6-2.8-2.8 1.8-.2 2.6-1 2.8-2.8z" />
        </svg>
        <span className="thinking-title">{active ? '思考中…' : seconds ? `思考了 ${seconds} 秒` : '思考过程'}</span>
        <Chevron open={open} />
      </button>
      {open && message.content && (
        <div className="thinking-body" ref={bodyRef}>{message.content}</div>
      )}
    </div>
  )
}

// 折叠箭头：用 SVG，不用 ▸ ▾ 字符（各字体里大小、基线都不一样）
function Chevron({ open }: { open?: boolean }) {
  return (
    <svg className={`chevron-icon${open ? ' open' : ''}`} width="10" height="10" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <polyline points="9 6 15 12 9 18" />
    </svg>
  )
}

// 内联 diff：默认折叠成「改了 N 格」，点开是原值 → 新值的小表
function ChangesTable({ changes }: { changes: ToolStepChanges }) {
  const more = changes.changed - changes.samples.length
  return (
    <details className="tool-step-diff">
      <summary>改了 {changes.changed} 格<Chevron /></summary>
      {changes.samples.length > 0 && (
        <table>
          <thead>
            <tr><th>单元格</th><th>原值</th><th>新值</th></tr>
          </thead>
          <tbody>
            {changes.samples.map(change => (
              <tr key={change.address}>
                <td className="addr">{change.address}</td>
                <td className={change.before ? '' : 'empty'}>{change.before || '（空）'}</td>
                <td className={change.after ? '' : 'empty'}>{change.after || '（空）'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {more > 0 && <div className="tool-step-diff-more">另有 {more} 格未列出</div>}
    </details>
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

const DECISION_TEXT: Record<string, string> = {
  default: '已批准，按每步确认执行',
  accept_writes: '已批准，本次会话自动应用写入',
  plan: '已批准',
  dismissed: '已放弃这个方案，可以继续说你的修改意见',
}

function PlanCard({ message, index, busy, onDecision }: {
  message: Message
  index: number
  busy: boolean
  onDecision?: (index: number, decision: PermissionMode | 'dismissed') => void
}) {
  const plan = message.plan!
  const decided = message.planDecision
  return (
    <div className={`message plan-card${decided ? ' decided' : ''}`}>
      <div className="plan-card-title">{decided ? '方案' : '方案待批准'}</div>
      <div className="plan-card-summary">{plan.summary}</div>
      <ol className="plan-card-steps">
        {plan.steps.map((step, i) => (
          <li key={i}>
            <span className="plan-step-action">{step.action}</span>
            {step.target && <span className="plan-step-target">{step.target}</span>}
            {step.detail && <div className="plan-step-detail">{step.detail}</div>}
          </li>
        ))}
      </ol>
      {plan.risks.length > 0 && (
        <div className="plan-card-risks">
          <div className="plan-card-risks-title">风险点</div>
          <ul>{plan.risks.map((r, i) => <li key={i}>{r}</li>)}</ul>
        </div>
      )}
      {decided ? (
        <div className="plan-card-decision">{DECISION_TEXT[decided] ?? '已处理'}</div>
      ) : onDecision && (
        <div className="plan-card-actions">
          <button type="button" className="plan-btn primary" disabled={busy}
            onClick={() => onDecision(index, 'default')}>批准并执行</button>
          <button type="button" className="plan-btn" disabled={busy}
            onClick={() => onDecision(index, 'accept_writes')}
            title="本次会话里写入不再逐个确认；删除、跑代码仍会确认">批准并自动应用</button>
          <button type="button" className="plan-btn ghost" disabled={busy}
            onClick={() => onDecision(index, 'dismissed')}>继续修改</button>
        </div>
      )}
    </div>
  )
}
