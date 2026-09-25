// ui_event 信封 → 消息列表（协议见 docs/ui-event-protocol.md）。
//
// 纯函数：输入旧消息列表和一个事件，返回新列表。放在组件外面是为了能单测——
// 工具开始/结束的配对、中断时补结束、错误卡片，这些以前全靠肉眼在 Excel 里验。

import type { Message, ToolStep, UiEvent } from '../types'
import { toolLabel } from './toolCatalog'

function lastToolGroupIndex(messages: Message[]): number {
  const last = messages.length - 1
  return last >= 0 && messages[last].role === 'tool' && messages[last].toolSteps ? last : -1
}

function updateStep(messages: Message[], id: string, patch: Partial<ToolStep>): Message[] {
  for (let i = messages.length - 1; i >= 0; i--) {
    const steps = messages[i].toolSteps
    if (!steps) continue
    const j = steps.findIndex(step => step.id === id)
    if (j < 0) continue
    const nextSteps = steps.slice()
    nextSteps[j] = { ...steps[j], ...patch }
    const next = messages.slice()
    next[i] = { ...messages[i], toolSteps: nextSteps }
    return next
  }
  return messages
}

/** 流式中的助手文本在工具步骤插入前要先封口，否则光标一直闪。 */
function closeStreaming(messages: Message[]): Message[] {
  return messages.some(m => m.streaming)
    ? messages.map(m => (m.streaming ? { ...m, streaming: false } : m))
    : messages
}

export function applyUiEvent(messages: Message[], event: UiEvent): Message[] {
  switch (event.kind) {
    case 'tool_start': {
      const step: ToolStep = {
        id: event.id,
        name: event.name,
        label: toolLabel(event.name, event.args),
        args: event.args,
        status: 'running',
      }
      const base = closeStreaming(messages)
      const idx = lastToolGroupIndex(base)
      if (idx >= 0) {
        const group = base[idx]
        const next = base.slice()
        next[idx] = {
          ...group,
          toolSteps: [...(group.toolSteps ?? []), step],
          toolGroup: [...(group.toolGroup ?? []), event.name],
        }
        return next
      }
      return [...base, {
        role: 'tool',
        content: '',
        toolName: event.name,
        toolGroup: [event.name],
        toolSteps: [step],
        // 超过 3 步时默认只显示最后 3 步，点标题展开
        expanded: false,
      }]
    }
    case 'tool_end':
      return updateStep(messages, event.id, {
        status: event.ok ? 'ok' : 'error',
        summary: event.summary,
        error: event.error,
        durationMs: event.duration_ms,
      })
    case 'compaction':
      return [...messages, {
        role: 'assistant',
        type: 'compacted',
        content: event.pre_tokens
          ? `对话已自动压缩（压缩前约 ${Math.round(event.pre_tokens / 1000)}k tokens），保留了关键上下文。`
          : '对话已自动压缩，保留了关键上下文。',
      }]
    case 'error':
      return [...closeStreaming(messages), {
        role: 'assistant',
        type: 'error',
        content: event.message,
        error: { code: event.code, message: event.message, hint: event.hint, retryable: event.retryable },
      }]
    case 'steer_delivered':
    case 'steer_deferred': {
      const next = event.kind === 'steer_delivered' ? 'delivered' : 'deferred'
      return messages.some(m => m.queued === 'pending')
        ? messages.map(m => (m.queued === 'pending' ? { ...m, queued: next } : m))
        : messages
    }
    case 'run_summary': {
      // 只在值得说的时候出现：失败、中断、达到轮次上限，或者做了不少步。
      // 一问一答的闲聊后面挂一行「完成」只是噪音。
      const notable = event.outcome !== 'success' || event.tool_calls >= 3 || event.failed_calls > 0
      if (!notable) return messages
      return [...closeStreaming(messages), {
        role: 'assistant',
        type: 'run_summary',
        content: runSummaryText(event),
        outcome: event.outcome,
      }]
    }
    default:
      return messages
  }
}

export function formatDuration(ms?: number): string {
  if (ms === undefined || ms === null) return ''
  if (ms < 1000) return `${ms}ms`
  if (ms < 60_000) return `${(ms / 1000).toFixed(1)}s`
  const m = Math.floor(ms / 60_000)
  const sec = Math.round((ms % 60_000) / 1000)
  return `${m}分${sec}秒`
}

const OUTCOME_TEXT: Record<string, string> = {
  success: '完成',
  interrupted: '已中断',
  max_turns: '达到最大轮次，任务可能没做完——回复「继续」接着做',
  error: '出错结束',
}

export function runSummaryText(e: Extract<UiEvent, { kind: 'run_summary' }>): string {
  const parts = [OUTCOME_TEXT[e.outcome] ?? e.outcome]
  if (e.tool_calls > 0) {
    parts.push(e.failed_calls > 0 ? `${e.tool_calls} 步（${e.failed_calls} 步失败）` : `${e.tool_calls} 步`)
  }
  const d = formatDuration(e.duration_ms)
  if (d) parts.push(d)
  return parts.join(' · ')
}
