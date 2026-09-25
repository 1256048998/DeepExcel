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

function findStep(messages: Message[], id: string): ToolStep | undefined {
  for (let i = messages.length - 1; i >= 0; i--) {
    const step = messages[i].toolSteps?.find(s => s.id === id)
    if (step) return step
  }
  return undefined
}

const CODE_ARG: Record<string, string> = { execute_vba: 'code', execute_jsa: 'code', execute_python: 'code' }

function generatingLabel(name: string, lines?: number, chars?: number): string {
  const what = name === 'execute_vba' ? 'VBA' : name === 'execute_jsa' ? 'JS 宏'
    : name === 'execute_python' ? 'Python' : ''
  if (what) return `正在编写 ${what}${lines ? `（${lines} 行）` : ''}…`
  return `正在准备「${toolLabel(name)}」${chars ? `（${chars} 字）` : ''}…`
}

function appendStep(messages: Message[], step: ToolStep): Message[] {
  const base = closeStreaming(messages)
  const idx = lastToolGroupIndex(base)
  if (idx >= 0) {
    const group = base[idx]
    const next = base.slice()
    next[idx] = {
      ...group,
      toolSteps: [...(group.toolSteps ?? []), step],
      toolGroup: [...(group.toolGroup ?? []), step.name],
    }
    return next
  }
  return [...base, {
    role: 'tool',
    content: '',
    toolName: step.name,
    toolGroup: [step.name],
    toolSteps: [step],
    // 超过 3 步时默认只显示最后 3 步，点标题展开
    expanded: false,
  }]
}

// 这些工具有专门的展示位置（计划胶囊），不再作为一步显示在步骤列表里
const HIDDEN_STEP_TOOLS = new Set(['todo_write'])

export function applyUiEvent(messages: Message[], event: UiEvent): Message[] {
  if ((event.kind === 'tool_gen' || event.kind === 'tool_start') && HIDDEN_STEP_TOOLS.has(event.name)) {
    return messages
  }
  switch (event.kind) {
    case 'tool_gen': {
      const patch = {
        label: generatingLabel(event.name, event.lines, event.chars),
        code: event.preview,
        genChars: event.chars,
      }
      if (findStep(messages, event.id)) return updateStep(messages, event.id, patch)
      return appendStep(messages, { id: event.id, name: event.name, status: 'generating', ...patch })
    }
    case 'tool_start': {
      const codeArg = CODE_ARG[event.name]
      const code = codeArg && typeof event.args?.[codeArg] === 'string' ? event.args[codeArg] as string : undefined
      const fields = {
        name: event.name,
        label: toolLabel(event.name, event.args),
        args: event.args,
        status: 'running' as const,
        code,
      }
      // 参数边写边显示过的步骤：原地变成「执行中」，不另起一行
      if (findStep(messages, event.id)) return updateStep(closeStreaming(messages), event.id, fields)
      return appendStep(messages, { id: event.id, ...fields })
    }
    case 'tool_end':
      return updateStep(messages, event.id, {
        status: event.ok ? 'ok' : 'error',
        summary: event.summary,
        error: event.error,
        durationMs: event.duration_ms,
        check: event.check,
        checkpointId: event.checkpoint_id,
        changes: event.changes,
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
