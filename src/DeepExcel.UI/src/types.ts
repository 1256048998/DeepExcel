export type Message = {
  role: 'user' | 'assistant' | 'tool'
  content: string
  streaming?: boolean
  toolName?: string
  result?: string
  type?: 'clarify' | 'compacted' | 'error' | 'run_summary'
  options?: string[]
  // 折叠工具调用组：当 role==='tool' 且是连续工具调用的首条时，
  // toolGroup 存该组所有工具名（按调用顺序），后续同组 tool 消息会被合并
  toolGroup?: string[]
  // 本次会话里实时收到的工具步骤（ui_event）；从历史恢复的旧消息只有 toolGroup
  toolSteps?: ToolStep[]
  // 该工具组是否处于展开状态
  expanded?: boolean
  // type==='error'
  error?: { code: string; message: string; hint?: string; retryable?: boolean }
  // type==='run_summary'
  outcome?: string
}

export type ToolStepError = { code: string; message: string; hint?: string | null }

export type ToolStep = {
  id: string
  name: string
  label: string
  args?: Record<string, unknown>
  status: 'running' | 'ok' | 'error'
  summary?: string
  error?: ToolStepError
  durationMs?: number
}

// 侧车 ui_event 信封里的 event（docs/ui-event-protocol.md）
export type UiEvent =
  | { v: number; kind: 'tool_start'; id: string; name: string; args?: Record<string, unknown>; ts?: number }
  | { v: number; kind: 'tool_end'; id: string; name: string; ok: boolean; duration_ms?: number;
      summary?: string; error?: ToolStepError; ts?: number }
  | { v: number; kind: 'status'; text: string; tool?: string; ts?: number }
  | { v: number; kind: 'compaction'; trigger: string; pre_tokens?: number; prev_pct?: number;
      curr_pct?: number; ts?: number }
  | { v: number; kind: 'error'; code: string; message: string; hint?: string; retryable?: boolean;
      detail?: string; ts?: number }
  | { v: number; kind: 'run_summary'; outcome: string; tool_calls: number; failed_calls: number;
      duration_ms: number; num_turns?: number; input_tokens?: number; output_tokens?: number; ts?: number }

export type ConnectionStatus = 'connecting' | 'connected' | 'disconnected'

// ★ 历史对话元信息（列表展示用，不含 messages）
export type ConversationSummary = {
  id: string
  title: string
  createdAt: string
  updatedAt?: string
  workbookName?: string
}

// ★ 模型配置弹窗用类型（对应 C# SafeConfig/SafeProvider 结构）

export type ProviderInfo = {
  displayName: string
  type: string
  baseUrl: string
  defaultModel: string
  supportsVision: boolean
  models: string[]
  hasApiKey: boolean
  apiKeyPreview: string
  /** ★ 最近一次测试连接是否成功（前端圆点据此显示，而非 hasApiKey） */
  connected: boolean
}

export type ModelConfig = {
  currentProvider: string
  currentModel: string
  /** ★ 全局默认厂商（前端 provider 列表排序最前，输入框下拉默认值取其 DefaultModel）。
   * 后端 Guaranteed non-null（SafeConfig 中 DefaultProvider 为空时回退到 CurrentProvider） */
  defaultProvider: string
  providers: Record<string, ProviderInfo>
  general: {
    maxRetries: number
    requestTimeoutSeconds: number
    autoCreateSnapshot: boolean
    requireConfirmation: boolean
    maxConversationHistory: number
    maxTurns: number
  }
  ui: {
    theme: string
    language: string
    showTokenUsage: boolean
    streamOutput: boolean
  }
}
