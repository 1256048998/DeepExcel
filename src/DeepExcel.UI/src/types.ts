export type Message = {
  role: 'user' | 'assistant' | 'tool'
  content: string
  streaming?: boolean
  toolName?: string
  result?: string
  // starter：欢迎语，下面带「首次使用」卡片（示例数据 / 为你的文件推荐）
  type?: 'clarify' | 'compacted' | 'error' | 'run_summary' | 'notice' | 'plan_proposal' | 'starter' | 'thinking'
  options?: string[]
  // type==='clarify'：提问卡的题目（多题、选项说明）；answered 是已发出的回答
  questions?: ClarifyQuestion[]
  answered?: string
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
  // 任务进行中发出的插话：pending 已发出 / delivered 已交给 AI / deferred 本轮结束后处理
  queued?: 'pending' | 'delivered' | 'deferred'
  // type==='plan_proposal'：模型用 present_plan 提交的方案，和用户的决定
  plan?: PlanProposal
  planDecision?: PermissionMode | 'dismissed'
  // type==='thinking'：思考过程卡片（content 是思考文字）
  thinkingId?: string
  thinkingActive?: boolean
  thinkingMs?: number
}

export type ClarifyOption = { label: string; description?: string }
export type ClarifyQuestion = { question: string; header?: string; multiSelect?: boolean; options: ClarifyOption[] }

// 权限模式（侧车 permission_modes.py）：每步确认 / 本次会话自动应用写入 / 只出方案。只在内存里，不保存
export type PermissionMode = 'default' | 'accept_writes' | 'plan'

export type PlanStep = { action: string; target?: string; detail?: string }
export type PlanProposal = { summary: string; steps: PlanStep[]; risks: string[] }

export type PlanItem = { content: string; status: 'pending' | 'in_progress' | 'completed' }

export type ToolStepError = { code: string; message: string; hint?: string | null }

export type ToolStep = {
  id: string
  name: string
  label: string
  args?: Record<string, unknown>
  // generating：模型还在生成参数（代码边写边显示），尚未真正调用
  status: 'generating' | 'running' | 'ok' | 'error'
  // 代码类工具（VBA / JS 宏 / Python）的代码；生成中是目前写到的部分
  code?: string
  genChars?: number
  summary?: string
  error?: ToolStepError
  durationMs?: number
  // 写后自动体检（新增公式错误、外部链接）；只在没通过时显示
  check?: ToolStepCheck
  // 这一步执行前单独存的检查点：有它才显示「回到这一步之前」
  checkpointId?: string
  // 内联 diff：这一步实际改了哪些格
  changes?: ToolStepChanges
}

export type ToolStepCheck = { ok: boolean; summary: string }

export type CellChange = { address: string; before: string; after: string }
export type ToolStepChanges = { changed: number; sheet?: string; samples: CellChange[] }

// 侧车 ui_event 信封里的 event（docs/ui-event-protocol.md）
export type UiEvent =
  | { v: number; kind: 'tool_start'; id: string; name: string; args?: Record<string, unknown>; ts?: number }
  | { v: number; kind: 'tool_end'; id: string; name: string; ok: boolean; duration_ms?: number;
      summary?: string; error?: ToolStepError; check?: ToolStepCheck; checkpoint_id?: string; changes?: ToolStepChanges; ts?: number }
  | { v: number; kind: 'status'; text: string; tool?: string; ts?: number }
  | { v: number; kind: 'tool_gen'; id: string; name: string; chars: number; lines?: number;
      preview?: string; ts?: number }
  | { v: number; kind: 'compaction'; trigger: string; pre_tokens?: number; prev_pct?: number;
      curr_pct?: number; ts?: number }
  | { v: number; kind: 'error'; code: string; message: string; hint?: string; retryable?: boolean;
      detail?: string; ts?: number }
  | { v: number; kind: 'steer_delivered' | 'steer_deferred'; count: number; ts?: number }
  | { v: number; kind: 'plan'; items: PlanItem[]; ts?: number }
  | { v: number; kind: 'thinking_start'; id: string; ts?: number }
  | { v: number; kind: 'thinking_delta'; id: string; text: string; ts?: number }
  | { v: number; kind: 'thinking_end'; id: string; chars?: number; duration_ms?: number; ts?: number }
  | { v: number; kind: 'plan_proposal'; summary: string; steps: PlanStep[]; risks: string[]; ts?: number }
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
