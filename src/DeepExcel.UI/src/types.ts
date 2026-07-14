export type Message = {
  role: 'user' | 'assistant' | 'tool'
  content: string
  streaming?: boolean
  toolName?: string
  result?: string
  type?: 'clarify' | 'compacted'
  options?: string[]
  // 折叠工具调用组：当 role==='tool' 且是连续工具调用的首条时，
  // toolGroup 存该组所有工具名（按调用顺序），后续同组 tool 消息会被合并
  toolGroup?: string[]
  // 该工具组是否处于展开状态（默认 false 折叠）
  expanded?: boolean
}

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
}

export type ModelConfig = {
  currentProvider: string
  currentModel: string
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
