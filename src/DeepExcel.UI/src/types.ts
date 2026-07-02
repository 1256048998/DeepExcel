export type Message = {
  role: 'user' | 'assistant' | 'tool'
  content: string
  streaming?: boolean
  toolName?: string
  result?: string
  type?: 'clarify'
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
