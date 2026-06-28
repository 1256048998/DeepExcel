export type Message = {
  role: 'user' | 'assistant' | 'tool'
  content: string
  streaming?: boolean
  toolName?: string
  result?: string
  type?: 'clarify'
  options?: string[]
}

export type ConnectionStatus = 'connecting' | 'connected' | 'disconnected'
