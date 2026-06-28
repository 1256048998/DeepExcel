import { useState, useRef, useEffect } from 'react'
import { sendToHost, onHostMessage } from './bridge'
import { MessageList } from './components/MessageList'
import { InputArea } from './components/InputArea'
import { StatusBar } from './components/StatusBar'
import type { Message, ConnectionStatus } from './types'

export default function App() {
  const [messages, setMessages] = useState<Message[]>([
    {
      role: 'assistant',
      content: '你好！我是 DeepExcel AI Agent，可以帮你执行Excel任务。\n\n试试说：\n• "在A1写入=SUM(B1:B10)"\n• "把Sheet1的A列数据清洗一下"\n• "根据Sheet1数据创建柱状图"'
    }
  ])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const [status, setStatus] = useState<ConnectionStatus>('connecting')
  const [isClarifying, setIsClarifying] = useState(false)

  // 监听来自C#主机的消息
  useEffect(() => {
    const unsubscribe = onHostMessage((data) => {
      if (data.type === 'stream_delta') {
        // 流式响应
        setMessages(prev => {
          const last = prev[prev.length - 1]
          if (last && last.role === 'assistant' && last.streaming) {
            return [
              ...prev.slice(0, -1),
              { ...last, content: last.content + data.payload.delta }
            ]
          }
          return [...prev, { role: 'assistant', content: data.payload.delta, streaming: true }]
        })
      } else if (data.type === 'stream_end') {
        setMessages(prev => {
          const last = prev[prev.length - 1]
          if (last && last.streaming) {
            return [...prev.slice(0, -1), { ...last, streaming: false }]
          }
          return prev
        })
        setLoading(false)
      } else if (data.type === 'tool_call') {
        // Agent 调用工具
        setMessages(prev => [...prev, {
          role: 'tool',
          content: `🔧 调用工具: ${data.payload.name}`,
          toolName: data.payload.name
        }])
      } else if (data.type === 'tool_result') {
        setMessages(prev => {
          const result = prev.map(m =>
            m.role === 'tool' && m.toolName === data.payload.name && !m.result
              ? { ...m, result: data.payload.result, streaming: false }
              : m
          )
          return result
        })
      } else if (data.type === 'clarify') {
        const { question, options } = data.payload
        const safeQuestion = question ?? ''
        const displayText = options && options.length > 0
          ? `${safeQuestion}\n\n选项：${(options as string[]).map((o, i) => `${i + 1}. ${o}`).join('\n')}`
          : safeQuestion
        setMessages(prev => [...prev, {
          role: 'assistant',
          content: displayText,
          type: 'clarify',
          options
        }])
        setIsClarifying(true)
        setLoading(false)
      } else if (data.type === 'error') {
        setMessages(prev => [...prev, { role: 'assistant', content: `❌ ${data.payload.message}` }])
        setLoading(false)
      } else if (data.type === 'connection_ok') {
        setStatus('connected')
      }
    })
    return unsubscribe
  }, [])

  const sendMessage = async () => {
    if (!input.trim() || loading) return

    const userMessage: Message = { role: 'user', content: input }
    setMessages(prev => [...prev, userMessage])
    const msg = input
    setInput('')
    setLoading(true)
    setIsClarifying(false)

    try {
      await sendToHost({
        type: 'user_message',
        payload: { content: msg }
      })
    } catch (err) {
      setMessages(prev => [...prev, { role: 'assistant', content: `❌ 发送失败: ${err}` }])
      setLoading(false)
    }
  }

  const stopGeneration = () => {
    sendToHost({ type: 'cancel', payload: {} })
    setLoading(false)
  }

  return (
    <div className="app">
      <header className="app-header">
        <div className="app-title">
          <span className="logo">◆</span>
          <span>DeepExcel</span>
        </div>
        <StatusBar status={status} />
      </header>

      <MessageList messages={messages} loading={loading} />

      <InputArea
        value={input}
        onChange={setInput}
        onSend={sendMessage}
        onStop={stopGeneration}
        disabled={loading}
        isClarifying={isClarifying}
      />
    </div>
  )
}
