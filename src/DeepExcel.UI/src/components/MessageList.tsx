import { useEffect, useRef } from 'react'
import type { Message } from '../types'

interface Props {
  messages: Message[]
  loading: boolean
}

export function MessageList({ messages, loading }: Props) {
  const endRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    endRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  return (
    <div className="messages">
      {messages.map((msg, idx) => (
        <MessageItem key={idx} message={msg} />
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

function MessageItem({ message }: { message: Message }) {
  return (
    <div className={`message ${message.role}`}>
      <div className="message-content">
        {message.content}
        {message.streaming && <span className="cursor">▊</span>}
      </div>
      {message.role === 'tool' && message.result && (
        <div className="tool-result">{message.result}</div>
      )}
    </div>
  )
}
