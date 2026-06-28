import type { ConnectionStatus } from '../types'

interface Props {
  status: ConnectionStatus
}

const STATUS_TEXT: Record<ConnectionStatus, string> = {
  connecting: '连接中',
  connected: '已连接',
  disconnected: '已断开'
}

const STATUS_COLOR: Record<ConnectionStatus, string> = {
  connecting: '#f59e0b',
  connected: '#10b981',
  disconnected: '#ef4444'
}

export function StatusBar({ status }: Props) {
  return (
    <div className="status-bar">
      <span className="status-dot" style={{ background: STATUS_COLOR[status] }}></span>
      <span className="status-text">{STATUS_TEXT[status]}</span>
    </div>
  )
}
