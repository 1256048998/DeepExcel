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
  connecting: 'var(--de-warning-solid)',
  connected: 'var(--de-success-solid)',
  disconnected: 'var(--de-danger-solid)'
}

export function StatusBar({ status }: Props) {
  return (
    <div className="status-bar">
      <span className="status-dot" style={{ background: STATUS_COLOR[status] }}></span>
      <span className="status-text">{STATUS_TEXT[status]}</span>
    </div>
  )
}
