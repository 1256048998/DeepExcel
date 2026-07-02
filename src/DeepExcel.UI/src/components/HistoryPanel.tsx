import { useState, useEffect } from 'react'
import { sendToHost, onHostMessage } from '../bridge'

export interface Snapshot {
  id: string
  workbookName: string
  createdAt: string
  timestamp: number
  reason: string
}

interface Props {
  open: boolean
  onClose: () => void
}

export function HistoryPanel({ open, onClose }: Props) {
  const [snapshots, setSnapshots] = useState<Snapshot[]>([])
  const [loading, setLoading] = useState(false)
  const [rollingBack, setRollingBack] = useState<string | null>(null)
  const [deleting, setDeleting] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  // 打开时拉取快照列表
  useEffect(() => {
    if (!open) return
    refreshList()
  }, [open])

  // 监听主机响应
  useEffect(() => {
    if (!open) return
    const unsubscribe = onHostMessage((data) => {
      if (data.type === 'snapshots') {
        setSnapshots(data.payload?.list ?? [])
        setLoading(false)
      } else if (data.type === 'rollback_result') {
        setRollingBack(null)
        if (data.payload?.success) {
          setError(null)
          // 回滚成功后关闭面板（工作簿已切换）
          setTimeout(() => onClose(), 600)
        } else {
          setError('回滚失败，请检查日志')
        }
      } else if (data.type === 'delete_snapshot_result') {
        setDeleting(null)
        if (data.payload?.success) {
          // 刷新列表
          refreshList()
        } else {
          setError('删除失败')
        }
      } else if (data.type === 'error') {
        setError(data.payload?.message ?? '未知错误')
        setLoading(false)
        setRollingBack(null)
        setDeleting(null)
      }
    })
    return unsubscribe
  }, [open])

  const refreshList = () => {
    setLoading(true)
    setError(null)
    sendToHost({ type: 'list_snapshots', payload: {} })
  }

  const handleRollback = (snapshot: Snapshot) => {
    if (!confirm(`确定要回滚到该快照吗？\n\n时间：${snapshot.createdAt}\n工作簿：${snapshot.workbookName}\n\n当前未保存的修改将丢失。`)) {
      return
    }
    setRollingBack(snapshot.id)
    setError(null)
    sendToHost({ type: 'rollback_snapshot', payload: { snapshot_id: snapshot.id } })
  }

  const handleDelete = (snapshot: Snapshot) => {
    if (!confirm(`确定要删除该快照吗？\n\n时间：${snapshot.createdAt}\n工作簿：${snapshot.workbookName}\n\n此操作不可撤销。`)) {
      return
    }
    setDeleting(snapshot.id)
    setError(null)
    sendToHost({ type: 'delete_snapshot', payload: { snapshot_id: snapshot.id } })
  }

  if (!open) return null

  return (
    <div className="history-panel-overlay">
      <div className="history-panel">
        <div className="history-header">
          <h3>历史版本</h3>
          <div className="history-actions">
            <button
              className="history-refresh-btn"
              onClick={refreshList}
              disabled={loading}
              title="刷新"
            >
              {loading ? '⟳' : '↻'}
            </button>
            <button
              className="history-close-btn"
              onClick={onClose}
              title="关闭"
            >
              ✕
            </button>
          </div>
        </div>

        {error && (
          <div className="history-error">{error}</div>
        )}

        <div className="history-list">
          {loading && snapshots.length === 0 && (
            <div className="history-empty">加载中...</div>
          )}
          {!loading && snapshots.length === 0 && (
            <div className="history-empty">
              暂无历史版本
              <div className="history-empty-hint">
                AI 执行操作前会自动创建快照
              </div>
            </div>
          )}
          {snapshots.map((s) => (
            <div key={s.id} className="history-item">
              <div className="history-item-info">
                <div className="history-item-title">
                  <span className="history-wb-icon" aria-hidden="true">📊</span>
                  <span className="history-wb-name">{s.workbookName}</span>
                </div>
                <div className="history-item-meta">
                  <span className="history-time">{s.createdAt}</span>
                  {s.reason && s.reason !== 'auto' && (
                    <span className="history-reason">{s.reason}</span>
                  )}
                </div>
              </div>
              <div className="history-item-actions">
                <button
                  className="history-rollback-btn"
                  onClick={() => handleRollback(s)}
                  disabled={rollingBack !== null || deleting !== null}
                >
                  {rollingBack === s.id ? '回滚中...' : '回滚'}
                </button>
                <button
                  className="history-delete-btn"
                  onClick={() => handleDelete(s)}
                  disabled={rollingBack !== null || deleting !== null}
                  title="删除"
                >
                  {deleting === s.id ? '...' : '🗑'}
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}
