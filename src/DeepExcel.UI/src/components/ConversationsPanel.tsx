import { useState, useEffect } from 'react'
import { sendToHost, sendToHostWithResponse } from '../bridge'
import type { ConversationSummary, Message } from '../types'

interface Props {
  open: boolean
  onClose: () => void
  onContinue: (messages: Message[]) => void
}

export function ConversationsPanel({ open, onClose, onContinue }: Props) {
  const [list, setList] = useState<ConversationSummary[]>([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [confirmingId, setConfirmingId] = useState<string | null>(null)

  // 打开时加载列表
  useEffect(() => {
    if (open) {
      loadList()
    }
  }, [open])

  const loadList = async () => {
    setLoading(true)
    setError('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'list_conversations', payload: {} },
        'conversations'
      )
      if (resp?.type === 'conversations' && Array.isArray(resp.payload?.list)) {
        setList(resp.payload.list)
      } else if (resp?.type === 'error') {
        setError(resp.payload?.message || '加载失败')
        setList([])
      } else {
        setList([])
      }
    } catch (e) {
      setError(`加载失败: ${e}`)
      setList([])
    } finally {
      setLoading(false)
    }
  }

  const handleContinue = async (id: string) => {
    setLoading(true)
    setError('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'continue_conversation', payload: { conversation_id: id } },
        'continue_conversation'
      )
      if (resp?.type === 'continue_conversation' && Array.isArray(resp.payload?.messages)) {
        // 转换 HistoryMessage → Message
        const msgs: Message[] = resp.payload.messages.map((m: any) => ({
          role: m.role,
          content: m.content || '',
          type: m.type,
          options: m.options,
          toolGroup: m.toolGroup,
          streaming: false
        }))
        onContinue(msgs)
        onClose()
      } else if (resp?.type === 'error') {
        setError(resp.payload?.message || '继续对话失败')
      } else {
        setError('继续对话失败：未收到响应')
      }
    } catch (e) {
      setError(`继续对话失败: ${e}`)
    } finally {
      setLoading(false)
      setConfirmingId(null)
    }
  }

  const handleDelete = async (id: string) => {
    try {
      await sendToHostWithResponse(
        { type: 'delete_conversation', payload: { conversation_id: id } },
        'delete_conversation'
      )
      await loadList()
    } catch (e) {
      console.warn('delete failed', e)
    }
  }

  if (!open) return null

  const formatDate = (iso: string) => {
    try {
      const d = new Date(iso)
      const now = new Date()
      const isToday = d.toDateString() === now.toDateString()
      const time = d.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' })
      if (isToday) return `今天 ${time}`
      return d.toLocaleDateString('zh-CN', { month: '2-digit', day: '2-digit' }) + ' ' + time
    } catch {
      return ''
    }
  }

  return (
    <div className="conv-panel-overlay" onClick={onClose}>
      <div className="conv-panel" onClick={e => e.stopPropagation()}>
        <div className="conv-header">
          <h3>历史对话</h3>
          <div className="conv-actions">
            <button
              className="conv-refresh-btn"
              onClick={loadList}
              disabled={loading}
              title="刷新"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <polyline points="23 4 23 10 17 10"></polyline>
                <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"></path>
              </svg>
            </button>
            <button className="conv-close-btn" onClick={onClose} title="关闭">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="18" y1="6" x2="6" y2="18"></line>
                <line x1="6" y1="6" x2="18" y2="18"></line>
              </svg>
            </button>
          </div>
        </div>

        {error && <div className="conv-error">{error}</div>}

        <div className="conv-list">
          {loading && list.length === 0 ? (
            <div className="conv-empty">加载中...</div>
          ) : list.length === 0 ? (
            <div className="conv-empty">
              <div>暂无历史对话</div>
              <div className="conv-empty-hint">开始新对话后会自动保存到历史</div>
            </div>
          ) : (
            list.map(conv => (
              <div key={conv.id} className="conv-item">
                <div className="conv-item-info">
                  <div className="conv-item-title">{conv.title || '未命名对话'}</div>
                  <div className="conv-item-meta">
                    {conv.updatedAt ? formatDate(conv.updatedAt) : formatDate(conv.createdAt)}
                  </div>
                </div>
                <div className="conv-item-actions">
                  {confirmingId === conv.id ? (
                    <>
                      <button
                        className="conv-confirm-btn"
                        onClick={() => handleContinue(conv.id)}
                        disabled={loading}
                      >
                        确认继续
                      </button>
                      <button
                        className="conv-cancel-btn"
                        onClick={() => setConfirmingId(null)}
                      >
                        取消
                      </button>
                    </>
                  ) : (
                    <>
                      <button
                        className="conv-continue-btn"
                        onClick={() => setConfirmingId(conv.id)}
                        title="继续该对话"
                      >
                        继续
                      </button>
                      <button
                        className="conv-delete-btn"
                        onClick={() => handleDelete(conv.id)}
                        title="删除"
                      >
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <polyline points="3 6 5 6 21 6"></polyline>
                          <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                        </svg>
                      </button>
                    </>
                  )}
                </div>
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  )
}
