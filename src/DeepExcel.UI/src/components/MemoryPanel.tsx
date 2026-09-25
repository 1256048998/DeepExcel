import { useEffect, useState } from 'react'
import { sendToHost, onHostMessage } from '../bridge'

interface MemoryState {
  available: boolean
  workbookName: string
  notes: string
  historyCount: number
  error?: string | null
}

interface Props {
  open: boolean
  onClose: () => void
}

// 新记忆的骨架：与侧车 workbook_memory.TEMPLATE 的四个小节一致
const TEMPLATE = `# 工作簿记忆

## 结构怪癖

## 用户偏好

## 做过的改动

## 禁区
`

/**
 * 工作簿记忆（CLAUDE.md 的工作簿版）：AI 在这个工作簿上记下的结构怪癖、你的偏好、做过的改动和禁区。
 * 只存在这台电脑上。禁区里的表和区域 AI 写不进去；只有你能在这里解除。
 */
export function MemoryPanel({ open, onClose }: Props) {
  const [memory, setMemory] = useState<MemoryState | null>(null)
  const [draft, setDraft] = useState('')
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState(false)

  useEffect(() => {
    if (!open) return
    setMemory(null)
    setError(null)
    setSaved(false)
    sendToHost({ type: 'memory_get', payload: {} })
  }, [open])

  useEffect(() => {
    if (!open) return
    return onHostMessage((data) => {
      if (data.type === 'memory') {
        const next = data.payload as MemoryState
        setMemory(next)
        setDraft(next.notes || '')
        setSaving(false)
        setError(next.error || null)
      } else if (data.type === 'error' && saving) {
        setSaving(false)
        setError(data.payload?.message ?? '保存失败')
      }
    })
  }, [open, saving])

  if (!open) return null

  const dirty = memory !== null && draft !== (memory.notes || '')

  const save = () => {
    setSaving(true)
    setError(null)
    setSaved(true)
    sendToHost({ type: 'memory_save', payload: { notes: draft } })
  }

  const clear = () => {
    if (!confirm(`清除「${memory?.workbookName || '这个工作簿'}」的全部记忆？\n\n包括笔记、禁区和操作记录，此操作不可撤销。`)) return
    setSaving(true)
    setSaved(false)
    sendToHost({ type: 'memory_clear', payload: {} })
  }

  return (
    <div className="history-panel-overlay">
      <div className="history-panel memory-panel">
        <div className="history-header">
          <h3>工作簿记忆</h3>
          <div className="history-actions">
            <button className="history-close-btn" onClick={onClose} title="关闭" type="button">✕</button>
          </div>
        </div>

        {error && <div className="history-error">{error}</div>}

        {memory === null && <div className="history-empty">加载中...</div>}

        {memory && !memory.available && (
          <div className="history-empty">
            这个工作簿还没保存过
            <div className="history-empty-hint">保存到磁盘后，AI 才能为它记下结构、偏好和禁区</div>
          </div>
        )}

        {memory && memory.available && (
          <div className="memory-body">
            <div className="memory-hint">
              AI 在「{memory.workbookName}」上记下的东西，只存在这台电脑上，下次打开对话时自动带给 AI。
              「## 禁区」下每行一个表名或区域（如 <code>- 汇总</code>、<code>- 明细!A1:D20</code>），
              AI 写不进去；只有你能在这里删掉。
            </div>
            <textarea
              className="memory-editor"
              value={draft}
              placeholder={TEMPLATE}
              onChange={e => { setDraft(e.target.value); setSaved(false) }}
              spellCheck={false}
            />
            <div className="memory-footer">
              <span className="memory-meta">
                {memory.historyCount > 0 ? `操作记录 ${memory.historyCount} 条` : '还没有操作记录'}
                {saved && !dirty && !error ? ' · 已保存' : ''}
              </span>
              <div className="memory-buttons">
                {!draft && (
                  <button className="memory-btn" type="button" onClick={() => setDraft(TEMPLATE)}>用模板</button>
                )}
                <button
                  className="memory-btn danger"
                  type="button"
                  onClick={clear}
                  disabled={saving || (!memory.notes && memory.historyCount === 0)}
                >
                  清除
                </button>
                <button className="memory-btn primary" type="button" onClick={save} disabled={saving || !dirty}>
                  {saving ? '保存中...' : '保存'}
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
