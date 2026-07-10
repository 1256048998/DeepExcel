// src/DeepExcel.UI/src/components/PromptManager.tsx
// ★ 提示词管理面板（模态框）：增删改
// 从历史消息保存时，预填充 content

import { useState, useEffect } from 'react'
import type { PromptTemplate } from '../utils/prompts'
import { addPrompt, updatePrompt, deletePrompt } from '../utils/prompts'

interface Props {
  visible: boolean
  prompts: PromptTemplate[]
  onChange: (prompts: PromptTemplate[]) => void  // 更新父组件状态
  onClose: () => void
  // ★ 从历史消息保存时传入预填内容
  prefillContent?: string
}

export function PromptManager({ visible, prompts, onChange, onClose, prefillContent }: Props) {
  // 编辑状态：null=新增，string=编辑指定 id
  const [editingId, setEditingId] = useState<string | null>(null)
  const [title, setTitle] = useState('')
  const [content, setContent] = useState('')

  // ★ 预填内容（从历史消息保存）
  useEffect(() => {
    if (visible && prefillContent) {
      setEditingId(null)
      setContent(prefillContent)
      setTitle('')
    }
  }, [visible, prefillContent])

  // ★ 打开面板时重置编辑状态
  useEffect(() => {
    if (visible && !prefillContent) {
      setEditingId(null)
      setTitle('')
      setContent('')
    }
  }, [visible])

  if (!visible) return null

  const handleSave = () => {
    if (!title.trim() && !content.trim()) return
    if (editingId) {
      onChange(updatePrompt(prompts, editingId, { title, content }))
    } else {
      onChange(addPrompt(prompts, title, content, prefillContent ? 'history' : 'manual'))
    }
    // 重置
    setEditingId(null)
    setTitle('')
    setContent('')
  }

  const handleEdit = (p: PromptTemplate) => {
    setEditingId(p.id)
    setTitle(p.title)
    setContent(p.content)
  }

  const handleDelete = (id: string) => {
    onChange(deletePrompt(prompts, id))
    if (editingId === id) {
      setEditingId(null)
      setTitle('')
      setContent('')
    }
  }

  const handleCancel = () => {
    setEditingId(null)
    setTitle('')
    setContent('')
  }

  return (
    <div className="conv-panel-overlay" onClick={onClose}>
      <div className="conv-panel prompt-manager" onClick={e => e.stopPropagation()}>
        <div className="conv-header">
          <h3>提示词管理</h3>
          <button className="conv-close-btn" onClick={onClose} title="关闭">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>

        {/* 编辑区 */}
        <div className="prompt-editor">
          <input
            type="text"
            className="prompt-editor-title"
            placeholder="提示词标题（如：清洗数据）"
            value={title}
            onChange={e => setTitle(e.target.value)}
            maxLength={50}
          />
          <textarea
            className="prompt-editor-content"
            placeholder="提示词内容（如：清洗A列数据，去除空格和特殊字符）"
            value={content}
            onChange={e => setContent(e.target.value)}
            rows={3}
            maxLength={500}
          />
          <div className="prompt-editor-actions">
            {editingId && (
              <button className="prompt-btn cancel" onClick={handleCancel}>取消编辑</button>
            )}
            <button
              className="prompt-btn save"
              onClick={handleSave}
              disabled={!title.trim() && !content.trim()}
            >
              {editingId ? '更新' : '保存'}
            </button>
          </div>
        </div>

        {/* 列表区 */}
        <div className="prompt-list">
          {prompts.length === 0 && (
            <div className="prompt-list-empty">还没有提示词，在上方创建或从对话中保存</div>
          )}
          {prompts.map(p => (
            <div key={p.id} className={`prompt-list-item ${editingId === p.id ? 'editing' : ''}`}>
              <div className="prompt-list-item-main" onClick={() => handleEdit(p)}>
                <div className="prompt-list-item-title">
                  <span className="prompt-list-slash">/</span>
                  {p.title}
                  {p.source === 'history' && <span className="prompt-list-badge">历史</span>}
                </div>
                <div className="prompt-list-item-content">{p.content}</div>
              </div>
              <button
                className="prompt-list-item-delete"
                onClick={() => handleDelete(p.id)}
                title="删除"
              >
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="3 6 5 6 21 6" />
                  <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                </svg>
              </button>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}
