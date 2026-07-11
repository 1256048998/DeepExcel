// src/DeepExcel.UI/src/components/PromptManager.tsx
// ★ 提示词 / 技能管理面板（模态框）：增删改
// - 提示词（prompt）：纯文本常用指令
// - 技能（skill）：遵循 Claude Agent Skills 标准（SKILL.md 格式：YAML frontmatter + Markdown 指令）
// 两者均为用户级（localStorage 持久化，跨工作簿保留）

import { useState, useEffect } from 'react'
import type { PromptTemplate, PromptType } from '../utils/prompts'
import { addPrompt, updatePrompt, deletePrompt, makeSkillTemplate } from '../utils/prompts'

interface Props {
  visible: boolean
  prompts: PromptTemplate[]
  onChange: (prompts: PromptTemplate[]) => void  // 更新父组件状态
  onClose: () => void
  // ★ 从历史消息保存时传入预填内容
  prefillContent?: string
  // ★ 初始类型（默认 prompt；从右上角"新建技能"入口可传 skill）
  initialType?: PromptType
}

export function PromptManager({ visible, prompts, onChange, onClose, prefillContent, initialType = 'prompt' }: Props) {
  // 编辑状态：null=新增，string=编辑指定 id
  const [editingId, setEditingId] = useState<string | null>(null)
  const [title, setTitle] = useState('')
  const [content, setContent] = useState('')
  // ★ 类型：prompt=普通提示词 / skill=Claude 标准技能
  const [type, setType] = useState<PromptType>('prompt')

  // ★ 预填内容（从历史消息保存）：用 initialType 初始化类型
  useEffect(() => {
    if (visible && prefillContent) {
      setEditingId(null)
      setContent(prefillContent)
      setTitle('')
      setType(initialType)
    }
  }, [visible, prefillContent, initialType])

  // ★ 打开面板时重置编辑状态（非预填模式）
  useEffect(() => {
    if (visible && !prefillContent) {
      setEditingId(null)
      setTitle('')
      setContent('')
      setType(initialType)
    }
  }, [visible])

  if (!visible) return null

  const handleSave = () => {
    if (!title.trim() && !content.trim()) return
    if (editingId) {
      onChange(updatePrompt(prompts, editingId, { title, content, type }))
    } else {
      onChange(addPrompt(prompts, title, content, type, prefillContent ? 'history' : 'manual'))
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
    setType(p.type)
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

  // ★ 切换类型：切到 skill 且 content 为空时，预填 SKILL.md 模板
  const handleTypeChange = (newType: PromptType) => {
    if (newType === 'skill' && !content.trim()) {
      setContent(makeSkillTemplate(title))
    }
    setType(newType)
  }

  // ★ title 变化时，如果是 skill 且 content 还是模板（未编辑），同步更新模板里的 name
  const handleTitleChange = (newTitle: string) => {
    setTitle(newTitle)
    if (type === 'skill' && content.trim() && content.startsWith('---')) {
      // content 仍是未修改的模板（包含占位提示文字）→ 同步 name
      if (content.includes('在这里描述') || content.includes('在这里编写')) {
        setContent(makeSkillTemplate(newTitle))
      }
    }
  }

  const isSkill = type === 'skill'
  const placeholderTitle = isSkill ? '技能名称（如：data-cleaning）' : '提示词标题（如：清洗数据）'
  const placeholderContent = isSkill
    ? 'SKILL.md 格式：YAML frontmatter + Markdown 指令'
    : '提示词内容（如：清洗A列数据，去除空格和特殊字符）'

  return (
    <div className="conv-panel-overlay" onClick={onClose}>
      <div className="conv-panel prompt-manager" onClick={e => e.stopPropagation()}>
        <div className="conv-header">
          <h3>提示词与技能管理</h3>
          <button className="conv-close-btn" onClick={onClose} title="关闭">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>

        {/* 编辑区 */}
        <div className="prompt-editor">
          {/* ★ 类型切换：提示词 / 技能 */}
          <div className="prompt-type-switch">
            <button
              type="button"
              className={`prompt-type-tab ${type === 'prompt' ? 'active' : ''}`}
              onClick={() => handleTypeChange('prompt')}
            >
              提示词
            </button>
            <button
              type="button"
              className={`prompt-type-tab ${type === 'skill' ? 'active' : ''}`}
              onClick={() => handleTypeChange('skill')}
            >
              技能 Skill
            </button>
            {isSkill && (
              <span className="prompt-type-hint">遵循 Claude Agent Skills 标准（SKILL.md）</span>
            )}
          </div>
          <input
            type="text"
            className="prompt-editor-title"
            placeholder={placeholderTitle}
            value={title}
            onChange={e => handleTitleChange(e.target.value)}
            maxLength={50}
          />
          <textarea
            className="prompt-editor-content"
            placeholder={placeholderContent}
            value={content}
            onChange={e => setContent(e.target.value)}
            rows={isSkill ? 8 : 3}
            maxLength={2000}
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
            <div className="prompt-list-empty">还没有提示词或技能，在上方创建或从对话中保存</div>
          )}
          {prompts.map(p => (
            <div key={p.id} className={`prompt-list-item ${editingId === p.id ? 'editing' : ''}`}>
              <div className="prompt-list-item-main" onClick={() => handleEdit(p)}>
                <div className="prompt-list-item-title">
                  <span className="prompt-list-slash">/</span>
                  {p.title}
                  <span className={`prompt-type-badge ${p.type === 'skill' ? 'prompt-type-skill' : 'prompt-type-prompt'}`}>
                    {p.type === 'skill' ? '技能' : '提示词'}
                  </span>
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
