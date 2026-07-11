// src/DeepExcel.UI/src/components/PromptDropdown.tsx
// ★ / 触发的提示词下拉列表（Claude Code 风格斜杠命令）
// 交互：↑↓ 选择，Enter 确认（填充到输入框），ESC 关闭，输入继续过滤

import { useState, useEffect, useRef } from 'react'
import type { PromptTemplate } from '../utils/prompts'
import { matchPrompts } from '../utils/prompts'

interface Props {
  query: string           // / 后面的查询字符串
  prompts: PromptTemplate[]
  onSelect: (prompt: PromptTemplate) => void   // 选中提示词
  onCreateNew: () => void  // 点击"+ 新建提示词"
  onClose: () => void      // ESC 关闭
}

export function PromptDropdown({ query, prompts, onSelect, onCreateNew, onClose }: Props) {
  const matched = matchPrompts(prompts, query)
  // ★ 选中索引：-1 表示选中"+ 新建"，0~n 表示选中对应提示词
  const [selectedIndex, setSelectedIndex] = useState(0)
  const listRef = useRef<HTMLDivElement>(null)

  // ★ 查询变化时重置选中索引
  useEffect(() => {
    setSelectedIndex(matched.length > 0 ? 0 : -1)
  }, [query])

  // ★ 键盘事件：在 InputArea 的 onKeyDown 中调用 handleKeyDown
  // 这里通过全局监听处理（简化实现，避免 prop 透传复杂度）
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'ArrowDown') {
        e.preventDefault()
        setSelectedIndex(idx => {
          const max = matched.length - 1
          // -1 是"新建"，0~n 是提示词；循环导航
          if (idx === -1) return 0
          return idx >= max ? -1 : idx + 1
        })
      } else if (e.key === 'ArrowUp') {
        e.preventDefault()
        setSelectedIndex(idx => {
          if (idx === -1) return matched.length - 1
          return idx <= 0 ? -1 : idx - 1
        })
      } else if (e.key === 'Enter') {
        e.preventDefault()
        e.stopPropagation()
        if (selectedIndex === -1) {
          onCreateNew()
        } else if (matched[selectedIndex]) {
          onSelect(matched[selectedIndex])
        }
      } else if (e.key === 'Escape') {
        e.preventDefault()
        onClose()
      }
    }
    // ★ 捕获阶段监听，确保在 InputArea 的 onKeyDown 之前处理
    document.addEventListener('keydown', onKey, true)
    return () => document.removeEventListener('keydown', onKey, true)
  }, [matched, selectedIndex, onSelect, onCreateNew, onClose])

  // ★ 滚动到选中项
  useEffect(() => {
    if (!listRef.current) return
    const el = listRef.current.querySelector(`[data-idx="${selectedIndex}"]`) as HTMLElement
    if (el) el.scrollIntoView({ block: 'nearest' })
  }, [selectedIndex])

  return (
    <div className="prompt-dropdown" ref={listRef}>
      <div className="prompt-dropdown-header">
        <span className="prompt-dropdown-title">提示词/技能</span>
        <span className="prompt-dropdown-count">{matched.length}</span>
      </div>
      <div className="prompt-dropdown-list">
        {matched.length === 0 && query.trim() && (
          <div className="prompt-dropdown-empty">无匹配提示词</div>
        )}
        {matched.map((p, i) => (
          <div
            key={p.id}
            data-idx={i}
            className={`prompt-dropdown-item ${selectedIndex === i ? 'selected' : ''}`}
            onClick={() => onSelect(p)}
            onMouseEnter={() => setSelectedIndex(i)}
          >
            <div className="prompt-dropdown-item-title">
              <span className="prompt-dropdown-slash">/</span>
              {p.title}
              <span className={`prompt-type-badge ${p.type === 'skill' ? 'prompt-type-skill' : 'prompt-type-prompt'}`}>
                {p.type === 'skill' ? '技能' : '提示词'}
              </span>
            </div>
            <div className="prompt-dropdown-item-content">{p.content}</div>
          </div>
        ))}
        {/* ★ "+ 新建提示词" 固定在底部 */}
        <div
          data-idx={-1}
          className={`prompt-dropdown-item create-new ${selectedIndex === -1 ? 'selected' : ''}`}
          onClick={onCreateNew}
          onMouseEnter={() => setSelectedIndex(-1)}
        >
          <div className="prompt-dropdown-item-title">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="12" y1="5" x2="12" y2="19" />
              <line x1="5" y1="12" x2="19" y2="12" />
            </svg>
            新建提示词/技能
          </div>
        </div>
      </div>
    </div>
  )
}
