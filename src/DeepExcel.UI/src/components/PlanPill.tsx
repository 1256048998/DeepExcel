import { useState } from 'react'
import type { PlanItem } from '../types'

// 计划胶囊：模型用 todo_write 维护的任务清单，显示在输入框上方。
// 收起时一行：「计划 2/5 · 正在：写合计公式」；点开看全部步骤。
export function PlanPill({ items }: { items: PlanItem[] }) {
  const [open, setOpen] = useState(false)
  if (items.length === 0) return null
  const done = items.filter(i => i.status === 'completed').length
  const current = items.find(i => i.status === 'in_progress')
  const allDone = done === items.length
  return (
    <div className={`plan-pill${open ? ' open' : ''}${allDone ? ' done' : ''}`}>
      <button className="plan-pill-header" type="button" onClick={() => setOpen(!open)} aria-expanded={open}>
        <span className="plan-pill-count">计划 {done}/{items.length}</span>
        <span className="plan-pill-current">
          {allDone ? '全部完成' : current ? `正在：${current.content}` : `下一步：${items.find(i => i.status === 'pending')?.content ?? ''}`}
        </span>
        <span className="chevron">{open ? '▾' : '▸'}</span>
      </button>
      {open && (
        <ol className="plan-pill-list">
          {items.map((item, i) => (
            <li key={i} className={`plan-item ${item.status}`}>
              <span className="plan-item-mark" aria-hidden="true">
                {item.status === 'completed' ? '☑' : item.status === 'in_progress' ? '◐' : '☐'}
              </span>
              <span>{item.content}</span>
            </li>
          ))}
        </ol>
      )}
    </div>
  )
}
