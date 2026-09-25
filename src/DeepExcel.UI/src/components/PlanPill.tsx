import { useState } from 'react'
import type { PlanItem } from '../types'

// 计划胶囊：模型用 todo_write 维护的任务清单，显示在输入框上方。
// 收起时一行：环形进度 +「正在：写合计公式」+ 2/5；点开看全部步骤。
export function PlanPill({ items }: { items: PlanItem[] }) {
  const [open, setOpen] = useState(false)
  if (items.length === 0) return null
  const done = items.filter(i => i.status === 'completed').length
  const current = items.find(i => i.status === 'in_progress')
  const allDone = done === items.length
  return (
    <div className={`plan-pill${open ? ' open' : ''}${allDone ? ' done' : ''}`}>
      <button className="plan-pill-header" type="button" onClick={() => setOpen(!open)} aria-expanded={open}>
        <ProgressRing done={done} total={items.length} />
        <span className="plan-pill-current">
          {allDone ? '全部完成' : current ? `正在：${current.content}` : `下一步：${items.find(i => i.status === 'pending')?.content ?? ''}`}
        </span>
        <span className="plan-pill-count">{done}/{items.length}</span>
        <svg className={`chevron-icon${open ? ' open-up' : ' up'}`} width="10" height="10" viewBox="0 0 24 24" fill="none"
          stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <polyline points="9 6 15 12 9 18" />
        </svg>
      </button>
      {open && (
        <ol className="plan-pill-list">
          {items.map((item, i) => (
            <li key={i} className={`plan-item ${item.status}`}>
              <span className="plan-item-mark" aria-hidden="true">
                {item.status === 'completed' && (
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3.2" strokeLinecap="round" strokeLinejoin="round">
                    <polyline points="20 6 9 17 4 12" />
                  </svg>
                )}
              </span>
              <span>{item.content}</span>
            </li>
          ))}
        </ol>
      )}
    </div>
  )
}

// 20px 环形进度：已完成步数 / 总步数
function ProgressRing({ done, total }: { done: number; total: number }) {
  const r = 7.5
  const c = 2 * Math.PI * r
  const ratio = total > 0 ? done / total : 0
  return (
    <svg className="plan-ring" width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
      <circle cx="10" cy="10" r={r} fill="none" strokeWidth="2.5" className="plan-ring-track" />
      <circle cx="10" cy="10" r={r} fill="none" strokeWidth="2.5" className="plan-ring-value"
        strokeDasharray={c} strokeDashoffset={c * (1 - ratio)} strokeLinecap="round" transform="rotate(-90 10 10)" />
    </svg>
  )
}
