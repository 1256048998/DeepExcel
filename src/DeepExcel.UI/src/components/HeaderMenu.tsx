import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'

export type HeaderMenuItem = {
  key: string
  label: string
  icon: ReactNode
  onSelect: () => void
  /** 开关类菜单项：显示勾选状态 */
  checked?: boolean
  /** 右侧的小徽标（如附件数） */
  badge?: number
}

/**
 * 顶栏「更多」菜单。窄侧栏里顶栏只放新建、历史、账户和它，
 * 其余入口（历史版本、提示词、技能、模型配置、附件、加载历史）都收在这里。
 */
export function HeaderMenu({ items }: { items: HeaderMenuItem[] }) {
  const [open, setOpen] = useState(false)
  const rootRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!open) return
    const onPointer = (e: MouseEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(false)
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false)
    }
    document.addEventListener('mousedown', onPointer)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', onPointer)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])

  const badgeTotal = items.reduce((n, item) => n + (item.badge ?? 0), 0)

  return (
    <div className="header-menu" ref={rootRef}>
      <button
        className={`header-btn icon-only${open ? ' on' : ''}`}
        onClick={() => setOpen(v => !v)}
        title="更多"
        aria-label="更多"
        aria-haspopup="menu"
        aria-expanded={open}
        type="button"
      >
        <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
          <circle cx="5" cy="12" r="1.8" />
          <circle cx="12" cy="12" r="1.8" />
          <circle cx="19" cy="12" r="1.8" />
        </svg>
        {badgeTotal > 0 && <span className="header-menu-dot" aria-hidden="true" />}
      </button>
      {open && (
        <div className="header-menu-list" role="menu">
          {items.map(item => (
            <button
              key={item.key}
              className="header-menu-item"
              role={item.checked === undefined ? 'menuitem' : 'menuitemcheckbox'}
              aria-checked={item.checked}
              type="button"
              onClick={() => {
                setOpen(false)
                item.onSelect()
              }}
            >
              <span className="header-menu-icon" aria-hidden="true">{item.icon}</span>
              <span className="header-menu-label">{item.label}</span>
              {item.badge ? <span className="header-menu-badge">{item.badge}</span> : null}
              {item.checked !== undefined && (
                <span className={`header-menu-check${item.checked ? ' on' : ''}`} aria-hidden="true">
                  {item.checked ? '开' : '关'}
                </span>
              )}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
