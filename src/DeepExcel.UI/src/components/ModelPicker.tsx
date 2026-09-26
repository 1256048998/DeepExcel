import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent } from 'react'
import { providerIcons } from '../providerIcons'
import type { ModelOption } from './InputArea'

interface Props {
  options: ModelOption[]
  /** `${provider}::${model}` */
  value?: string
  onChange: (provider: string, model: string) => void
  disabled?: boolean
  /** 弹层底部「管理模型」：打开模型配置 */
  onManage?: () => void
}

const keyOf = (o: ModelOption) => `${o.provider}::${o.model}`

/**
 * 输入框里的模型选择：胶囊按钮 + 向上弹出的分组列表（替掉原生 select）。
 * 键盘：↑↓ 移动、Enter 选中、Esc 关闭；点外面关闭。
 * 只展示已连接供应商的模型；成本、速度这类信息要等服务端目录下发，客户端不自己编。
 */
export function ModelPicker({ options, value, onChange, disabled, onManage }: Props) {
  const [open, setOpen] = useState(false)
  const [active, setActive] = useState(0)
  const rootRef = useRef<HTMLDivElement>(null)
  const listRef = useRef<HTMLDivElement>(null)

  const current = options.find(o => keyOf(o) === value) ?? options[0]
  const groups = useMemo(() => {
    const map = new Map<string, ModelOption[]>()
    for (const o of options) {
      if (!map.has(o.provider)) map.set(o.provider, [])
      map.get(o.provider)!.push(o)
    }
    return [...map.values()]
  }, [options])
  const flat = useMemo(() => groups.flat(), [groups])

  useEffect(() => {
    if (!open) return
    setActive(Math.max(0, flat.findIndex(o => keyOf(o) === value)))
    const onPointer = (e: MouseEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown', onPointer)
    return () => document.removeEventListener('mousedown', onPointer)
  }, [open, flat, value])

  useEffect(() => {
    if (!open) return
    listRef.current?.querySelector<HTMLElement>(`[data-index="${active}"]`)?.scrollIntoView({ block: 'nearest' })
  }, [open, active])

  const choose = (o: ModelOption) => {
    setOpen(false)
    if (keyOf(o) !== value) onChange(o.provider, o.model)
  }

  const onKeyDown = (e: KeyboardEvent) => {
    if (!open) {
      if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); setOpen(true) }
      return
    }
    if (e.key === 'Escape') { e.preventDefault(); setOpen(false) }
    else if (e.key === 'ArrowDown') { e.preventDefault(); setActive(i => (i + 1) % flat.length) }
    else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(i => (i - 1 + flat.length) % flat.length) }
    else if (e.key === 'Enter') { e.preventDefault(); if (flat[active]) choose(flat[active]) }
  }

  if (!current) return null
  const CurrentIcon = providerIcons[current.provider] || providerIcons.custom
  let index = -1

  return (
    <div className="model-picker" ref={rootRef} onKeyDown={onKeyDown}>
      <button
        type="button"
        className={`model-picker-btn${open ? ' open' : ''}`}
        onClick={() => setOpen(v => !v)}
        disabled={disabled}
        title={`${current.providerDisplayName} · ${current.model}（对话输出结束后切换）`}
        aria-haspopup="listbox"
        aria-expanded={open}
      >
        <CurrentIcon size={16} />
        <span className="model-picker-name">{current.model}</span>
        <svg className="model-picker-caret" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor"
          strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <polyline points="6 15 12 9 18 15" />
        </svg>
      </button>
      {open && (
        <div className="model-picker-pop" role="listbox" aria-label="选择模型" ref={listRef}>
          {groups.map(group => {
            const Icon = providerIcons[group[0].provider] || providerIcons.custom
            return (
              <div key={group[0].provider} className="model-picker-group">
                <div className="model-picker-group-title">
                  <Icon size={14} />
                  {group[0].providerDisplayName}
                </div>
                {group.map(o => {
                  index += 1
                  const i = index
                  const selected = keyOf(o) === value
                  return (
                    <button
                      key={keyOf(o)}
                      type="button"
                      role="option"
                      aria-selected={selected}
                      data-index={i}
                      className={`model-picker-item${i === active ? ' active' : ''}${selected ? ' selected' : ''}`}
                      onMouseEnter={() => setActive(i)}
                      onClick={() => choose(o)}
                    >
                      <span className="model-picker-item-name">{o.model}</span>
                      {o.isPrimary && <span className="model-picker-badge">主模型</span>}
                      {selected && (
                        <svg className="model-picker-check" width="13" height="13" viewBox="0 0 24 24" fill="none"
                          stroke="currentColor" strokeWidth="2.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                          <polyline points="20 6 9 17 4 12" />
                        </svg>
                      )}
                    </button>
                  )
                })}
              </div>
            )
          })}
          {onManage && (
            <button type="button" className="model-picker-manage" onClick={() => { setOpen(false); onManage() }}>
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
                strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <line x1="4" y1="21" x2="4" y2="14" /><line x1="4" y1="10" x2="4" y2="3" />
                <line x1="12" y1="21" x2="12" y2="12" /><line x1="12" y1="8" x2="12" y2="3" />
                <line x1="20" y1="21" x2="20" y2="16" /><line x1="20" y1="12" x2="20" y2="3" />
                <line x1="1" y1="14" x2="7" y2="14" /><line x1="9" y1="8" x2="15" y2="8" /><line x1="17" y1="16" x2="23" y2="16" />
              </svg>
              管理模型与密钥…
            </button>
          )}
        </div>
      )}
    </div>
  )
}
