import { useEffect, useState } from 'react'
import { sendToHostWithResponse } from '../bridge'
import type { ReactNode } from 'react'
import {
  NEVER_COLLECTED, PRIVACY_SUMMARY, ROUTE_TEXT, SENT_TO_MODEL, SKILL_SYNC_TEXT, STORED_LOCALLY,
  TELEMETRY_EVENTS, TELEMETRY_INTRO, TELEMETRY_SWITCH,
} from '../utils/privacy'

type UsageStats = { supported: boolean; enabled: boolean }

function UsageStatsSwitch() {
  const [stats, setStats] = useState<UsageStats | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    let alive = true
    sendToHostWithResponse({ type: 'get_usage_stats', payload: {} }, 'usage_stats')
      .then(resp => { if (alive && resp?.payload) setStats(resp.payload as UsageStats) })
      .catch(() => { if (alive) setError('读不到这个设置') })
    return () => { alive = false }
  }, [])

  if (error && !stats) return <p className="privacy-text muted">{error}</p>
  if (!stats) return null
  if (!stats.supported) return <p className="privacy-text">{TELEMETRY_SWITCH.unsupported}</p>

  const toggle = async () => {
    setBusy(true)
    setError(null)
    try {
      const resp = await sendToHostWithResponse(
        { type: 'set_usage_stats', payload: { enabled: !stats.enabled } }, 'usage_stats')
      if (resp?.payload) setStats(resp.payload as UsageStats)
      else setError('没能保存这个设置，请稍后重试')
    } catch {
      setError('没能保存这个设置，请稍后重试')
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="privacy-switch-row">
      <div className="privacy-switch-text">
        <span className="privacy-switch-label">{TELEMETRY_SWITCH.label}</span>
        <span className="privacy-switch-hint">{error ?? (stats.enabled ? TELEMETRY_SWITCH.on : TELEMETRY_SWITCH.off)}</span>
      </div>
      <button type="button" role="switch" aria-checked={stats.enabled} aria-label={TELEMETRY_SWITCH.label}
        className={`privacy-switch${stats.enabled ? ' on' : ''}`} disabled={busy} onClick={() => void toggle()}>
        <span className="privacy-switch-knob" />
      </button>
    </div>
  )
}

interface Props {
  open: boolean
  onClose: () => void
  /** 账号状态里的出口（服务端给的）；这里只展示，不参与路由判断 */
  mode: string | null
  serverUrl: string | null
}

function Section({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section className="privacy-section">
      <h4 className="privacy-section-title">{title}</h4>
      {children}
    </section>
  )
}

export function PrivacyPanel({ open, onClose, mode, serverUrl }: Props) {
  const [fieldsOpen, setFieldsOpen] = useState(false)
  if (!open) return null
  const hosted = mode === 'hosted'
  const schemaUrl = serverUrl ? `${serverUrl.replace(/\/+$/, '')}/api/v1/telemetry/schema` : null

  return (
    <div className="history-panel-overlay" onClick={onClose}>
      <div className="history-panel privacy-panel" onClick={e => e.stopPropagation()} role="dialog" aria-label="数据与隐私">
        <div className="history-header">
          <h3>数据与隐私</h3>
          <div className="history-actions">
            <button className="history-close-btn" onClick={onClose} title="关闭" type="button">✕</button>
          </div>
        </div>

        <div className="privacy-body">
          <div className="privacy-summary">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
              strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
              <polyline points="9 12 11 14 15 10" />
            </svg>
            <span>{PRIVACY_SUMMARY}</span>
          </div>

          <Section title="发给模型的内容">
            <ul className="privacy-list">
              {SENT_TO_MODEL.map(t => <li key={t}>{t}</li>)}
            </ul>
            <div className="privacy-routes">
              <div className={`privacy-route${!hosted ? ' current' : ''}`}>
                {!hosted && <span className="privacy-route-tag">当前</span>}
                {ROUTE_TEXT.byok}
              </div>
              <div className={`privacy-route${hosted ? ' current' : ''}`}>
                {hosted && <span className="privacy-route-tag">当前</span>}
                {ROUTE_TEXT.hosted}
              </div>
            </div>
          </Section>

          <Section title="只存在这台电脑上">
            <ul className="privacy-list">
              {STORED_LOCALLY.map(t => <li key={t}>{t}</li>)}
            </ul>
          </Section>

          <Section title="使用统计">
            <p className="privacy-text">{TELEMETRY_INTRO}</p>
            <UsageStatsSwitch />
            <div className="privacy-never">
              <span className="privacy-never-label">从不收集</span>
              {NEVER_COLLECTED.map(t => <span key={t} className="privacy-chip">{t}</span>)}
            </div>
            <button type="button" className="privacy-toggle" aria-expanded={fieldsOpen}
              onClick={() => setFieldsOpen(v => !v)}>
              {fieldsOpen ? '收起' : '查看'}会发送的全部字段
            </button>
            {fieldsOpen && (
              <dl className="privacy-fields">
                {Object.entries(TELEMETRY_EVENTS).map(([key, ev]) => (
                  <div key={key} className="privacy-field-row">
                    <dt>{ev.label}</dt>
                    <dd>{Object.values(ev.fields).join('、')}</dd>
                  </div>
                ))}
              </dl>
            )}
            {schemaUrl && (
              <p className="privacy-text muted">服务器公开的白名单：<code>{schemaUrl}</code></p>
            )}
          </Section>

          <Section title="技能同步">
            <p className="privacy-text">{SKILL_SYNC_TEXT}</p>
          </Section>
        </div>
      </div>
    </div>
  )
}
