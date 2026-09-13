import { useState, useEffect, useCallback } from 'react'
import { sendToHostWithResponse } from '../bridge'

interface Props {
  open: boolean
  onClose: () => void
  /** 登录状态变化时通知外层，用于刷新状态栏 */
  onStatusChange?: (status: AccountStatus) => void
}

export interface Entitlement {
  plan: string
  status: string
  task_limit: number | null
  tasks_used: number
  tasks_remaining: number | null
}

export interface AccountStatus {
  state: 'signedout' | 'signedin' | 'offline' | 'expired'
  server_url: string | null
  email: string | null
  /** 'byok' | 'hosted'，由服务端决定，客户端只做展示 */
  mode: string | null
  entitlement: Entitlement | null
}

const PLAN_LABELS: Record<string, string> = {
  beta: '内测',
  free: '免费版',
  pro: '专业版',
  team: '团队版',
  byok: '自带密钥',
}

const STATE_LABELS: Record<AccountStatus['state'], string> = {
  signedout: '未登录',
  signedin: '已登录',
  offline: '离线（使用缓存的登录状态）',
  expired: '登录已失效',
}

/** 密码下限与服务端 RegisterRequest 保持一致 */
const MIN_PASSWORD_LENGTH = 10

export function AccountPanel({ open, onClose, onStatusChange }: Props) {
  const [status, setStatus] = useState<AccountStatus | null>(null)
  const [mode, setMode] = useState<'signin' | 'register'>('signin')
  const [serverUrl, setServerUrl] = useState('')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [inviteCode, setInviteCode] = useState('')
  const [inviteRequired, setInviteRequired] = useState<boolean | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  // 托管模式下的实际用量，来自代理——它是权威计数点
  const [usage, setUsage] = useState<{
    applicable: boolean
    period_days?: number
    calls?: number
    input_tokens?: number
    output_tokens?: number
  } | null>(null)

  const applyStatus = useCallback(
    (next: AccountStatus) => {
      setStatus(next)
      onStatusChange?.(next)
    },
    [onStatusChange],
  )

  const loadStatus = useCallback(async () => {
    try {
      const resp = await sendToHostWithResponse({ type: 'account_status', payload: {} }, 'account_status')
      const next = resp?.payload as AccountStatus
      if (next) {
        applyStatus(next)
        if (next.server_url) setServerUrl(next.server_url)
        if (next.mode === 'hosted') {
          try {
            const u = await sendToHostWithResponse(
              { type: 'account_usage', payload: {} },
              'account_usage',
            )
            setUsage(u?.payload ?? null)
          } catch {
            // 用量拿不到不该影响账号面板本身
            setUsage(null)
          }
        } else {
          setUsage(null)
        }
      }
    } catch {
      // 状态读取失败不该阻塞面板，保持未登录展示即可
    }
  }, [applyStatus])

  useEffect(() => {
    if (open) {
      setError('')
      void loadStatus()
    }
  }, [open, loadStatus])

  /**
   * 是否需要邀请码由服务端决定，不在客户端硬编码。
   * 这样改注册策略不需要重新发版。
   */
  const probeServer = useCallback(async (url: string) => {
    if (!url.trim()) {
      setInviteRequired(null)
      return
    }
    try {
      const resp = await sendToHostWithResponse(
        { type: 'account_server_meta', payload: { server_url: url.trim() } },
        'account_server_meta',
      )
      setInviteRequired(Boolean(resp?.payload?.invite_required))
      setError('')
    } catch {
      setInviteRequired(null)
    }
  }, [])

  const submit = async () => {
    setError('')

    if (!serverUrl.trim()) return setError('请填写服务器地址')
    if (!email.trim()) return setError('请填写邮箱')
    if (mode === 'register' && password.length < MIN_PASSWORD_LENGTH) {
      return setError(`密码至少 ${MIN_PASSWORD_LENGTH} 位`)
    }
    if (!password) return setError('请填写密码')

    setBusy(true)
    try {
      const resp = await sendToHostWithResponse(
        {
          type: mode === 'signin' ? 'account_sign_in' : 'account_register',
          payload: {
            server_url: serverUrl.trim(),
            email: email.trim(),
            password,
            invite_code: inviteCode.trim() || undefined,
          },
        },
        'account_status',
      )
      if (!resp) throw new Error('服务器无响应，请稍后重试')
      applyStatus(resp.payload as AccountStatus)
      // 不保留在内存里，登录态由 C# 侧以 DPAPI 加密持久化
      setPassword('')
      setInviteCode('')
    } catch (e: any) {
      setError(e?.message || (mode === 'signin' ? '登录失败' : '注册失败'))
    } finally {
      setBusy(false)
    }
  }

  const signOut = async () => {
    setBusy(true)
    setError('')
    try {
      const resp = await sendToHostWithResponse({ type: 'account_sign_out', payload: {} }, 'account_status')
      if (!resp) throw new Error('服务器无响应，请稍后重试')
      applyStatus(resp.payload as AccountStatus)
    } catch (e: any) {
      setError(e?.message || '退出登录失败')
    } finally {
      setBusy(false)
    }
  }

  if (!open) return null

  const signedIn = status?.state === 'signedin' || status?.state === 'offline'

  return (
    <div className="config-overlay" onClick={onClose}>
      <div className="config-panel" onClick={(e) => e.stopPropagation()}>
        <div className="config-header">
          <h3>账号</h3>
          <button className="config-close-btn" onClick={onClose} aria-label="关闭">
            ×
          </button>
        </div>

        <div className="config-body">
          {signedIn ? (
            <div className="config-section account-summary">
              <div className="account-row">
                <span className="account-label">状态</span>
                <span className={status.state === 'offline' ? 'account-warn' : 'account-ok'}>
                  {STATE_LABELS[status.state]}
                </span>
              </div>
              <div className="account-row">
                <span className="account-label">账号</span>
                <span>{status.email}</span>
              </div>
              <div className="account-row">
                <span className="account-label">服务器</span>
                <span className="account-mono">{status.server_url}</span>
              </div>
              {status.entitlement && (
                <>
                  <div className="account-row">
                    <span className="account-label">套餐</span>
                    <span>
                      {PLAN_LABELS[status.entitlement.plan] ?? status.entitlement.plan}
                    </span>
                  </div>
                  {status.entitlement.task_limit !== null && (
                    <div className="account-row">
                      <span className="account-label">本期额度</span>
                      <span>
                        已用 {status.entitlement.tasks_used} / {status.entitlement.task_limit}
                      </span>
                    </div>
                  )}
                </>
              )}
              {usage?.applicable && (
                <div className="account-row">
                  <span className="account-label">近 {usage.period_days} 天</span>
                  <span>
                    {usage.calls} 次调用 · {((usage.input_tokens ?? 0) + (usage.output_tokens ?? 0)).toLocaleString()} tokens
                  </span>
                </div>
              )}
              <div className="account-row">
                <span className="account-label">模型出口</span>
                <span>
                  {status.mode === 'hosted' ? 'DeepExcel 托管转发' : '本地配置的供应商（自带密钥）'}
                </span>
              </div>

              {status.state === 'offline' && (
                <p className="config-apikey-hint">
                  暂时无法连接服务器，正在使用缓存的登录状态。联网后会自动恢复。
                </p>
              )}

              <button className="config-toggle-btn" onClick={signOut} disabled={busy}>
                {busy ? '处理中…' : '退出登录'}
              </button>
            </div>
          ) : (
            <div className="config-section account-form">
              {status?.state === 'expired' && (
                <p className="config-apikey-hint account-warn">登录状态已失效，请重新登录。</p>
              )}

              <div className="account-tabs">
                <button
                  className={mode === 'signin' ? 'active' : ''}
                  onClick={() => {
                    setMode('signin')
                    setError('')
                  }}
                >
                  登录
                </button>
                <button
                  className={mode === 'register' ? 'active' : ''}
                  onClick={() => {
                    setMode('register')
                    setError('')
                  }}
                >
                  注册
                </button>
              </div>

              <label>
                服务器地址
                <input
                  className="config-input"
                  type="text"
                  value={serverUrl}
                  placeholder="https://api.deepexcel.com"
                  onChange={(e) => setServerUrl(e.target.value)}
                  onBlur={(e) => void probeServer(e.target.value)}
                  disabled={busy}
                />
              </label>

              <label>
                邮箱
                <input
                  className="config-input"
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  disabled={busy}
                />
              </label>

              <label>
                密码
                <input
                  className="config-input"
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && !busy) void submit()
                  }}
                  disabled={busy}
                />
              </label>

              {mode === 'register' && inviteRequired !== false && (
                <label>
                  邀请码{inviteRequired === null ? '（如服务器要求）' : ''}
                  <input
                    type="text"
                    value={inviteCode}
                    onChange={(e) => setInviteCode(e.target.value)}
                    disabled={busy}
                  />
                </label>
              )}

              {error && <p className="config-error">{error}</p>}

              <button className="config-save-btn" onClick={() => void submit()} disabled={busy}>
                {busy ? '处理中…' : mode === 'signin' ? '登录' : '注册'}
              </button>

              <p className="config-apikey-hint">
                不登录也可以使用：在模型设置里填写自己的 API Key 即可。登录用于内测准入与使用统计。
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
