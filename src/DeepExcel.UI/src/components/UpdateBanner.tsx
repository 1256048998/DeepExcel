import { useCallback, useEffect, useState } from 'react'
import { sendToHostWithResponse } from '../bridge'

/**
 * 新版本就绪提示。
 *
 * 只在"已下载并验签通过、等待重启"时出现，占一行。检查、下载、校验全部在后台
 * 完成，用户第一次听说这件事就是它可以装了——自动更新的价值就在于用户不参与
 * 过程，所以这里刻意不显示"正在检查""正在下载"之类的中间态。
 *
 * 关掉只在本次会话内有效：下次打开面板还会提示，但不会在同一次会话里反复打扰。
 */

interface UpdateStatus {
  state: string
  detail?: string | null
  installed_version?: string
  available_version?: string | null
  notes?: string | null
  size?: number
}

const DISMISS_KEY = 'deepexcel.update.dismissed'
const POLL_INTERVAL_MS = 5 * 60 * 1000

function formatSize(bytes?: number): string {
  if (!bytes || bytes <= 0) return ''
  return ` · ${(bytes / 1024 / 1024).toFixed(0)} MB`
}

export function UpdateBanner() {
  const [status, setStatus] = useState<UpdateStatus | null>(null)
  const [dismissed, setDismissed] = useState<string | null>(
    () => sessionStorage.getItem(DISMISS_KEY)
  )
  const [installing, setInstalling] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const poll = useCallback(async () => {
    const response = await sendToHostWithResponse(
      { type: 'update_status', payload: {} },
      'update_status'
    )
    if (response?.type === 'update_status') {
      setStatus(response.payload as UpdateStatus)
    }
  }, [])

  useEffect(() => {
    void poll()
    const timer = setInterval(() => void poll(), POLL_INTERVAL_MS)
    return () => clearInterval(timer)
  }, [poll])

  const version = status?.available_version
  // blocked 与 ready 一起显示，是因为 blocked 恰恰是用户唯一需要动手的情况：
  // 更新已经下好、验签也过了，但在这台机器上装不上（多半被安全软件拦了）。
  // 把它藏起来等于让用户一直停在旧版本且毫不知情。
  const blocked = status?.state === 'blocked'
  if (!status || (status.state !== 'ready' && !blocked) || !version) return null
  if (dismissed === version) return null

  const handleInstall = async () => {
    setInstalling(true)
    setError(null)
    // Excel 会先弹未保存提示，用户可能取消；那不是错误，更新会留到下次重启。
    const response = await sendToHostWithResponse(
      { type: 'update_install', payload: {} },
      'update_install',
      30000
    )
    if (!response || response.type === 'error') {
      setInstalling(false)
      setError(response?.payload?.message ?? '启动更新失败')
    }
  }

  const handleDismiss = () => {
    sessionStorage.setItem(DISMISS_KEY, version)
    setDismissed(version)
  }

  return (
    <div className={`update-banner${blocked ? ' blocked' : ''}`} role="status">
      {blocked ? (
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"
             strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
          <line x1="12" y1="9" x2="12" y2="13" />
          <line x1="12" y1="17" x2="12.01" y2="17" />
        </svg>
      ) : (
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"
             strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
          <polyline points="7 10 12 15 17 10" />
          <line x1="12" y1="15" x2="12" y2="3" />
        </svg>
      )}
      <span className="update-banner-text" title={status.detail ?? status.notes ?? undefined}>
        {error ??
          (blocked
            ? `v${version} 多次安装未成功，请手动下载安装`
            : `新版本 v${version} 已就绪${formatSize(status.size)}`)}
      </span>
      {!blocked && (
      <button
        className="header-btn primary"
        onClick={handleInstall}
        disabled={installing}
        title="关闭 Excel 后自动安装，安装完成会重新打开 Excel"
        type="button"
      >
        {installing ? '正在关闭 Excel…' : '重启安装'}
      </button>
      )}
      <button
        className="update-banner-close"
        onClick={handleDismiss}
        title="本次会话内不再提示"
        aria-label="关闭"
        type="button"
      >
        ×
      </button>
    </div>
  )
}
