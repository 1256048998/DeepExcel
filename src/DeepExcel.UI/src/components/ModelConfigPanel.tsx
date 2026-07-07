import { useState, useEffect } from 'react'
import { sendToHostWithResponse } from '../bridge'
import { providerIcons, providerOrder } from '../providerIcons'
import type { ModelConfig, ProviderInfo } from '../types'

interface Props {
  open: boolean
  onClose: () => void
}

const KEEP_PLACEHOLDER = '***keep***'

export function ModelConfigPanel({ open, onClose }: Props) {
  const [config, setConfig] = useState<ModelConfig | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [selectedProvider, setSelectedProvider] = useState('')
  const [model, setModel] = useState('')
  const [apiKey, setApiKey] = useState(KEEP_PLACEHOLDER)
  const [baseUrl, setBaseUrl] = useState('')
  const [maxTurns, setMaxTurns] = useState(20)
  const [showAdvanced, setShowAdvanced] = useState(false)
  const [showApiKey, setShowApiKey] = useState(false)
  const [testing, setTesting] = useState(false)
  const [testResult, setTestResult] = useState<{ success: boolean; latencyMs?: number; error?: string | null } | null>(null)
  const [saving, setSaving] = useState(false)
  const [saveMsg, setSaveMsg] = useState('')

  // 打开时加载配置
  useEffect(() => {
    if (open) {
      loadConfig()
      setTestResult(null)
      setSaveMsg('')
    }
  }, [open])

  // 加载配置后同步表单
  useEffect(() => {
    if (config) {
      setSelectedProvider(config.currentProvider)
    }
  }, [config])

  // 切换 provider 时更新表单字段
  useEffect(() => {
    if (config && selectedProvider && config.providers[selectedProvider]) {
      const p = config.providers[selectedProvider]
      // 如果选中的是当前 provider，用 currentModel；否则用 defaultModel
      setModel(selectedProvider === config.currentProvider ? config.currentModel : (p.defaultModel || p.models[0] || ''))
      setApiKey(KEEP_PLACEHOLDER)
      setBaseUrl(p.baseUrl || '')
      setMaxTurns(config.general?.maxTurns ?? 20)
      setTestResult(null)
      setSaveMsg('')
    }
  }, [selectedProvider, config])

  async function loadConfig() {
    setLoading(true)
    setError('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'get_model_config', payload: {} },
        'model_config'
      )
      if (resp?.type === 'model_config' && resp.payload?.providers) {
        setConfig(resp.payload as ModelConfig)
      } else if (resp?.type === 'error') {
        setError(resp.payload?.message || '加载配置失败')
      } else {
        setError('加载配置失败：未收到响应')
      }
    } catch (e) {
      setError(`加载配置失败: ${e}`)
    } finally {
      setLoading(false)
    }
  }

  async function handleSave() {
    if (!selectedProvider || !model) {
      setSaveMsg('请先选择厂商和模型')
      return
    }
    setSaving(true)
    setSaveMsg('')
    try {
      const resp = await sendToHostWithResponse(
        {
          type: 'save_model_config',
          payload: {
            provider: selectedProvider,
            model,
            apiKey,
            baseUrl,
            maxTurns
          }
        },
        'config_saved'
      )
      if (resp?.type === 'config_saved' && resp.payload?.success) {
        setSaveMsg('✓ 已保存并应用')
        // 重新加载配置以更新 hasApiKey 状态
        await loadConfig()
      } else if (resp?.type === 'error') {
        setSaveMsg('✗ ' + (resp.payload?.message || '保存失败'))
      } else {
        setSaveMsg('✗ 保存失败：未收到响应')
      }
    } catch (e) {
      setSaveMsg(`✗ 保存失败: ${e}`)
    } finally {
      setSaving(false)
    }
  }

  async function handleTest() {
    if (!selectedProvider || !model) {
      setTestResult({ success: false, error: '请先选择厂商和模型' })
      return
    }
    // ★ 当 apiKey 是 KEEP_PLACEHOLDER 时，检查是否已配置 key
    // 已配置 key（hasApiKey=true）允许直接测试，传 ***keep*** 给后端，后端读取已存储的 key
    if (!apiKey || apiKey === KEEP_PLACEHOLDER) {
      if (!currentProviderInfo?.hasApiKey) {
        setTestResult({ success: false, error: '请先输入 API Key' })
        return
      }
      // 已配置 key，用 ***keep*** 标志让后端读取已存储的 key
    }
    setTesting(true)
    setTestResult(null)
    try {
      const resp = await sendToHostWithResponse(
        {
          type: 'test_api_key',
          payload: {
            provider: selectedProvider,
            apiKey: apiKey || KEEP_PLACEHOLDER,
            baseUrl,
            model
          }
        },
        'api_test_result',
        20000  // 测试连接超时 20s
      )
      if (resp?.type === 'api_test_result') {
        setTestResult({
          success: resp.payload?.success,
          latencyMs: resp.payload?.latencyMs,
          error: resp.payload?.error
        })
      } else {
        setTestResult({ success: false, error: '测试超时或未收到响应' })
      }
    } catch (e) {
      setTestResult({ success: false, error: `测试失败: ${e}` })
    } finally {
      setTesting(false)
    }
  }

  if (!open) return null

  const currentProviderInfo: ProviderInfo | undefined = config?.providers?.[selectedProvider]

  return (
    <div className="config-overlay" onClick={onClose}>
      <div className="config-panel" onClick={e => e.stopPropagation()}>
        <div className="config-header">
          <h3>模型配置</h3>
          <button className="config-close-btn" onClick={onClose} title="关闭">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="18" y1="6" x2="6" y2="18"></line>
              <line x1="6" y1="6" x2="18" y2="18"></line>
            </svg>
          </button>
        </div>

        {loading && !config ? (
          <div className="config-loading">加载中...</div>
        ) : error ? (
          <div className="config-error">
            <div>{error}</div>
            <button onClick={loadConfig} className="config-retry-btn">重试</button>
          </div>
        ) : config ? (
          <div className="config-body">
            {/* 左侧厂商列表 */}
            <aside className="config-provider-list">
              {providerOrder.filter(k => config.providers[k]).map(key => {
                const p = config.providers[key]
                const Icon = providerIcons[key] || providerIcons.custom
                const isActive = key === selectedProvider
                const isCurrent = key === config.currentProvider
                return (
                  <div
                    key={key}
                    className={`config-provider-item ${isActive ? 'active' : ''}`}
                    onClick={() => setSelectedProvider(key)}
                  >
                    <Icon size={24} />
                    <span className="config-provider-name">{p.displayName}</span>
                    {p.hasApiKey && <span className="config-provider-dot" title="已配置 API Key" />}
                    {isCurrent && <span className="config-provider-current" title="当前使用" />}
                  </div>
                )
              })}
            </aside>

            {/* 右侧配置区 */}
            <main className="config-provider-config">
              {currentProviderInfo ? (
                <>
                  <div className="config-section">
                    <label className="config-label">模型</label>
                    <select
                      className="config-select"
                      value={model}
                      onChange={e => setModel(e.target.value)}
                    >
                      {currentProviderInfo.models.map(m => (
                        <option key={m} value={m}>{m}</option>
                      ))}
                    </select>
                    {currentProviderInfo.supportsVision && (
                      <span className="config-badge">支持视觉</span>
                    )}
                  </div>

                  <div className="config-section">
                    <label className="config-label">API Key</label>
                    <div className="config-apikey-row">
                      <input
                        type={showApiKey ? 'text' : 'password'}
                        className="config-input"
                        value={apiKey === KEEP_PLACEHOLDER ? '' : apiKey}
                        placeholder={currentProviderInfo.hasApiKey ? `已配置（${currentProviderInfo.apiKeyPreview}）` : '输入 API Key'}
                        onChange={e => setApiKey(e.target.value)}
                      />
                      <button
                        className="config-toggle-btn"
                        onClick={() => setShowApiKey(!showApiKey)}
                        title={showApiKey ? '隐藏' : '显示'}
                        type="button"
                      >
                        {showApiKey ? '隐藏' : '显示'}
                      </button>
                    </div>
                  </div>

                  {/* 高级设置折叠区 */}
                  <div className="config-advanced-toggle" onClick={() => setShowAdvanced(!showAdvanced)}>
                    <svg
                      width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
                      style={{ transform: showAdvanced ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s' }}
                    >
                      <polyline points="9 18 15 12 9 6"></polyline>
                    </svg>
                    <span>高级设置</span>
                  </div>

                  {showAdvanced && (
                    <div className="config-advanced">
                      <div className="config-section">
                        <label className="config-label">Base URL</label>
                        <input
                          type="text"
                          className="config-input"
                          value={baseUrl}
                          onChange={e => setBaseUrl(e.target.value)}
                          placeholder="https://..."
                        />
                      </div>
                      <div className="config-section">
                        <label className="config-label">MaxTurns（工具调用循环上限）</label>
                        <input
                          type="number"
                          className="config-input config-input-narrow"
                          value={maxTurns}
                          min={1}
                          max={200}
                          onChange={e => setMaxTurns(parseInt(e.target.value) || 20)}
                        />
                      </div>
                    </div>
                  )}

                  {/* 测试结果 */}
                  {testResult && (
                    <div className={`config-test-result ${testResult.success ? 'success' : 'error'}`}>
                      {testResult.success
                        ? `✓ 连接成功（${testResult.latencyMs}ms）`
                        : `✗ ${testResult.error}`}
                    </div>
                  )}

                  {/* 保存消息 */}
                  {saveMsg && (
                    <div className={`config-save-msg ${saveMsg.startsWith('✓') ? 'success' : 'error'}`}>
                      {saveMsg}
                    </div>
                  )}

                  {/* 底部按钮 */}
                  <div className="config-actions">
                    <button
                      className="config-test-btn"
                      onClick={handleTest}
                      disabled={testing}
                      type="button"
                    >
                      {testing ? '测试中...' : '测试连接'}
                    </button>
                    <button
                      className="config-save-btn"
                      onClick={handleSave}
                      disabled={saving}
                      type="button"
                    >
                      {saving ? '保存中...' : '保存并应用'}
                    </button>
                  </div>
                </>
              ) : (
                <div className="config-empty">请选择一个厂商</div>
              )}
            </main>
          </div>
        ) : null}
      </div>
    </div>
  )
}
