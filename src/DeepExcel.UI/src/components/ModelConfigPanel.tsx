import { useState, useEffect, useRef } from 'react'
import { sendToHostWithResponse } from '../bridge'
import { providerIcons, providerOrder } from '../providerIcons'
import type { ModelConfig, ProviderInfo } from '../types'

interface Props {
  open: boolean
  onClose: () => void
}

const KEEP_PLACEHOLDER = '***keep***'
/** ★ 单个厂商最多保留的模型数，与后端 HandleSetProviderModels 的上限保持一致 */
const MAX_MODELS = 100

/** ★ 数组内移动元素（拖拽排序用），返回新数组 */
function moveItem<T>(list: T[], from: number, to: number): T[] {
  if (from === to || from < 0 || to < 0 || from >= list.length || to >= list.length) return list
  const next = list.slice()
  const [item] = next.splice(from, 1)
  next.splice(to, 0, item)
  return next
}

/** ★ 优先级标签：第 0 个是主模型，其余是备用 1、备用 2 …… */
function priorityLabel(index: number): string {
  return index === 0 ? '主模型' : `备用 ${index}`
}

export function ModelConfigPanel({ open, onClose }: Props) {
  const [config, setConfig] = useState<ModelConfig | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [selectedProvider, setSelectedProvider] = useState('')
  // ★ 当前厂商的"模型优先级"有序列表：数组顺序即优先级，models[0] 为主模型（= DefaultModel）
  const [models, setModels] = useState<string[]>([])
  // 拖拽中的行索引（null 表示没有拖拽）
  const [dragIndex, setDragIndex] = useState<number | null>(null)
  const [modelsMsg, setModelsMsg] = useState('')
  // 手动添加模型的输入态
  const [addingModel, setAddingModel] = useState(false)
  const [newModelName, setNewModelName] = useState('')
  // ★ 导入模型弹窗
  const [importOpen, setImportOpen] = useState(false)
  const [importAll, setImportAll] = useState<string[]>([])
  const [importChecked, setImportChecked] = useState<string[]>([])
  const [importQuery, setImportQuery] = useState('')
  const [apiKey, setApiKey] = useState(KEEP_PLACEHOLDER)
  const [baseUrl, setBaseUrl] = useState('')
  const [maxTurns, setMaxTurns] = useState(20)
  const [showAdvanced, setShowAdvanced] = useState(false)
  // ★ API Key 显示状态：'hidden'（默认 password 模式）/ 'revealed'（明文可编辑）/ 'empty'（未配置）
  const [apiKeyMode, setApiKeyMode] = useState<'hidden' | 'revealed' | 'empty'>('hidden')
  const [apiKeyLoading, setApiKeyLoading] = useState(false)
  const [deletingKey, setDeletingKey] = useState(false)
  const [testing, setTesting] = useState(false)
  const [refreshingModels, setRefreshingModels] = useState(false)
  const [modelRefreshMsg, setModelRefreshMsg] = useState('')
  const [testResult, setTestResult] = useState<{ success: boolean; latencyMs?: number; error?: string | null } | null>(null)
  const [saving, setSaving] = useState(false)
  const [saveMsg, setSaveMsg] = useState('')
  // ★ 全局默认厂商切换的提示消息（per-provider "设为默认模型"按钮已移除，避免"都变成了默认"混淆）
  const [defaultProviderMsg, setDefaultProviderMsg] = useState('')
  // ★ 测试进行中的标志：防止 loadConfig 触发的 useEffect 清空 testResult
  // 否则测试完成后调用 loadConfig，useEffect 会 setTestResult(null) 把刚显示的结果清掉
  const testingRef = useRef(false)

  // ★ 拖拽过程中需要读取"最新"的模型顺序（setState 是异步的），用 ref 同步一份
  const modelsRef = useRef<string[]>([])
  function applyModels(next: string[]) {
    modelsRef.current = next
    setModels(next)
  }

  // 打开时加载配置
  useEffect(() => {
    if (open) {
      loadConfig()
      setTestResult(null)
      setSaveMsg('')
      setDefaultProviderMsg('')
      setModelRefreshMsg('')
      setModelsMsg('')
      setImportOpen(false)
      setAddingModel(false)
    }
  }, [open])

  // 加载配置后同步表单
  useEffect(() => {
    if (config) {
      setSelectedProvider(config.currentProvider)
    }
  }, [config])

  // 切换 provider 时更新表单字段
  // ★ 用 ref 区分"用户主动切换 provider"和"loadConfig 触发的 config 变化"：
  //   - 用户切换 provider：userSwitchedRef.current=true，执行清空
  //   - loadConfig 触发 config 变化：userSwitchedRef.current=false，跳过清空（保护 testResult）
  const userSwitchedRef = useRef(false)
  useEffect(() => {
    if (config && selectedProvider && config.providers[selectedProvider]) {
      const p = config.providers[selectedProvider]
      // ★ 模型优先级列表：直接用后端保存的数组顺序（第 0 个为主模型）
      applyModels(Array.isArray(p.models) ? p.models : [])
      // ★ 重置 API Key 状态：已配置→hidden（password 模式显示 ********），未配置→empty
      setApiKey(KEEP_PLACEHOLDER)
      setApiKeyMode(p.hasApiKey ? 'hidden' : 'empty')
      setBaseUrl(p.baseUrl || '')
      setMaxTurns(config.general?.maxTurns ?? 20)
      // 只有用户主动切换 provider 才清空 testResult，loadConfig 触发的不清空
      if (userSwitchedRef.current) {
        setTestResult(null)
        setSaveMsg('')
        setDefaultProviderMsg('')
        setModelRefreshMsg('')
        setModelsMsg('')
        setAddingModel(false)
        setNewModelName('')
        userSwitchedRef.current = false
      }
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

  // ★ 点击"显示"按钮：从后端拉取完整 key，切换到明文可编辑模式
  async function handleRevealApiKey() {
    if (!selectedProvider) return
    setApiKeyLoading(true)
    try {
      const resp = await sendToHostWithResponse(
        { type: 'get_api_key', payload: { provider: selectedProvider } },
        'api_key'
      )
      if (resp?.type === 'api_key') {
        const fullKey = resp.payload?.apiKey || ''
        setApiKey(fullKey)
        setApiKeyMode('revealed')
      } else if (resp?.type === 'error') {
        setSaveMsg('✗ ' + (resp.payload?.message || '获取 API Key 失败'))
      }
    } catch (e) {
      setSaveMsg(`✗ 获取 API Key 失败: ${e}`)
    } finally {
      setApiKeyLoading(false)
    }
  }

  // ★ 点击"隐藏"按钮：切回 password 模式，key 保留在 input 中但不显示明文
  function handleHideApiKey() {
    setApiKeyMode('hidden')
  }

  // ★ 点击"清空"按钮：删除已配置的 key
  async function handleDeleteApiKey() {
    if (!selectedProvider) return
    if (!confirm(`确定要删除 ${currentProviderInfo?.displayName || selectedProvider} 的 API Key 吗？`)) return
    setDeletingKey(true)
    setSaveMsg('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'delete_api_key', payload: { provider: selectedProvider } },
        'api_key_deleted'
      )
      if (resp?.type === 'api_key_deleted' && resp.payload?.success) {
        setApiKey(KEEP_PLACEHOLDER)
        setApiKeyMode('empty')
        setSaveMsg('✓ API Key 已删除')
        await loadConfig()
      } else if (resp?.type === 'error') {
        setSaveMsg('✗ ' + (resp.payload?.message || '删除失败'))
      } else {
        setSaveMsg('✗ 删除失败：未收到响应')
      }
    } catch (e) {
      setSaveMsg(`✗ 删除失败: ${e}`)
    } finally {
      setDeletingKey(false)
    }
  }

  // ★ 点击"设为默认厂商"：将该 provider 设为全局默认厂商（DefaultProvider）。
  // 全局唯一，后端会覆盖旧值。前端 provider 列表把默认厂商排第一。
  // 注意：与 per-provider DefaultModel 不同——DefaultModel 是该 provider 内部的默认模型，
  // DefaultProvider 是全局唯一默认厂商。
  async function handleSetDefaultProvider(providerKey: string) {
    if (!providerKey) return
    setDefaultProviderMsg('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'set_default_provider', payload: { provider: providerKey } },
        'default_provider_set'
      )
      if (resp?.type === 'default_provider_set' && resp.payload?.success) {
        setDefaultProviderMsg('✓ 已设为默认厂商')
        await loadConfig()
      } else if (resp?.type === 'error') {
        setDefaultProviderMsg('✗ ' + (resp.payload?.message || '设置失败'))
      } else {
        setDefaultProviderMsg('✗ 设置失败：未收到响应')
      }
    } catch (e) {
      setDefaultProviderMsg(`✗ 设置失败: ${e}`)
    }
  }

  // ★ 保存"模型优先级"有序列表到后端（导入 / 新增 / 删除 / 拖拽排序后调用）。
  // 立即落盘而不是等"保存并应用"：排序是独立于"切换当前厂商"的操作，
  // 用户拖完就关掉弹窗也不应该丢失顺序。
  async function persistModels(next: string[], okMsg: string): Promise<boolean> {
    if (!selectedProvider) return false
    if (next.length === 0) {
      setModelsMsg('✗ 至少需要保留一个模型')
      return false
    }
    try {
      const resp = await sendToHostWithResponse(
        { type: 'set_provider_models', payload: { provider: selectedProvider, models: next } },
        'provider_models_saved'
      )
      if (resp?.type === 'provider_models_saved' && resp.payload?.success) {
        const saved: string[] = Array.isArray(resp.payload.models) ? resp.payload.models : next
        applyModels(saved)
        setConfig(previous => {
          if (!previous || !previous.providers[selectedProvider]) return previous
          return {
            ...previous,
            currentModel: resp.payload.currentModel || previous.currentModel,
            providers: {
              ...previous.providers,
              [selectedProvider]: {
                ...previous.providers[selectedProvider],
                models: saved,
                defaultModel: resp.payload.defaultModel || saved[0]
              }
            }
          }
        })
        setModelsMsg(okMsg)
        return true
      }
      setModelsMsg('✗ ' + (resp?.payload?.error || resp?.payload?.message || '保存模型列表失败'))
    } catch (e) {
      setModelsMsg(`✗ 保存模型列表失败: ${e}`)
    }
    // 保存失败时回滚到后端的真实顺序，避免界面和配置不一致
    const fallback = config?.providers?.[selectedProvider]?.models
    if (Array.isArray(fallback)) applyModels(fallback)
    return false
  }

  // ★ 拖拽排序：dragover 时实时换位（所见即所得），dragend 时才落盘
  function handleDragOverItem(index: number) {
    if (dragIndex === null || dragIndex === index) return
    applyModels(moveItem(modelsRef.current, dragIndex, index))
    setDragIndex(index)
  }

  function handleDragEnd() {
    if (dragIndex === null) return
    setDragIndex(null)
    const current = modelsRef.current
    const saved = config?.providers?.[selectedProvider]?.models || []
    // 顺序没变就不用打扰后端
    if (saved.length === current.length && saved.every((m, i) => m === current[i])) return
    persistModels(current, '✓ 优先级已保存')
  }

  // ★ 上移/下移：任务窗格很窄，拖拽不好操作，保留按钮方式
  function handleMove(index: number, delta: number) {
    const target = index + delta
    if (target < 0 || target >= models.length) return
    persistModels(moveItem(models, index, target), '✓ 优先级已保存')
  }

  function handleRemoveModel(name: string) {
    if (models.length <= 1) {
      setModelsMsg('✗ 至少需要保留一个模型')
      return
    }
    persistModels(models.filter(m => m !== name), '✓ 已移除该模型')
  }

  function handleAddModel() {
    const name = newModelName.trim()
    if (!name) return
    if (models.some(m => m.toLowerCase() === name.toLowerCase())) {
      setModelsMsg('✗ 该模型已在列表中')
      return
    }
    if (models.length >= MAX_MODELS) {
      setModelsMsg(`✗ 最多添加 ${MAX_MODELS} 个模型`)
      return
    }
    setNewModelName('')
    setAddingModel(false)
    persistModels([...models, name], '✓ 已添加模型')
  }

  // ★ 导入弹窗确认：勾选项按"已有顺序优先 + 新增追加"合并，保持用户排好的优先级
  function handleConfirmImport() {
    const checkedSet = new Set(importChecked)
    const kept = models.filter(m => checkedSet.has(m))
    const added = importAll.filter(m => checkedSet.has(m) && !models.includes(m))
    const next = [...kept, ...added].slice(0, MAX_MODELS)
    if (next.length === 0) {
      setModelsMsg('✗ 至少需要保留一个模型')
      return
    }
    setImportOpen(false)
    persistModels(next, `✓ 已更新模型列表（${next.length} 个）`)
  }

  function toggleImportChecked(name: string) {
    setImportChecked(prev =>
      prev.includes(name) ? prev.filter(m => m !== name) : [...prev, name]
    )
  }

  async function handleSave() {
    const model = models[0] || ''
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
        // ★ 保存后重置 API Key 状态：已配置→hidden，未配置→empty
        const p = config?.providers?.[selectedProvider]
        if (p) {
          setApiKey(KEEP_PLACEHOLDER)
          setApiKeyMode(p.hasApiKey ? 'hidden' : 'empty')
        }
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
    // ★ 用主模型（优先级第 1 个）测试连接
    const model = models[0] || ''
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
    testingRef.current = true  // ★ 标记测试进行中，防止 loadConfig 触发的 useEffect 清空 testResult
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
        // ★ 测试完成后重新加载 config，让 provider 列表圆点根据 Connected 状态立即更新
        await loadConfig()
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
      testingRef.current = false  // ★ 测试结束，恢复 useEffect 对 testResult 的清空行为
    }
  }

  // ★ 从服务商拉取可用模型 → 打开"导入模型"弹窗由用户勾选。
  // 拉取本身不再直接覆盖配置，避免把用户排好的优先级冲掉。
  async function handleRefreshModels() {
    if (!selectedProvider) return
    if ((!apiKey || apiKey === KEEP_PLACEHOLDER) && !currentProviderInfo?.hasApiKey) {
      setModelRefreshMsg('✗ 请先输入 API Key')
      return
    }
    if (!baseUrl) {
      setModelRefreshMsg('✗ 请先填写 Base URL')
      return
    }

    setRefreshingModels(true)
    setModelRefreshMsg('')
    setModelsMsg('')
    try {
      const resp = await sendToHostWithResponse(
        {
          type: 'refresh_models',
          payload: {
            provider: selectedProvider,
            apiKey: apiKey || KEEP_PLACEHOLDER,
            baseUrl
          }
        },
        'models_refreshed',
        45000  // 后端最多串行尝试 5 个候选端点（每个 8s）
      )
      if (resp?.type === 'models_refreshed' && resp.payload?.success) {
        const fetched: string[] = Array.isArray(resp.payload.models) ? resp.payload.models : []
        // 已在列表里但服务商没返回的（手工添加的模型）也列出来，避免用户无感知丢失
        const extras = models.filter(m => !fetched.includes(m))
        setImportAll([...fetched, ...extras])
        setImportChecked(models.slice())
        setImportQuery('')
        setImportOpen(true)
        setModelRefreshMsg(`✓ 拉取到 ${fetched.length} 个模型`)
      } else {
        setModelRefreshMsg('✗ ' + (resp?.payload?.error || '刷新失败：未收到响应'))
      }
    } catch (e) {
      setModelRefreshMsg(`✗ 刷新失败: ${e}`)
    } finally {
      setRefreshingModels(false)
    }
  }

  if (!open) return null

  const currentProviderInfo: ProviderInfo | undefined = config?.providers?.[selectedProvider]
  // ★ 拉取模型列表的前置条件：先有 API Key（已保存的或刚输入的）+ Base URL。
  // 这也是面板的步骤顺序：① API Key → ② Base URL → ③ 拉取模型 → ④ 排优先级。
  const hasUsableKey = !!currentProviderInfo?.hasApiKey ||
    (!!apiKey && apiKey !== KEEP_PLACEHOLDER)
  const canFetchModels = hasUsableKey && !!baseUrl

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
            {/* ★ 排序：默认厂商置顶，其余按 providerOrder 原顺序 */}
            <aside className="config-provider-list">
              {(() => {
                const defaultProvider = config.defaultProvider || config.currentProvider
                const allKeys = providerOrder.filter(k => config.providers[k])
                // 默认厂商排第一
                const sortedKeys = [
                  ...allKeys.filter(k => k === defaultProvider),
                  ...allKeys.filter(k => k !== defaultProvider)
                ]
                return sortedKeys.map(key => {
                  const p = config.providers[key]
                  const Icon = providerIcons[key] || providerIcons.custom
                  const isActive = key === selectedProvider
                  const isCurrent = key === config.currentProvider
                  const isDefault = key === defaultProvider
                  return (
                    <div
                      key={key}
                      className={`config-provider-item ${isActive ? 'active' : ''} ${isDefault ? 'is-default' : ''}`}
                      onClick={() => {
                        userSwitchedRef.current = true  // ★ 标记用户主动切换，让 useEffect 清空 testResult
                        setSelectedProvider(key)
                      }}
                    >
                      <Icon size={24} />
                      <span className="config-provider-name">{p.displayName}</span>
                      {/* ★ 默认厂商星形开关：点击设为默认（已默认则点击无操作）。
                          全局唯一，只有 1 家能成为默认。 */}
                      <button
                        className={`config-default-star ${isDefault ? 'is-default' : ''}`}
                        onClick={e => {
                          e.stopPropagation()
                          if (!isDefault) handleSetDefaultProvider(key)
                        }}
                        title={isDefault ? '默认厂商' : '设为默认厂商'}
                        type="button"
                      >
                        <svg width="14" height="14" viewBox="0 0 24 24"
                          fill={isDefault ? 'currentColor' : 'none'}
                          stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                          <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" />
                        </svg>
                      </button>
                      {/* ★ 圆点表示"连接成功"，而非"已配置 key"。
                          未测试/测试失败/删除 key 后都不显示圆点 */}
                      {p.connected && <span className="config-provider-dot" title="连接测试通过" />}
                      {isCurrent && <span className="config-provider-current" title="当前使用" />}
                    </div>
                  )
                })
              })()}
            </aside>

            {/* 右侧配置区 */}
            <main className="config-provider-config">
              {currentProviderInfo ? (
                <>
                  {/* ★ 厂商标题：显示当前厂商 + 默认厂商/视觉能力徽章 */}
                  <div className="config-provider-header">
                    <span className="config-provider-title">{currentProviderInfo.displayName}</span>
                    {(config.defaultProvider === selectedProvider ||
                      (!config.defaultProvider && config.currentProvider === selectedProvider)) && (
                      <span className="config-default-badge" title="默认厂商">默认厂商</span>
                    )}
                    {currentProviderInfo.supportsVision && (
                      <span className="config-badge">支持视觉</span>
                    )}
                  </div>

                  <div className="config-section">
                    <label className="config-label"><span className="config-step">1</span>API Key</label>
                    <div className="config-apikey-row">
                      {/* ★ 三种模式：
                          - hidden: 已配置，password 模式显示虚拟 ********，只读
                          - revealed: 已配置且用户点击"显示"，text 模式显示真实 key，可编辑
                          - empty: 未配置，password 模式空输入框，可输入新 key */}
                      <input
                        type={apiKeyMode === 'revealed' ? 'text' : 'password'}
                        className="config-input"
                        value={
                          apiKeyMode === 'hidden' ? '••••••••••••••••' :
                          apiKey === KEEP_PLACEHOLDER ? '' : apiKey
                        }
                        placeholder={apiKeyMode === 'empty' ? '输入 API Key' : ''}
                        readOnly={apiKeyMode === 'hidden'}
                        onChange={e => {
                          // 用户在 empty 或 revealed 模式下输入新值
                          setApiKey(e.target.value)
                          if (apiKeyMode === 'empty') {
                            // 用户开始输入新 key，切换到 revealed 模式
                            // （保留 apiKeyMode='empty' 也行，但 revealed 更直观）
                          }
                        }}
                      />
                      {/* 显示/隐藏按钮：仅在已配置状态下显示 */}
                      {apiKeyMode === 'hidden' && (
                        <button
                          className="config-toggle-btn"
                          onClick={handleRevealApiKey}
                          disabled={apiKeyLoading}
                          title="显示 API Key"
                          type="button"
                        >
                          {apiKeyLoading ? '...' : '显示'}
                        </button>
                      )}
                      {apiKeyMode === 'revealed' && (
                        <button
                          className="config-toggle-btn"
                          onClick={handleHideApiKey}
                          title="隐藏 API Key"
                          type="button"
                        >
                          隐藏
                        </button>
                      )}
                      {/* 清空按钮：已配置状态下显示 */}
                      {(apiKeyMode === 'hidden' || apiKeyMode === 'revealed') && (
                        <button
                          className="config-toggle-btn config-delete-btn"
                          onClick={handleDeleteApiKey}
                          disabled={deletingKey}
                          title="删除已配置的 API Key"
                          type="button"
                        >
                          {deletingKey ? '...' : '清空'}
                        </button>
                      )}
                    </div>
                    {apiKeyMode === 'hidden' && currentProviderInfo.apiKeyPreview && (
                      <div className="config-apikey-hint">已配置（{currentProviderInfo.apiKeyPreview}）</div>
                    )}
                  </div>

                  <div className="config-section">
                    <label className="config-label"><span className="config-step">2</span>Base URL</label>
                    <input
                      type="text"
                      className="config-input"
                      value={baseUrl}
                      onChange={e => setBaseUrl(e.target.value)}
                      placeholder="https://..."
                    />
                    <div className="config-apikey-hint">留空或保持默认即可，自建网关/中转站在此填写</div>
                  </div>

                  {/* ★ 第 3 步：填好 Key 和 Base URL 后，从服务商拉取可选模型 */}
                  <div className="config-section">
                    <label className="config-label"><span className="config-step">3</span>获取模型列表</label>
                    <button
                      className="model-fetch-btn"
                      onClick={handleRefreshModels}
                      disabled={refreshingModels || testing || !canFetchModels}
                      title={canFetchModels ? '从服务商拉取可用模型' : '请先填写 API Key 和 Base URL'}
                      type="button"
                    >
                      {refreshingModels ? '拉取中...' : '↻ 从服务商拉取模型列表'}
                    </button>
                    {!canFetchModels && (
                      <div className="config-apikey-hint">填好上面的 API Key 和 Base URL 后即可拉取</div>
                    )}
                    {modelRefreshMsg && (
                      <div className={`config-default-msg ${modelRefreshMsg.startsWith('✓') ? 'success' : 'error'}`}>
                        {modelRefreshMsg}
                      </div>
                    )}
                  </div>

                  <div className="config-section">
                    <div className="config-label-row">
                      <label className="config-label">
                        <span className="config-step">4</span>模型优先级（至少添加一个）
                      </label>
                    </div>

                    {/* ★ 可拖拽排序的模型优先级列表：数组顺序即优先级，第 1 项为主模型 */}
                    <ul className="model-priority-list">
                      {models.map((m, i) => (
                        <li
                          key={m}
                          className={`model-priority-item ${dragIndex === i ? 'dragging' : ''}`}
                          draggable
                          onDragStart={e => {
                            setDragIndex(i)
                            e.dataTransfer.effectAllowed = 'move'
                          }}
                          onDragOver={e => {
                            e.preventDefault()
                            e.dataTransfer.dropEffect = 'move'
                            handleDragOverItem(i)
                          }}
                          onDrop={e => e.preventDefault()}
                          onDragEnd={handleDragEnd}
                        >
                          <span className="model-drag-handle" title="拖拽调整优先级">
                            <svg width="10" height="14" viewBox="0 0 10 14" fill="currentColor">
                              <circle cx="2.5" cy="2" r="1.2" /><circle cx="7.5" cy="2" r="1.2" />
                              <circle cx="2.5" cy="7" r="1.2" /><circle cx="7.5" cy="7" r="1.2" />
                              <circle cx="2.5" cy="12" r="1.2" /><circle cx="7.5" cy="12" r="1.2" />
                            </svg>
                          </span>
                          <span className={`model-priority-tag ${i === 0 ? 'primary' : ''}`}>
                            {priorityLabel(i)}
                          </span>
                          <span className="model-priority-name" title={m}>{m}</span>
                          <div className="model-priority-actions">
                            <button
                              className="model-icon-btn"
                              onClick={() => handleMove(i, -1)}
                              disabled={i === 0}
                              title="上移"
                              type="button"
                            >↑</button>
                            <button
                              className="model-icon-btn"
                              onClick={() => handleMove(i, 1)}
                              disabled={i === models.length - 1}
                              title="下移"
                              type="button"
                            >↓</button>
                            <button
                              className="model-icon-btn danger"
                              onClick={() => handleRemoveModel(m)}
                              disabled={models.length <= 1}
                              title="从列表移除"
                              type="button"
                            >✕</button>
                          </div>
                        </li>
                      ))}

                      {/* 手动添加模型（服务商没开放列表接口时的兜底入口） */}
                      <li className="model-priority-add">
                        {addingModel ? (
                          <div className="model-add-row">
                            <input
                              className="config-input"
                              autoFocus
                              value={newModelName}
                              placeholder="输入模型 ID，例如 deepseek-v4-pro"
                              onChange={e => setNewModelName(e.target.value)}
                              onKeyDown={e => {
                                if (e.key === 'Enter') handleAddModel()
                                if (e.key === 'Escape') { setAddingModel(false); setNewModelName('') }
                              }}
                            />
                            <button className="config-toggle-btn" onClick={handleAddModel} type="button">添加</button>
                            <button
                              className="config-toggle-btn"
                              onClick={() => { setAddingModel(false); setNewModelName('') }}
                              type="button"
                            >取消</button>
                          </div>
                        ) : (
                          <button
                            className="model-add-btn"
                            onClick={() => { setAddingModel(true); setModelsMsg('') }}
                            type="button"
                          >＋ 添加模型</button>
                        )}
                      </li>
                    </ul>

                    <div className="model-priority-hint">
                      拖拽或用 ↑↓ 调整优先级，第 1 项为主模型（默认使用）
                    </div>

                    {modelsMsg && (
                      <div className={`config-default-msg ${modelsMsg.startsWith('✓') ? 'success' : 'error'}`}>
                        {modelsMsg}
                      </div>
                    )}
                    {defaultProviderMsg && (
                      <div className={`config-default-msg ${defaultProviderMsg.startsWith('✓') ? 'success' : 'error'}`}>
                        {defaultProviderMsg}
                      </div>
                    )}
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
                        <label className="config-label">MaxTurns（工具调用循环上限）</label>
                        <input
                          type="number"
                          className="config-input config-input-narrow"
                          value={maxTurns}
                          min={1}
                          max={50}
                          // 与 C# HandleSaveModelConfig 的上限一致：超过 50 会被后端静默丢弃
                          onChange={e => setMaxTurns(Math.min(50, Math.max(1, parseInt(e.target.value) || 20)))}
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

        {/* ★ 导入模型弹窗：勾选要加入"模型优先级"列表的模型 */}
        {importOpen && (() => {
          const query = importQuery.trim().toLowerCase()
          const visible = query ? importAll.filter(m => m.toLowerCase().includes(query)) : importAll
          return (
            <div className="import-overlay" onClick={() => setImportOpen(false)}>
              <div className="import-dialog" onClick={e => e.stopPropagation()}>
                <div className="import-header">
                  <h4>导入模型</h4>
                  <button className="config-close-btn" onClick={() => setImportOpen(false)} title="关闭" type="button">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <line x1="18" y1="6" x2="6" y2="18"></line>
                      <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                  </button>
                </div>
                <div className="import-desc">
                  从服务商拉取到 {importAll.length} 个模型，已添加模型会默认勾选；
                  取消勾选会从当前列表删除（最多 {MAX_MODELS} 个）
                </div>
                <input
                  className="config-input"
                  value={importQuery}
                  placeholder="搜索模型 ID"
                  onChange={e => setImportQuery(e.target.value)}
                />
                <div className="import-meta">
                  <span>已选 {importChecked.length} 个</span>
                  <button className="import-clear-btn" onClick={() => setImportChecked([])} type="button">清空</button>
                </div>
                <ul className="import-list">
                  {visible.length === 0 && (
                    <li className="import-empty">没有匹配的模型</li>
                  )}
                  {visible.map(m => {
                    const checked = importChecked.includes(m)
                    return (
                      <li
                        key={m}
                        className={`import-item ${checked ? 'checked' : ''}`}
                        onClick={() => toggleImportChecked(m)}
                      >
                        <span className={`import-check ${checked ? 'on' : ''}`}>{checked ? '✓' : ''}</span>
                        <span className="import-name" title={m}>{m}</span>
                        {models.includes(m) && <span className="import-added">已添加</span>}
                      </li>
                    )
                  })}
                </ul>
                <div className="import-actions">
                  <button className="config-test-btn" onClick={() => setImportOpen(false)} type="button">取消</button>
                  <button
                    className="config-save-btn"
                    onClick={handleConfirmImport}
                    disabled={importChecked.length === 0}
                    type="button"
                  >
                    更新列表（{importChecked.length}/{importAll.length}）
                  </button>
                </div>
              </div>
            </div>
          )
        })()}
      </div>
    </div>
  )
}
