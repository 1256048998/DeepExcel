// src/DeepExcel.Wps/config-store.js
// 模型配置读写（对应 C# 端 Config/AppConfig.cs + ConfigManager）
//
// ★ 关键设计：直接复用 Excel 加载项的配置文件 %APPDATA%\DeepExcel\config.json，
// 键名沿用 C# 的 PascalCase（C# 用默认命名策略序列化），读取时大小写不敏感。
// 这样用户在 Excel 里配好的厂商/模型优先级/API Key，在 WPS 里立刻可用，反之亦然。
//
// 发给前端的 payload 则是 camelCase，与 C# SecurityManager.GetSafeConfig 经
// MessageBridge（camelCase 策略）序列化后的结构一一对应，前端组件无需区分宿主。

const fs = require('fs')
const path = require('path')
const CredentialStore = require('./credential-store')

/** ★ 内置模型目录，必须与 C# ConfigManager.LatestModelCatalog 保持一致 */
const MODEL_CATALOG = {
  anthropic: {
    models: ['claude-sonnet-5', 'claude-opus-5', 'claude-opus-4.8', 'claude-haiku-5', 'claude-haiku-4-5-20251001'],
    defaultModel: 'claude-sonnet-5',
  },
  deepseek: { models: ['deepseek-v4-pro', 'deepseek-v4-flash'], defaultModel: 'deepseek-v4-pro' },
  stepfun: { models: ['step-3.7-flash', 'step-3.5-flash'], defaultModel: 'step-3.7-flash' },
  openai: { models: ['gpt-5.5', 'gpt-5.5-pro', 'gpt-5'], defaultModel: 'gpt-5.5' },
  kimi: { models: ['kimi-k2.7-code', 'kimi-k2.6', 'kimi-k2-thinking'], defaultModel: 'kimi-k2.7-code' },
  qwen: { models: ['qwen3.7-max', 'qwen3-max', 'qwen3-coder-plus'], defaultModel: 'qwen3.7-max' },
  zhipu: { models: ['glm-5.2', 'glm-5.1', 'glm-4.7-flash'], defaultModel: 'glm-5.2' },
  minimax: { models: ['MiniMax-M2.5', 'MiniMax-M2'], defaultModel: 'MiniMax-M2.5' },
  doubao: { models: ['doubao-seed-2.1-pro', 'doubao-seed-2.1', 'doubao-seed-1.6'], defaultModel: 'doubao-seed-2.1-pro' },
}

/** ★ 厂商元信息，必须与 C# AppConfig.CreateDefault 保持一致 */
const PROVIDER_DEFAULTS = {
  anthropic: { type: 'anthropic', displayName: 'Claude (Anthropic)', baseUrl: 'https://api.anthropic.com', supportsVision: true },
  deepseek: { type: 'anthropic', displayName: 'DeepSeek', baseUrl: 'https://api.deepseek.com/anthropic', supportsVision: false },
  stepfun: { type: 'anthropic', displayName: '阶跃星辰 (Step)', baseUrl: 'https://api.stepfun.com/step_plan', supportsVision: true },
  openai: { type: 'openai', displayName: 'OpenAI', baseUrl: 'https://api.openai.com/v1', supportsVision: true },
  kimi: { type: 'anthropic', displayName: 'Kimi (月之暗面)', baseUrl: 'https://api.moonshot.cn/anthropic', supportsVision: true },
  qwen: { type: 'anthropic', displayName: '通义千问 (阿里)', baseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/anthropic', supportsVision: true },
  zhipu: { type: 'anthropic', displayName: '智谱 (GLM)', baseUrl: 'https://api.z.ai/api/anthropic', supportsVision: true },
  minimax: { type: 'anthropic', displayName: 'Minimax', baseUrl: 'https://api.minimax.io/anthropic', supportsVision: false },
  doubao: { type: 'anthropic', displayName: '豆包 (火山引擎)', baseUrl: 'https://ark.cn-beijing.volces.com/api/compatible', supportsVision: true },
  custom: { type: 'openai', displayName: '自定义 (OpenAI兼容)', baseUrl: '', supportsVision: false },
}

/** ★ 与 C# AppConfig.CurrentModelCatalogVersion 保持一致 */
const CURRENT_MODEL_CATALOG_VERSION = 2
const MAX_MODELS = 100

/** 大小写不敏感取值：C# 写 PascalCase，手工编辑过的文件可能是 camelCase */
function pick(obj, key, fallback) {
  if (!obj || typeof obj !== 'object') return fallback
  if (Object.prototype.hasOwnProperty.call(obj, key)) return obj[key]
  const lower = key.toLowerCase()
  for (const k of Object.keys(obj)) {
    if (k.toLowerCase() === lower) return obj[k]
  }
  return fallback
}

function isValidModelName(id) {
  // 与 C# MessageBridge.IsValidModelName 一致：非空、长度受限、不含空白与控制字符
  if (typeof id !== 'string' || id.length === 0 || id.length > 200) return false
  for (let i = 0; i < id.length; i++) {
    const code = id.charCodeAt(i)
    if (code <= 0x20 || code === 0x7f) return false
  }
  return true
}

class ConfigStore {
  constructor(options = {}) {
    this.dir = options.dir || path.join(process.env.APPDATA || '', 'DeepExcel')
    this.filePath = options.filePath || path.join(this.dir, 'config.json')
    this.credentials = options.credentials || new CredentialStore({ pythonPath: options.pythonPath })
    this.config = null
  }

  // ============= 加载 / 保存 =============

  load() {
    let raw = null
    try {
      if (fs.existsSync(this.filePath)) raw = JSON.parse(fs.readFileSync(this.filePath, 'utf8'))
    } catch (e) {
      console.error('[ConfigStore] read config.json failed, falling back to defaults:', e.message)
    }
    this.config = this._normalize(raw)
    return this.config
  }

  current() {
    return this.config || this.load()
  }

  save() {
    try {
      if (!fs.existsSync(this.dir)) fs.mkdirSync(this.dir, { recursive: true })
      fs.writeFileSync(this.filePath, JSON.stringify(this.config, null, 2), 'utf8')
      return true
    } catch (e) {
      console.error('[ConfigStore] save failed:', e.message)
      return false
    }
  }

  /**
   * 归一化 + 迁移：补齐缺失厂商、空模型列表兜底、按版本推送内置目录。
   * 迁移规则与 C# ConfigManager.MigrateConfig 一致：
   *   - 只有本地目录版本落后才推送内置模型列表（每个版本一次）
   *   - ModelsCustomized=true（用户自己导入/排序过）的厂商永不覆盖
   */
  _normalize(raw) {
    const source = raw && typeof raw === 'object' ? raw : {}
    const catalogVersion = Number(pick(source, 'ModelCatalogVersion', 0)) || 0
    const catalogOutdated = catalogVersion < CURRENT_MODEL_CATALOG_VERSION
    const rawProviders = pick(source, 'Providers', {}) || {}

    const providers = {}
    for (const key of Object.keys(PROVIDER_DEFAULTS)) {
      const defaults = PROVIDER_DEFAULTS[key]
      const catalog = MODEL_CATALOG[key] || { models: ['custom-model'], defaultModel: 'custom-model' }
      const existing = pick(rawProviders, key, null)

      let models = pick(existing, 'Models', null)
      if (!Array.isArray(models)) models = null
      models = models ? models.filter(isValidModelName) : null

      // ★ 内置目录只"补充"不"覆盖"：用户已有的模型名一个都不删（内置目录必然滞后于厂商）
      const customized = pick(existing, 'ModelsCustomized', false) === true
      const isEmpty = !models || models.length === 0
      if (isEmpty) {
        models = catalog.models.slice()
      } else if (catalogOutdated && !customized) {
        for (const model of catalog.models) {
          if (!models.some(m => m.toLowerCase() === model.toLowerCase())) models.push(model)
        }
      }

      let defaultModel = pick(existing, 'DefaultModel', '')
      if (!defaultModel || models.indexOf(defaultModel) < 0) defaultModel = models[0]

      providers[key] = {
        Type: pick(existing, 'Type', defaults.type),
        DisplayName: pick(existing, 'DisplayName', defaults.displayName) || defaults.displayName,
        ApiKey: '', // 明文永不落盘，凭据在 credentials/*.crypt
        BaseUrl: pick(existing, 'BaseUrl', defaults.baseUrl),
        Models: models,
        DefaultModel: defaultModel,
        Headers: pick(existing, 'Headers', {}) || {},
        SupportsVision: pick(existing, 'SupportsVision', defaults.supportsVision) === true,
        LastTestSuccess: pick(existing, 'LastTestSuccess', false) === true,
        ModelsCustomized: customized,
      }
    }

    // 保留 config.json 里出现过的第三方自定义厂商（Excel 端可能扩展）
    for (const key of Object.keys(rawProviders)) {
      if (providers[key]) continue
      const existing = rawProviders[key]
      let models = pick(existing, 'Models', [])
      models = Array.isArray(models) ? models.filter(isValidModelName) : []
      if (models.length === 0) continue
      providers[key] = {
        Type: pick(existing, 'Type', 'openai'),
        DisplayName: pick(existing, 'DisplayName', key),
        ApiKey: '',
        BaseUrl: pick(existing, 'BaseUrl', ''),
        Models: models,
        DefaultModel: pick(existing, 'DefaultModel', models[0]),
        Headers: pick(existing, 'Headers', {}) || {},
        SupportsVision: pick(existing, 'SupportsVision', false) === true,
        LastTestSuccess: pick(existing, 'LastTestSuccess', false) === true,
        ModelsCustomized: pick(existing, 'ModelsCustomized', false) === true,
      }
    }

    let currentProvider = pick(source, 'CurrentProvider', 'anthropic')
    if (!providers[currentProvider]) currentProvider = 'anthropic'
    let currentModel = pick(source, 'CurrentModel', '')
    if (!currentModel || providers[currentProvider].Models.indexOf(currentModel) < 0) {
      currentModel = providers[currentProvider].DefaultModel
    }
    const defaultProvider = pick(source, 'DefaultProvider', null)

    const general = pick(source, 'General', {}) || {}
    const ui = pick(source, 'UI', {}) || {}

    return {
      CurrentProvider: currentProvider,
      CurrentModel: currentModel,
      DefaultProvider: providers[defaultProvider] ? defaultProvider : null,
      ModelCatalogVersion: CURRENT_MODEL_CATALOG_VERSION,
      Providers: providers,
      General: {
        MaxRetries: Number(pick(general, 'MaxRetries', 2)) || 2,
        RequestTimeoutSeconds: Number(pick(general, 'RequestTimeoutSeconds', 60)) || 60,
        AutoCreateSnapshot: pick(general, 'AutoCreateSnapshot', true) !== false,
        RequireConfirmation: pick(general, 'RequireConfirmation', true) !== false,
        MaxConversationHistory: Number(pick(general, 'MaxConversationHistory', 10)) || 10,
        MaxTurns: Number(pick(general, 'MaxTurns', 20)) || 20,
      },
      UI: {
        Theme: pick(ui, 'Theme', 'light'),
        Language: pick(ui, 'Language', 'zh-CN'),
        ShowTokenUsage: pick(ui, 'ShowTokenUsage', true) !== false,
        StreamOutput: pick(ui, 'StreamOutput', true) !== false,
      },
    }
  }

  // ============= 前端视图 =============

  /** 对应 C# SecurityManager.GetSafeConfig：camelCase + API Key 脱敏 */
  getSafeConfig() {
    const cfg = this.current()
    const providers = {}
    for (const key of Object.keys(cfg.Providers)) {
      const p = cfg.Providers[key]
      const apiKey = this.credentials.get(key)
      const hasApiKey = apiKey.length > 0
      providers[key] = {
        displayName: p.DisplayName,
        type: p.Type,
        baseUrl: p.BaseUrl,
        defaultModel: p.DefaultModel,
        supportsVision: p.SupportsVision,
        models: p.Models,
        hasApiKey,
        apiKeyPreview: CredentialStore.mask(apiKey),
        // Connected = 最近一次测试成功 且 key 仍然存在
        connected: p.LastTestSuccess && hasApiKey,
      }
    }
    return {
      currentProvider: cfg.CurrentProvider,
      currentModel: cfg.CurrentModel,
      defaultProvider: cfg.DefaultProvider || cfg.CurrentProvider,
      providers,
      general: {
        maxRetries: cfg.General.MaxRetries,
        requestTimeoutSeconds: cfg.General.RequestTimeoutSeconds,
        autoCreateSnapshot: cfg.General.AutoCreateSnapshot,
        requireConfirmation: cfg.General.RequireConfirmation,
        maxConversationHistory: cfg.General.MaxConversationHistory,
        maxTurns: cfg.General.MaxTurns,
      },
      ui: {
        theme: cfg.UI.Theme,
        language: cfg.UI.Language,
        showTokenUsage: cfg.UI.ShowTokenUsage,
        streamOutput: cfg.UI.StreamOutput,
      },
    }
  }

  /** 当前会话要发给 sidecar 的配置（对应 C# SendConfigToSession） */
  getSidecarConfig() {
    const cfg = this.current()
    const providerKey = cfg.CurrentProvider
    const provider = cfg.Providers[providerKey]
    if (!provider) return null
    let baseUrl = provider.BaseUrl || 'https://api.anthropic.com'
    if (providerKey === 'deepseek' && !/\/anthropic$/i.test(baseUrl)) {
      baseUrl = 'https://api.deepseek.com/anthropic'
    }
    return {
      provider: providerKey,
      baseUrl,
      model: cfg.CurrentModel || provider.DefaultModel || provider.Models[0],
      apiKey: this.credentials.get(providerKey),
    }
  }

  // ============= 变更操作 =============

  hasProvider(provider) {
    return !!(provider && this.current().Providers[provider])
  }

  /** 模型优先级列表：数组顺序即优先级，第 0 个为主模型 */
  updateProviderModels(provider, models) {
    const cfg = this.current()
    const target = cfg.Providers[provider]
    if (!target) return { success: false, error: '未知的模型供应商' }

    const cleaned = []
    for (const item of Array.isArray(models) ? models : []) {
      const id = typeof item === 'string' ? item.trim() : ''
      if (!isValidModelName(id)) continue
      if (cleaned.some(m => m.toLowerCase() === id.toLowerCase())) continue
      cleaned.push(id)
      if (cleaned.length >= MAX_MODELS) break
    }
    if (cleaned.length === 0) return { success: false, error: '至少需要保留一个模型' }

    target.Models = cleaned
    target.DefaultModel = cleaned[0]
    target.ModelsCustomized = true
    if (cfg.CurrentProvider === provider && cleaned.indexOf(cfg.CurrentModel) < 0) {
      cfg.CurrentModel = cleaned[0]
    }
    this.save()
    return { success: true, models: cleaned, defaultModel: cleaned[0], currentModel: cfg.CurrentModel }
  }

  switchProvider(provider, model) {
    const cfg = this.current()
    const target = cfg.Providers[provider]
    if (!target) return { success: false, error: `未知的 provider: ${provider}` }
    if (model && target.Models.indexOf(model) < 0) {
      return { success: false, error: `模型 ${model} 不在 ${provider} 的支持列表中` }
    }
    cfg.CurrentProvider = provider
    cfg.CurrentModel = model || target.DefaultModel || target.Models[0]
    this.save()
    return { success: true, provider, model: cfg.CurrentModel }
  }

  setDefaultProvider(provider) {
    const cfg = this.current()
    if (!cfg.Providers[provider]) return { success: false, error: `未知的 provider: ${provider}` }
    cfg.DefaultProvider = provider
    this.save()
    return { success: true, defaultProvider: provider }
  }

  setConnected(provider, connected) {
    const cfg = this.current()
    if (!cfg.Providers[provider]) return
    cfg.Providers[provider].LastTestSuccess = !!connected
    this.save()
  }

  /** 对应 save_model_config：baseUrl / apiKey / maxTurns + 切换当前厂商模型 */
  saveModelConfig({ provider, model, apiKey, baseUrl, maxTurns }) {
    const cfg = this.current()
    const target = cfg.Providers[provider]
    if (!target) return { success: false, error: `未知的 provider: ${provider}` }

    if (baseUrl) {
      if (!/^https?:\/\//i.test(baseUrl)) {
        return { success: false, error: 'Base URL 格式无效，必须是 http:// 或 https:// 开头的完整 URL' }
      }
      target.BaseUrl = baseUrl
    }

    if (apiKey && apiKey !== '***keep***') {
      if (!this.credentials.set(provider, apiKey)) {
        return { success: false, error: 'API Key 加密保存失败（DPAPI 错误），请重试' }
      }
    }

    const switched = this.switchProvider(provider, model)
    if (!switched.success) return switched

    const turns = Number(maxTurns)
    if (turns > 0 && turns <= 50) cfg.General.MaxTurns = turns

    this.save()
    return { success: true }
  }
}

module.exports = ConfigStore
module.exports.MODEL_CATALOG = MODEL_CATALOG
module.exports.PROVIDER_DEFAULTS = PROVIDER_DEFAULTS
module.exports.CURRENT_MODEL_CATALOG_VERSION = CURRENT_MODEL_CATALOG_VERSION
