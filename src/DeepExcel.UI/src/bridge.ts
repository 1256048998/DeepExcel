/**
 * 桥接层（多宿主支持）
 *
 * 支持两种宿主环境：
 * 1. Excel WebView2（通过 window.chrome.webview 与 C# 通信）
 * 2. WPS taskpane（通过 window.parent.postMessage 与 JS 加载项通信）
 *
 * API 保持不变（sendToHost / onHostMessage / sendToHostWithResponse），
 * 内部按环境自动选择实现。
 */

export type HostMessage = {
  type: string
  payload: any
}

// ★ 消息监听器（两种宿主共用）
let listeners: ((msg: HostMessage) => void)[] = []

// ============= 环境检测 =============

/** Excel WebView2 环境 */
const isInWebView = typeof (window as any).chrome !== 'undefined' &&
  (window as any).chrome.webview

/** WPS taskpane environment. Current WPS exposes Application inside panes. */
const isInWpsTaskpane = !isInWebView &&
  typeof window !== 'undefined' &&
  (typeof (window as any).Application !== 'undefined' ||
   typeof (window as any).wps !== 'undefined' ||
   window.parent !== window)

/** 开发环境（Vite dev server，无宿主） */
const isDev = !isInWebView && !isInWpsTaskpane

// ★ 防止重复注册 listener
let listenerInitialized = false
let wpsChannel: BroadcastChannel | null = null

function ensureListener() {
  if (listenerInitialized) return
  listenerInitialized = true

  if (isInWebView) {
    // ★ WebView2: chrome.webview.addEventListener('message', ...)
    const webview = (window as any).chrome.webview
    webview.addEventListener('message', (e: MessageEvent) => {
      try {
        const raw = e.data
        const data = typeof raw === 'string' ? JSON.parse(raw) : raw
        _dispatch(data)
      } catch (err) {
        console.error('Parse host message error:', err, 'raw:', e.data)
      }
    })
  } else if (isInWpsTaskpane) {
    // WPS task panes are separate browser surfaces, not guaranteed iframes.
    // BroadcastChannel provides a same-origin bridge to the hidden add-in page.
    if (typeof BroadcastChannel === 'function') {
      wpsChannel = new BroadcastChannel('deepexcel-wps')
      wpsChannel.addEventListener('message', (e: MessageEvent) => {
        try {
          const raw = e.data
          _dispatch(typeof raw === 'string' ? JSON.parse(raw) : raw)
        } catch (err) {
          console.error('Parse WPS channel message error:', err, 'raw:', e.data)
        }
      })
    }
    // Compatibility fallback for WPS builds that embed the pane as an iframe.
    window.addEventListener('message', (e: MessageEvent) => {
      try {
        const raw = e.data
        const data = typeof raw === 'string' ? JSON.parse(raw) : raw
        _dispatch(data)
      } catch (err) {
        console.error('Parse WPS host message error:', err, 'raw:', e.data)
      }
    })
  }
}

/** 分发消息到所有监听器 */
function _dispatch(data: HostMessage) {
  const snapshot = [...listeners]
  snapshot.forEach(l => {
    try { l(data) } catch (err) { console.error('Listener error:', err) }
  })
}

// ============= 公共 API（保持与原 bridge.ts 兼容） =============

export async function sendToHost(message: HostMessage): Promise<void> {
  ensureListener()
  if (isInWebView) {
    // ★ Excel WebView2: chrome.webview.postMessage
    ;(window as any).chrome.webview.postMessage(message)
  } else if (isInWpsTaskpane) {
    if (wpsChannel) {
      wpsChannel.postMessage(message)
    } else if (window.parent !== window) {
      window.parent.postMessage(message, '*')
    } else {
      _dispatch({ type: 'error', payload: { message: 'WPS 消息桥接不可用，请升级 WPS 后重试。' } })
    }
  } else {
    // ★ 开发环境：模拟响应
    console.log('[Bridge→Host]', message)
    setTimeout(() => mockHostResponse(message), 500)
  }
}

/**
 * ★ 请求-响应模式：发送消息并等待指定类型的响应。
 * 协议与 C# 端一致：注册临时监听器匹配 expectedType，匹配后自动移除。
 */
export async function sendToHostWithResponse(
  message: HostMessage,
  expectedType: string,
  timeout = 5000
): Promise<HostMessage | null> {
  ensureListener()
  return new Promise((resolve) => {
    let resolved = false
    const handler = (msg: HostMessage) => {
      if (!resolved && msg.type === expectedType) {
        resolved = true
        clearTimeout(timer)
        removeHandler()
        resolve(msg)
      } else if (!resolved && msg.type === 'error') {
        resolved = true
        clearTimeout(timer)
        removeHandler()
        resolve(msg)
      }
    }
    function removeHandler() {
      const idx = listeners.indexOf(handler)
      if (idx >= 0) listeners.splice(idx, 1)
    }
    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true
        removeHandler()
        console.warn(`[Bridge] sendToHostWithResponse timeout: expectedType=${expectedType}`)
        resolve(null)
      }
    }, timeout)

    listeners.push(handler)
    sendToHost(message)
  })
}

export function onHostMessage(callback: (msg: HostMessage) => void): () => void {
  ensureListener()
  listeners.push(callback)

  if (isDev) {
    // 开发环境：连接ok信号
    setTimeout(() => callback({ type: 'connection_ok', payload: {} }), 100)
  }

  return () => {
    const idx = listeners.indexOf(callback)
    if (idx >= 0) listeners.splice(idx, 1)
  }
}

// ============= 环境信息导出（供前端判断宿主类型） =============

export const hostType: 'excel' | 'wps' | 'dev' = isInWebView ? 'excel' : isInWpsTaskpane ? 'wps' : 'dev'

// ============= 开发环境模拟响应 =============

/**
 * ★ 开发环境的场景开关：?mock=<scenario>
 *
 * 面板的不少状态在真机上很难复现——"一个 Key 都没配"只在全新安装的那几分钟
 * 里存在，"托管模式"需要一个配好的服务端。结果就是这些分支只能靠脑补验证，
 * 而其中一条（托管用户被 setupNeeded 误拦）真的漏出去过。
 *
 *   ?mock=nokey    一个 Key 都没配、未登录 -> 应当出现"先配置一个供应商"
 *   ?mock=hosted   一个 Key 都没配、托管模式 -> **不应**出现（本地无 Key 是正常的）
 *   其他/不带      已配 DeepSeek、未登录 -> 正常可用
 */
const mockScenario: string = isDev
  ? new URLSearchParams(window.location.search).get('mock') || 'default'
  : 'default'

// ★ 开发环境的模型配置假数据：让"模型配置"弹窗（模型优先级 / 导入模型）在 vite dev 下可调试
const mockProviders: Record<string, any> = {
  deepseek: {
    displayName: 'DeepSeek', type: 'anthropic', baseUrl: 'https://api.deepseek.com/anthropic',
    defaultModel: 'deepseek-v4-pro', supportsVision: false,
    models: ['deepseek-v4-pro', 'deepseek-v4-flash', 'deepseek-flash'],
    hasApiKey: true, apiKeyPreview: 'sk-5***...cee', connected: true
  },
  anthropic: {
    displayName: 'Claude (Anthropic)', type: 'anthropic', baseUrl: 'https://api.anthropic.com',
    defaultModel: 'claude-sonnet-5', supportsVision: true,
    models: ['claude-sonnet-5', 'claude-opus-5', 'claude-opus-4.8', 'claude-haiku-5', 'claude-haiku-4-5-20251001'],
    hasApiKey: false, apiKeyPreview: '', connected: false
  }
}

function mockHostResponse(message: HostMessage) {
  const emit = (type: string, payload: any) => listeners.forEach(l => l({ type, payload }))

  switch (message.type) {
    case 'account_status':
      // 没有这个 mock 的时候，App 启动时那次 account_status 会一直等到超时，
      // 而 accountStatus 停在 null 就意味着 setupNeeded 永远不成立——dev 下
      // 根本看不到那条引导。
      emit('account_status', mockScenario === 'hosted'
        ? { state: 'signedin', server_url: 'https://mock.local', email: 'dev@example.com',
            mode: 'hosted', entitlement: { plan: 'pro', status: 'active', routing_mode: 'hosted',
              task_limit: 1000, tasks_used: 12, tasks_remaining: 988, expires_at: null } }
        : { state: 'signedout', server_url: null, email: null, mode: null, entitlement: null })
      return
    case 'get_model_config':
      emit('model_config', {
        currentProvider: 'deepseek', currentModel: 'deepseek-v4-pro', defaultProvider: 'deepseek',
        // nokey / hosted 两个场景都是"本地一个 Key 都没有"，差别只在路由模式。
        providers: (mockScenario === 'nokey' || mockScenario === 'hosted')
          ? Object.fromEntries(Object.entries(mockProviders).map(
              ([k, v]) => [k, { ...v, hasApiKey: false, connected: false, apiKeyPreview: '' }]))
          : mockProviders,
        general: { maxRetries: 2, requestTimeoutSeconds: 60, autoCreateSnapshot: true, requireConfirmation: true, maxConversationHistory: 10, maxTurns: 20 },
        ui: { theme: 'light', language: 'zh-CN', showTokenUsage: true, streamOutput: true }
      })
      return
    case 'set_provider_models': {
      const { provider, models } = message.payload
      if (mockProviders[provider]) {
        mockProviders[provider].models = models
        mockProviders[provider].defaultModel = models[0]
      }
      emit('provider_models_saved', { success: true, models, defaultModel: models[0], currentModel: models[0] })
      return
    }
    case 'refresh_models':
      emit('models_refreshed', {
        success: true,
        models: ['deepseek-v4-pro', 'deepseek-v4-flash', 'deepseek-flash', 'deepseek-v3.2', 'deepseek-reasoner', 'deepseek-chat'],
        selected: mockProviders[message.payload.provider]?.models || [],
        currentModel: 'deepseek-v4-pro'
      })
      return
    case 'test_api_key':
      emit('api_test_result', { success: true, latencyMs: 420, error: null })
      return
    case 'get_api_key':
      emit('api_key', { apiKey: 'sk-5xxxxxxxxxxxxxxxxxxxxxcee' })
      return
    case 'save_model_config':
      emit('config_saved', { success: true })
      return
  }

  if (message.type === 'user_message') {
    const content = message.payload.content

    // 模拟 clarify 响应
    if (content.includes('统计') || content.includes('clarify')) {
      setTimeout(() => {
        listeners.forEach(l => l({
          type: 'clarify',
          payload: {
            question: '检测到A列同时包含数字和文本，请问您要统计什么？',
            options: ['SUM（仅数字求和）', 'COUNTA（非空单元格计数）', 'COUNT（数字单元格计数）']
          }
        }))
      }, 500)
      return
    }

    // 模拟流式输出
    const response = `收到你的需求："${content}"\n\n[开发模式 - 模拟响应]\n在生产环境中，这里会通过桥接层调用宿主 → 感知表格 → 调用AI模型 → 生成工具调用 → 执行操作。`

    let i = 0
    const interval = setInterval(() => {
      if (i < response.length) {
        const delta = response[i]
        i++
        listeners.forEach(l => l({ type: 'stream_delta', payload: { delta } }))
      } else {
        clearInterval(interval)
        listeners.forEach(l => l({ type: 'stream_end', payload: {} }))
      }
    }, 20)
  }
}
