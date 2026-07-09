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

/** WPS taskpane 环境（taskpane 在 iframe 中，parent 是 main.js 的窗口） */
const isInWpsTaskpane = !isInWebView &&
  typeof window !== 'undefined' &&
  window.parent !== window &&
  typeof (window as any).wps === 'undefined'  // taskpane 内部没有 wps 全局对象

/** 开发环境（Vite dev server，无宿主） */
const isDev = !isInWebView && !isInWpsTaskpane

// ★ 防止重复注册 listener
let listenerInitialized = false

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
    // ★ WPS taskpane: window.addEventListener('message', ...)
    // main.js 通过 taskpane.postMessage(json) 发送，前端用 window.message 事件接收
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
    // ★ WPS taskpane: window.parent.postMessage
    // main.js 的 _setupTaskpaneMessageListener 接收
    window.parent.postMessage(message, '*')
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

function mockHostResponse(message: HostMessage) {
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
