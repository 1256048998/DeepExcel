/**
 * WebView2 桥接层
 * 在浏览器环境中（Vite dev）提供 mock 实现
 * 在WebView2中通过window.chrome.webview与C#通信
 */

export type HostMessage = {
  type: string
  payload: any
}

// 检测是否在WebView2中
const isInWebView = typeof (window as any).chrome !== 'undefined' &&
  (window as any).chrome.webview

// 消息监听器
let listeners: ((msg: HostMessage) => void)[] = []
// ★ 防止重复注册 webview message listener（原 bug：每次调用 onHostMessage 都注册一次，
//   导致 N 个监听器 × N 次注册 = N² 倍消息分发）
let webviewListenerInitialized = false

function ensureWebviewListener() {
  if (webviewListenerInitialized) return
  webviewListenerInitialized = true

  if (isInWebView) {
    // ★ WebView2 中 C# 用 PostWebMessageAsString 发送消息，
    // 必须用 chrome.webview.addEventListener('message', ...) 监听（不是 window.addEventListener）。
    // PostWebMessageAsString 发送的是字符串，event.data 就是原始字符串（不会自动 JSON.parse）。
    // C# MessageBridge.SendToUi 发送的格式是 JSON 字符串：{"type":"stream_delta","payload":{"delta":"..."}}
    const webview = (window as any).chrome.webview
    webview.addEventListener('message', (e: MessageEvent) => {
      try {
        const raw = e.data
        const data = typeof raw === 'string' ? JSON.parse(raw) : raw
        // 分发给所有监听器（复制数组避免遍历中被修改）
        const snapshot = [...listeners]
        snapshot.forEach(l => {
          try { l(data) } catch (err) { console.error('Listener error:', err) }
        })
      } catch (err) {
        console.error('Parse host message error:', err, 'raw:', e.data)
      }
    })
  }
}

export async function sendToHost(message: HostMessage): Promise<void> {
  ensureWebviewListener()
  if (isInWebView) {
    ;(window as any).chrome.webview.postMessage(message)
  } else {
    // 开发环境：模拟C#响应
    console.log('[Bridge→Host]', message)
    setTimeout(() => mockHostResponse(message), 500)
  }
}

/**
 * ★ 请求-响应模式：发送消息并等待指定类型的响应。
 * C# 端 HandleMessage 是同步返回的（WebMessageReceived 事件中立即 PostWebMessageAsString），
 * 但前端 postMessage 是异步的，响应通过 onHostMessage 全局监听器接收。
 * 此函数注册临时监听器匹配 expectedType，匹配后自动移除。
 */
export async function sendToHostWithResponse(
  message: HostMessage,
  expectedType: string,
  timeout = 5000
): Promise<HostMessage | null> {
  ensureWebviewListener()
  return new Promise((resolve) => {
    let resolved = false
    const handler = (msg: HostMessage) => {
      if (!resolved && msg.type === expectedType) {
        resolved = true
        clearTimeout(timer)
        removeHandler()
        resolve(msg)
      } else if (!resolved && msg.type === 'error') {
        // 错误响应也算匹配，返回给调用方处理
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
  ensureWebviewListener()
  listeners.push(callback)

  if (!isInWebView) {
    // 开发环境：连接ok信号
    setTimeout(() => callback({ type: 'connection_ok', payload: {} }), 100)
  }

  return () => {
    const idx = listeners.indexOf(callback)
    if (idx >= 0) listeners.splice(idx, 1)
  }
}

// 开发环境模拟响应
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
    const response = `收到你的需求："${content}"\n\n[开发模式 - 模拟响应]\n在生产环境中，这里会通过桥接层调用C# → 感知Excel → 调用AI模型 → 生成VBA → 在Excel中执行。`

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
