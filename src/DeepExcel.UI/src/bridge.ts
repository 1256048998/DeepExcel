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

export async function sendToHost(message: HostMessage): Promise<void> {
  if (isInWebView) {
    ;(window as any).chrome.webview.postMessage(message)
  } else {
    // 开发环境：模拟C#响应
    console.log('[Bridge→Host]', message)
    setTimeout(() => mockHostResponse(message), 500)
  }
}

export function onHostMessage(callback: (msg: HostMessage) => void): () => void {
  listeners.push(callback)

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
        // 分发给所有监听器
        listeners.forEach(l => l(data))
      } catch (err) {
        console.error('Parse host message error:', err, 'raw:', e.data)
      }
    })
  } else {
    // 开发环境：连接ok信号
    setTimeout(() => callback({ type: 'connection_ok', payload: {} }), 100)
  }

  return () => {
    listeners = listeners.filter(l => l !== callback)
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
