// src/DeepExcel.Wps/sidecar-host.js
// Python sidecar 子进程管理（对应 C# 端 PythonSidecar.cs）
// 职责：
// 1. child_process.spawn 启动 Python sidecar
// 2. stdin/stdout JSON Lines IPC（与 C# 端完全一致）
// 3. tool_call 消息路由到 tool-dispatcher.js
// 4. stream_delta / stream_end / clarify / permission_request 消息转发到 taskpane
// 5. 进程健康检查 + 异常退出通知
// 6. 环境变量 PYTHONUTF8=1（避免编码问题，与 C# 端一致）

const { spawn } = require('child_process')
const path = require('path')
const fs = require('fs')
const readline = require('readline')
const ToolDispatcher = require('./tool-dispatcher')

class SidecarHost {
  constructor(options = {}) {
    this.pythonPath = options.pythonPath || this._detectPython()
    this.sidecarPath = options.sidecarPath || this._detectSidecarPath()
    this.process = null
    this.exited = false
    this.stopping = false
    // ★ 工具调度器（对应 ToolDispatcher.cs）
    this.dispatcher = new ToolDispatcher()
    // ★ taskpane 消息监听器（main.js 注册，用于接收 sidecar 事件转发到前端）
    this.onEvent = null
  }

  // 启动 Python sidecar 子进程
  start() {
    if (this.process && !this.process.killed) {
      console.log('[SidecarHost] process already running, skip')
      return
    }
    this.exited = false
    this.stopping = false

    const args = [this.sidecarPath]
    const env = {
      ...process.env,
      // ★ 强制 Python 全局 UTF-8 模式（与 C# 端 PYTHONUTF8=1 一致）
      PYTHONUTF8: '1',
      PYTHONIOENCODING: 'utf-8',
    }

    this.process = spawn(this.pythonPath, args, {
      env,
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,  // 不显示控制台窗口
    })

    // ★ stdout：逐行读取 JSON Lines
    const rl = readline.createInterface({ input: this.process.stdout })
    rl.on('line', line => this._onStdoutLine(line))

    // ★ stderr：诊断日志，不转发到前端
    const errRl = readline.createInterface({ input: this.process.stderr })
    errRl.on('line', line => console.log(`[sidecar stderr] ${line}`))

    this.process.on('exit', (code, signal) => this._onExit(code, signal))
    this.process.on('error', err => this._onError(err))

    console.log(`[SidecarHost] Started, pid=${this.process.pid}, python=${this.pythonPath}, sidecar=${this.sidecarPath}`)
  }

  stop() {
    this.stopping = true
    if (this.process && !this.process.killed) {
      try { this._writeLine(JSON.stringify({ type: 'cancel' })) } catch (e) {}
      setTimeout(() => {
        if (this.process && !this.process.killed) {
          try { this.process.kill() } catch (e) {}
        }
      }, 2000)
    }
  }

  restart() {
    console.log('[SidecarHost] Restarting...')
    this.stop()
    setTimeout(() => {
      this.stopping = false
      this.start()
    }, 500)
  }

  // ============= 发送消息（JS → Python）=============
  // 协议与 C# 端完全一致：每行一个 JSON 对象

  sendUserMessage(text, sessionId, context) {
    this._writeLine(JSON.stringify({
      type: 'user_message', text, session_id: sessionId, context,
    }))
  }

  sendCancel() {
    this._writeLine(JSON.stringify({ type: 'cancel' }))
  }

  sendPermissionResponse(requestId, decision) {
    this._writeLine(JSON.stringify({
      type: 'permission_response', request_id: requestId, decision,
    }))
    console.log(`[SidecarHost] SendPermissionResponse: req_id=${requestId}, decision=${decision}`)
  }

  updateConfig(baseUrl, model, apiKey) {
    this._writeLine(JSON.stringify({
      type: 'config', base_url: baseUrl, model, api_key: apiKey,
    }))
  }

  sendClarifyAnswer(answer) {
    this._writeLine(JSON.stringify({ type: 'clarify_answer', answer }))
  }

  sendRestoreHistory(messages) {
    if (!messages || messages.length === 0) return
    this._writeLine(JSON.stringify({ type: 'restore_history', messages }))
  }

  // ★ 工具执行结果返回给 sidecar（tool-dispatcher 调用）
  sendToolResult(callId, success, data, error, suggestion, context) {
    this._writeLine(JSON.stringify({
      type: 'tool_result', call_id: callId, success, data, error, suggestion, context,
    }))
  }

  // ============= 接收消息（Python → JS）=============
  // 协议与 C# 端 SidecarProtocol.cs 一致

  _onStdoutLine(line) {
    if (!line) return
    let msg
    try { msg = JSON.parse(line) } catch (e) {
      console.error('[SidecarHost] parse stdout failed:', e.message, 'raw:', line.slice(0, 200))
      return
    }

    const type = msg.type
    console.log(`[SidecarHost] OnStdoutLine: type=${type}, len=${line.length}`)

    switch (type) {
      case 'stream_delta':
        this._emit({ type: 'stream_delta', payload: { delta: msg.text || '' } })
        break

      case 'tool_call':
        // ★ 路由到 tool-dispatcher 执行 WPS JS API 调用
        this._handleToolCall(msg)
        break

      case 'tool_use':
        this._emit({ type: 'tool_call', payload: { name: msg.tool, args: msg.args || {} } })
        break

      case 'clarify':
        this._emit({ type: 'clarify', payload: { question: msg.question, options: msg.options || [] } })
        break

      case 'stream_end':
        this._emit({ type: 'stream_end', payload: { input_tokens: msg.input_tokens || 0, output_tokens: msg.output_tokens || 0 } })
        break

      case 'permission_request':
        this._emit({
          type: 'permission_request',
          payload: { request_id: msg.request_id, tool: msg.tool, args: msg.args || {} },
        })
        break

      default:
        console.warn(`[SidecarHost] UNKNOWN type=${type}`)
    }
  }

  // ★ 工具调用：异步执行后返回结果
  async _handleToolCall(msg) {
    const callId = msg.call_id
    const toolName = msg.tool
    const args = msg.args || {}

    console.log(`[SidecarHost] HandleToolCall START: tool=${toolName}, call_id=${callId}`)

    try {
      const result = await this.dispatcher.execute(toolName, args)
      const context = this.dispatcher.buildExcelSnapshot()
      this.sendToolResult(callId, result.success, result.data, result.error, result.suggestion, context)
    } catch (err) {
      console.error(`[SidecarHost] HandleToolCall FAILED: ${toolName}`, err)
      this.sendToolResult(callId, false, {}, err.message, '', {})
    }
  }

  _onExit(code, signal) {
    if (this.exited) return
    this.exited = true
    console.log(`[SidecarHost] process exited, code=${code}, signal=${signal}, stopping=${this.stopping}`)
    // ★ 异常退出才通知前端（主动 stop 不发 error）
    if (!this.stopping) {
      this._emit({ type: 'error', payload: { message: `AI 助手进程异常退出 (code=${code})，请重新发送消息` } })
      this._emit({ type: 'stream_end', payload: {} })
    }
  }

  _onError(err) {
    console.error('[SidecarHost] process error:', err)
    this._emit({ type: 'error', payload: { message: `AI 助手启动失败: ${err.message}` } })
  }

  _writeLine(text) {
    if (!this.process || !this.process.stdin || this.process.killed) {
      console.warn('[SidecarHost] stdin not writable, dropping message:', text.slice(0, 100))
      return
    }
    try {
      this.process.stdin.write(text + '\n')
    } catch (e) {
      console.error('[SidecarHost] write stdin failed:', e)
    }
  }

  // ★ 转发事件到 main.js（main.js 再 postMessage 到 taskpane）
  _emit(event) {
    if (typeof this.onEvent === 'function') {
      try { this.onEvent(event) } catch (e) { console.error('[SidecarHost] onEvent failed:', e) }
    }
  }

  _detectPython() {
    // 优先使用项目内置的 python embeddable
    const projectRoot = path.resolve(__dirname, '..', '..')
    const candidates = [
      path.join(projectRoot, 'python-3.11-embed-amd64', 'python.exe'),
      path.join(projectRoot, 'python', 'python.exe'),
      'python',
      'python3',
    ]
    for (const p of candidates) {
      if (p.startsWith('python') || fs.existsSync(p)) return p
    }
    return 'python'
  }

  _detectSidecarPath() {
    const projectRoot = path.resolve(__dirname, '..', '..')
    return path.join(projectRoot, 'src', 'DeepExcel.Sidecar', 'sidecar.py')
  }
}

module.exports = SidecarHost
