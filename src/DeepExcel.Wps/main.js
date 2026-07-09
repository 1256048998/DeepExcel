// src/DeepExcel.Wps/main.js
// WPS JS 加载项入口（对应 C# 端 ThisAddIn.cs + TaskPaneControl.cs）
// 职责：
// 1. Ribbon 按钮回调：打开 taskpane 承载 React 前端
// 2. 启动 Python sidecar 子进程
// 3. taskpane ↔ sidecar 消息桥接（postMessage ↔ stdin/stdout）
// 4. 管理生命周期（加载/卸载）

const SidecarHost = require('./sidecar-host')

// ★ 全局 sidecar 实例（单例，整个加载项生命周期内复用）
let sidecar = null
// ★ taskpane 引用（WPS ShowTaskPane 返回的对象）
let taskpane = null

// ============= WPS 加载项生命周期回调 =============

/**
 * 加载项初始化（WPS 加载加载项时调用）
 */
function OnPluginInit() {
  console.log('[DeepExcel] OnPluginInit: starting sidecar...')
  try {
    sidecar = new SidecarHost()
    // ★ 注册事件转发器：sidecar 事件 → taskpane
    sidecar.onEvent = (event) => {
      _forwardToTaskpane(event)
    }
    sidecar.start()
    console.log('[DeepExcel] OnPluginInit: sidecar started')
  } catch (err) {
    console.error('[DeepExcel] OnPluginInit FAILED:', err)
  }
}

/**
 * 加载项卸载
 */
function OnPluginDestroy() {
  console.log('[DeepExcel] OnPluginDestroy')
  if (sidecar) {
    sidecar.stop()
    sidecar = null
  }
}

// ============= Ribbon 回调 =============

/**
 * Ribbon onLoad
 */
function OnRibbonLoad(ribbonUI) {
  console.log('[DeepExcel] OnRibbonLoad')
}

/**
 * "打开面板"按钮回调
 * ★ 仅打开/显示面板，不 toggle（与 Excel 端行为一致，符合用户偏好）
 */
function OnAction(control) {
  console.log('[DeepExcel] OnAction: opening taskpane...')
  try {
    // ★ WPS ShowTaskPane API
    // 如果 taskpane 已存在，仅显示；否则创建
    if (!taskpane) {
      taskpane = wps.ShowTaskPane({
        url: _getTaskpaneUrl(),
        title: 'DeepExcel AI',
        width: 400,
        // ★ 右侧停靠
        dockRight: true,
      })
      // ★ 监听 taskpane 的 postMessage（来自 React 前端的消息）
      _setupTaskpaneMessageListener()
    } else {
      taskpane.Visible = true
    }
  } catch (err) {
    console.error('[DeepExcel] OnAction FAILED:', err)
  }
}

/**
 * "帮助"按钮回调
 */
function OnShowHelp(control) {
  console.log('[DeepExcel] OnShowHelp')
  try {
    wps.ShowHelp('https://github.com/yourusername/DeepExcel')
  } catch (e) {
    // WPS 可能不支持 ShowHelp，用 ShowDialog 替代
    wps.ShowDialog({
      url: 'data:text/html,<html><body><h1>DeepExcel 使用帮助</h1><p>1. 点击"打开面板"按钮启动 AI 对话</p><p>2. 在面板中输入你的 Excel 任务</p><p>3. AI 会自动调用工具操作 WPS 表格</p></body></html>',
      title: '使用帮助',
      width: 400,
      height: 300,
    })
  }
}

// ============= taskpane ↔ sidecar 消息桥接 =============

/**
 * 监听 taskpane 中 React 前端的 postMessage
 * 前端通过 window.parent.postMessage 发送消息到 main.js
 */
function _setupTaskpaneMessageListener() {
  // ★ WPS JS 加载项的消息监听机制
  // 具体实现取决于 WPS 版本，这里用通用方式
  // 如果 WPS 提供了 taskpane.OnMessage API，用那个；否则用 window.addEventListener
  try {
    if (taskpane && typeof taskpane.OnMessage === 'function') {
      taskpane.OnMessage((msg) => {
        _handleFrontendMessage(typeof msg === 'string' ? JSON.parse(msg) : msg)
      })
    }
  } catch (e) {
    console.warn('[DeepExcel] taskpane.OnMessage setup failed, falling back:', e.message)
  }
}

/**
 * 处理来自 React 前端的消息
 * 对应 C# 端 MessageBridge.HandleMessage
 */
function _handleFrontendMessage(msg) {
  if (!msg || !msg.type) return
  console.log(`[DeepExcel] Frontend→Host: type=${msg.type}`)

  if (!sidecar) {
    console.warn('[DeepExcel] sidecar not running, cannot handle frontend message')
    return
  }

  switch (msg.type) {
    case 'user_message':
      // ★ 用户发送消息：转发给 sidecar
      sidecar.sendUserMessage(msg.payload.content, _getSessionId(), _buildContext())
      break

    case 'cancel':
      sidecar.sendCancel()
      break

    case 'permission_response':
      sidecar.sendPermissionResponse(msg.payload.request_id, msg.payload.decision)
      break

    case 'new_conversation':
      // ★ 新建对话：重启 sidecar 清空上下文
      sidecar.restart()
      break

    case 'list_conversations':
      // ★ 列出历史对话（WPS 端暂用本地存储）
      _listConversations()
      break

    case 'continue_conversation':
      // ★ 继续历史对话
      _continueConversation(msg.payload.conversation_id)
      break

    case 'list_attachments':
      // ★ 列出附件
      _listAttachments()
      break

    case 'upload_attachment':
      // ★ 上传附件
      _uploadAttachment(msg.payload.file_name, msg.payload.file_base64)
      break

    case 'delete_attachment':
      _deleteAttachment(msg.payload.file_name)
      break

    case 'clarify_answer':
      sidecar.sendClarifyAnswer(msg.payload.answer)
      break

    default:
      console.warn(`[DeepExcel] Unknown frontend message type: ${msg.type}`)
  }
}

/**
 * 转发 sidecar 事件到 taskpane（前端）
 * 对应 C# 端 PostWebMessageAsString
 */
function _forwardToTaskpane(event) {
  if (!taskpane) {
    console.warn('[DeepExcel] taskpane not available, cannot forward event:', event.type)
    return
  }
  try {
    const json = JSON.stringify(event)
    // ★ WPS taskpane 的消息发送 API
    if (typeof taskpane.postMessage === 'function') {
      taskpane.postMessage(json)
    } else if (typeof taskpane.SendToWeb === 'function') {
      taskpane.SendToWeb(json)
    } else {
      console.warn('[DeepExcel] taskpane does not support postMessage/SendToWeb')
    }
  } catch (e) {
    console.error('[DeepExcel] _forwardToTaskpane failed:', e)
  }
}

// ============= 辅助函数 =============

function _getTaskpaneUrl() {
  // ★ taskpane.html 的绝对路径
  const path = require('path')
  const htmlPath = path.join(__dirname, 'taskpane.html')
  // WPS ShowTaskPane 接受 file:// URL 或绝对路径
  return 'file:///' + htmlPath.replace(/\\/g, '/')
}

function _getSessionId() {
  // ★ 会话 ID：用时间戳生成（简化版，实际可用 UUID）
  return 'wps-' + Date.now()
}

function _buildContext() {
  // ★ 上下文信息（对应 C# 端 BuildExcelSnapshot）
  try {
    const app = wps.Application
    const wb = app.ActiveWorkbook
    return {
      host_type: 'wps',
      workbook: wb ? wb.Name : '',
      path: wb ? wb.FullName : '',
      activeSheet: (wb && wb.ActiveSheet) ? wb.ActiveSheet.Name : '',
    }
  } catch (e) {
    return { host_type: 'wps', error: e.message }
  }
}

// ============= 历史对话和附件（简化实现，后续可扩展） =============

const conversations = []
const attachments = []

function _listConversations() {
  _forwardToTaskpane({
    type: 'conversations',
    payload: { list: conversations },
  })
}

function _continueConversation(conversationId) {
  const conv = conversations.find(c => c.id === conversationId)
  if (conv) {
    sidecar.restart()
    sidecar.sendRestoreHistory(conv.messages)
    _forwardToTaskpane({
      type: 'continue_conversation',
      payload: { messages: conv.messages },
    })
  }
}

function _listAttachments() {
  _forwardToTaskpane({
    type: 'attachments',
    payload: { list: attachments },
  })
}

function _uploadAttachment(fileName, fileBase64) {
  attachments.push({ fileName, size: Math.floor(fileBase64.length * 0.75) })
  _forwardToTaskpane({ type: 'uploaded', payload: { success: true, fileName } })
}

function _deleteAttachment(fileName) {
  const idx = attachments.findIndex(a => a.fileName === fileName)
  if (idx >= 0) attachments.splice(idx, 1)
  _forwardToTaskpane({ type: 'attachment_deleted', payload: { fileName } })
}

// ============= 导出（WPS 加载项入口） =============

module.exports = {
  OnPluginInit,
  OnPluginDestroy,
  OnRibbonLoad,
  OnAction,
  OnShowHelp,
}

// ★ 如果 WPS 直接 require main.js 并调用导出函数，上面的 module.exports 生效
// ★ 如果 WPS 期望全局函数，则挂到 global
if (typeof global !== 'undefined') {
  global.OnPluginInit = OnPluginInit
  global.OnPluginDestroy = OnPluginDestroy
  global.OnRibbonLoad = OnRibbonLoad
  global.OnAction = OnAction
  global.OnShowHelp = OnShowHelp
}
