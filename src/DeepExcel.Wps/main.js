// DeepExcel WPS JS add-in entry. Keep this file browser-safe.

// ★ 诊断日志。WPS 端曾经在用户机器上"装了但用不了"：选项卡在，点"打开面板"
// 毫无反应，而且不留任何痕迹——因为出错分支用 app.ShowDialog('data:text/html;...')
// 报错，WPS 若拒绝 data: URL，这个兜底也失败，用户和我们都什么都看不到。
// 干净沙箱 + 真实 WPS 已复现。任何排查的前提是先有一行日志。
//
// 运行时能力不确定（require 在 WPS 的 JS 宿主里未必存在），所以逐个降级尝试，
// 全部失败也绝不抛异常——诊断代码自己把加载项搞挂是最坏的结果。
var DIAG_PATH_HINT = 'DeepExcel\\logs\\wps-panel.log'
function _diag(message) {
  var line = '[' + new Date().toISOString() + '] ' + message + '\n'
  try {
    var fs = require('fs')
    var path = require('path')
    var base = (typeof process !== 'undefined' && process.env &&
                (process.env.LOCALAPPDATA || process.env.APPDATA)) || 'C:\\'
    var dir = path.join(base, 'DeepExcel', 'logs')
    try { fs.mkdirSync(dir, { recursive: true }) } catch (_) {}
    fs.appendFileSync(path.join(dir, 'wps-panel.log'), line)
    return
  } catch (_) {}
  try { console.log('[DeepExcel] ' + message); return } catch (_) {}
}

var sidecar = null
var taskpane = null
var messageChannel = null
// ★ 模型配置（厂商 / 模型优先级 / API Key）：与 Excel 加载项共用
// %APPDATA%\DeepExcel\config.json 和 credentials\*.crypt，两端配置一次即可。
var configStore = null
var modelService = null
var KEEP_API_KEY = '***keep***'
// ★ 对话历史 + 附件：按工作簿隔离的会话状态，历史文件与 Excel 加载项共用
// %LOCALAPPDATA%\DeepExcel\history\，同一个工作簿在两个宿主里能接着聊。
var conversationModule = null
var AttachmentStore = null
var conversationHistory = null
var sessions = {}

function GetUrlPath() {
  var url = decodeURI(document.location.toString())
  var slash = url.lastIndexOf('/')
  return slash >= 0 ? url.substring(0, slash) : url
}

function _application() {
  return window.Application || (window.wps && window.wps.Application) || window.wps
}

function _ensureMessageBridge() {
  if (!messageChannel && typeof BroadcastChannel === 'function') {
    messageChannel = new BroadcastChannel('deepexcel-wps')
    messageChannel.onmessage = function (event) { _handleFrontendMessage(event.data) }
  }
  if (!window.__deepExcelMessageListenerInstalled) {
    window.__deepExcelMessageListenerInstalled = true
    window.addEventListener('message', function (event) { _handleFrontendMessage(event.data) })
  }
}

function _ensureSidecar() {
  if (sidecar) return true
  if (typeof require !== 'function') {
    console.error('[DeepExcel] WPS runtime does not expose Node require; sidecar unavailable')
    return false
  }
  try {
    var SidecarHost = require('./sidecar-host')
    sidecar = new SidecarHost()
    sidecar.onEvent = function (event) {
      // ★ 先记进对话历史，再转发给前端：stream_end 时整段对话落盘
      _recordSidecarEvent(event)
      _forwardToTaskpane(event)
    }
    sidecar.start()
    // ★ 启动后立刻下发模型配置，否则 sidecar 拿不到 base_url / model / api_key
    if (_ensureConfigStore()) {
      configStore.credentials.setPythonPath(sidecar.pythonPath)
      _pushConfigToSidecar()
    }
    return true
  } catch (error) {
    console.error('[DeepExcel] sidecar start failed:', error)
    sidecar = null
    return false
  }
}

// ★ 懒加载模型配置模块。WPS 若不提供 Node require，则退化为"面板可用但配置不可用"，
// 而不是整个面板报错——所以这里独立于 _ensureSidecar 做判断。
function _ensureConfigStore() {
  if (configStore) return true
  if (typeof require !== 'function') return false
  try {
    var ConfigStore = require('./config-store')
    modelService = require('./model-service')
    configStore = new ConfigStore({ pythonPath: sidecar && sidecar.pythonPath })
    configStore.load()
    return true
  } catch (error) {
    console.error('[DeepExcel] config store init failed:', error)
    configStore = null
    modelService = null
    return false
  }
}

// ★ 懒加载对话历史 / 附件模块（同样容忍无 require 的运行环境）
function _ensureSessionStores() {
  if (conversationModule && AttachmentStore) return true
  if (typeof require !== 'function') return false
  try {
    conversationModule = require('./conversation-store')
    AttachmentStore = require('./attachment-store')
    conversationHistory = new conversationModule.ConversationHistory()
    return true
  } catch (error) {
    console.error('[DeepExcel] session stores init failed:', error)
    conversationModule = null
    AttachmentStore = null
    return false
  }
}

// ★ 工作簿标识：与 C# GetWorkbookKey 同口径（优先 FullName，其次 Name），
// 这样同一个文件在 Excel 和 WPS 里命中同一份历史。
// 缓存 2 秒：流式回复期间每个 stream_delta 都会查一次会话，
// 不缓存的话等于每个字符都走一次 COM 调用。
var _workbookKeyCache = { key: null, name: '', at: 0 }
var WORKBOOK_KEY_TTL_MS = 2000

function _readWorkbookIdentity() {
  try {
    var app = _application()
    var wb = app && app.ActiveWorkbook
    if (!wb) return { key: 'workbook_unknown', name: '' }
    var fullName = wb.FullName
    var key = (fullName && (fullName.indexOf('\\') >= 0 || fullName.indexOf('/') >= 0))
      ? fullName
      : (wb.Name || 'workbook_unknown')
    return { key: key, name: wb.Name || '' }
  } catch (error) {
    return { key: 'workbook_unknown', name: '' }
  }
}

function _workbookIdentity() {
  var now = Date.now()
  if (_workbookKeyCache.key && now - _workbookKeyCache.at < WORKBOOK_KEY_TTL_MS) {
    return _workbookKeyCache
  }
  var identity = _readWorkbookIdentity()
  _workbookKeyCache = { key: identity.key, name: identity.name, at: now }
  return _workbookKeyCache
}

/** 取当前活动工作簿的会话（对话历史 + 附件），没有就建一个 */
function _session() {
  if (!_ensureSessionStores()) return null
  var identity = _workbookIdentity()
  var key = identity.key
  if (!sessions[key]) {
    sessions[key] = {
      key: key,
      conversation: new conversationModule.ConversationSession(key, identity.name, conversationHistory),
      attachments: new AttachmentStore(key),
    }
  }
  sessions[key].conversation.setWorkbookName(identity.name)
  return sessions[key]
}

// ★ 把 sidecar 事件记进当前对话（对应 C# MessageBridge 里对 session 的 Append*）
function _recordSidecarEvent(event) {
  var session = _session()
  if (!session || !event || !event.type) return
  var payload = event.payload || {}
  try {
    switch (event.type) {
      case 'stream_delta':
        if (payload.delta) session.conversation.appendAssistantDelta(payload.delta)
        break
      case 'tool_call':
        // 计划清单由侧车自己处理，不是工作簿操作，不进对话历史
        if (payload.name && !/todo_write$/.test(payload.name)) session.conversation.appendToolCall(payload.name)
        break
      case 'clarify':
        session.conversation.appendClarify(payload.question || '', payload.options || [])
        break
      case 'stream_end':
        session.conversation.onStreamEnd()
        break
      default:
        break
    }
  } catch (error) {
    console.error('[DeepExcel] record conversation failed:', error)
  }
}

// ★ 把当前 provider/model/apiKey 下发给 sidecar（对应 C# SendConfigToSession）。
// 之前 WPS 端从来没发过 config，sidecar 只能用默认值，这是配置面板落地的必要一环。
function _pushConfigToSidecar() {
  if (!configStore || !sidecar) return
  try {
    var conf = configStore.getSidecarConfig()
    if (!conf) return
    sidecar.updateConfig(conf.baseUrl, conf.model, conf.apiKey)
  } catch (error) {
    console.error('[DeepExcel] push config to sidecar failed:', error)
  }
}

function OnRibbonLoad(ribbonUI) {
  var app = _application()
  if (app && typeof app.ribbonUI !== 'object') app.ribbonUI = ribbonUI
  _ensureMessageBridge()
  _ensureSidecar()
  return true
}

function OnPluginInit() {
  _ensureMessageBridge()
  _ensureSidecar()
  return true
}

function OnPluginDestroy() {
  // ★ 退出前把还没落盘的对话存起来（没走到 stream_end 就关 WPS 的情况）
  try {
    for (var key in sessions) {
      if (!Object.prototype.hasOwnProperty.call(sessions, key)) continue
      var conversation = sessions[key].conversation
      if (conversation.messages.length > 0 && conversation.currentConversationId) {
        conversation.saveCurrent()
      }
    }
  } catch (error) {
    console.error('[DeepExcel] save conversations on destroy failed:', error)
  }
  if (sidecar) {
    sidecar.stop()
    sidecar = null
  }
  if (messageChannel) {
    messageChannel.close()
    messageChannel = null
  }
}

function OnAction() {
  _diag('OnAction entered')
  var app = _application()
  if (!app) {
    _diag('FATAL: WPS Application API unavailable (window.Application / wps both missing)')
    console.error('[DeepExcel] WPS Application API unavailable')
    return false
  }
  _diag('Application API present; typeof CreateTaskPane=' + (typeof app.CreateTaskPane))

  try {
    var storedId = app.PluginStorage && app.PluginStorage.getItem('deepexcel_taskpane_id')
    if (storedId) {
      try { taskpane = app.GetTaskPane(storedId) } catch (_) { taskpane = null }
    }
    if (!taskpane) {
      var paneUrl = GetUrlPath() + '/taskpane.html'
      _diag('creating taskpane at: ' + paneUrl)
      taskpane = app.CreateTaskPane(paneUrl, 'DeepExcel AI')
      if (!taskpane) throw new Error('CreateTaskPane returned undefined (URL rejected or unsupported)')
      _diag('taskpane created, id=' + taskpane.ID)
      if (app.PluginStorage) app.PluginStorage.setItem('deepexcel_taskpane_id', taskpane.ID)
      try {
        var right = app.Enum && app.Enum.msoCTPDockPositionRight
        taskpane.DockPosition = right === undefined ? 2 : right
      } catch (_) {}
      try { taskpane.Width = 420 } catch (_) {}
    }
    taskpane.Visible = true
    _ensureMessageBridge()
    if (!_ensureSidecar()) {
      _forwardToTaskpane({
        type: 'error',
        payload: { message: 'WPS 已打开面板，但当前 WPS 运行环境无法启动本地 AI 进程。' },
      })
    }
    return true
  } catch (error) {
    var detail = String((error && (error.stack || error.message)) || error)
    _diag('OnAction FAILED: ' + detail)
    console.error('[DeepExcel] open taskpane failed:', error)

    // 报错必须真的到达用户。原先只用 ShowDialog('data:text/html;...')，
    // 一旦 WPS 拒绝 data: URL，用户看到的就是"点了没反应"。按可靠性降级：
    // Alert（最朴素、最可能被支持）→ ShowDialog → 日志兜底。
    var shown = false
    try {
      if (typeof app.Alert === 'function') {
        app.Alert('DeepExcel 面板打开失败：' + detail)
        shown = true
      }
    } catch (_) {}
    if (!shown) {
      try {
        app.ShowDialog(
          'data:text/html;charset=utf-8,' + encodeURIComponent(
            '<h3>DeepExcel 面板打开失败</h3><p>' + detail + '</p>'
          ),
          'DeepExcel',
          480,
          260,
          false
        )
        shown = true
      } catch (_) {}
    }
    if (!shown) _diag('could not surface the error to the user at all; log only')
    return false
  }
}

function OnShowHelp() {
  var app = _application()
  if (!app) return false
  try {
    app.ShowDialog(
      'data:text/html;charset=utf-8,' + encodeURIComponent(
        '<h2>DeepExcel 使用帮助</h2><p>点击“打开面板”，在右侧输入你的表格任务。</p>'
      ),
      'DeepExcel 使用帮助',
      480,
      300,
      false
    )
    return true
  } catch (error) {
    console.error('[DeepExcel] help dialog failed:', error)
    return false
  }
}

function GetImage(control) {
  return control && control.Id === 'btnHelp' ? 'images/help.svg' : 'images/panel.svg'
}

function _respond(type, payload) {
  _forwardToTaskpane({ type: type, payload: payload })
}

var CONFIG_MESSAGE_TYPES = [
  'get_model_config', 'save_model_config', 'set_provider_models', 'set_default_provider',
  'switch_model', 'get_api_key', 'delete_api_key', 'refresh_models', 'test_api_key',
]

// ★ 模型配置类消息：不依赖 sidecar 进程（用户可能正是因为没配 key 才打开面板）
function _handleConfigMessage(type, payload) {
  if (CONFIG_MESSAGE_TYPES.indexOf(type) < 0) return false
  if (!_ensureConfigStore()) {
    _respond('error', { message: '当前 WPS 运行环境无法读写模型配置（缺少 Node 运行时）。' })
    return true
  }
  if (sidecar && sidecar.pythonPath) configStore.credentials.setPythonPath(sidecar.pythonPath)

  var provider = payload.provider
  switch (type) {
    case 'get_model_config':
      configStore.load()
      _respond('model_config', configStore.getSafeConfig())
      return true

    case 'save_model_config': {
      var saved = configStore.saveModelConfig({
        provider: provider,
        model: payload.model,
        apiKey: payload.apiKey,
        baseUrl: payload.baseUrl,
        maxTurns: payload.maxTurns,
      })
      if (saved.success) _pushConfigToSidecar()
      _respond('config_saved', saved)
      return true
    }

    case 'set_provider_models':
      _respond('provider_models_saved', configStore.updateProviderModels(provider, payload.models))
      return true

    case 'set_default_provider':
      _respond('default_provider_set', configStore.setDefaultProvider(provider))
      return true

    case 'switch_model': {
      var switched = configStore.switchProvider(provider, payload.model)
      if (!switched.success) {
        _respond('error', { message: switched.error || '切换模型失败' })
        return true
      }
      // 与 Excel 端一致：切模型要重启 sidecar，否则 base_url 还是旧的
      if (sidecar) {
        sidecar.restart()
        setTimeout(_pushConfigToSidecar, 800)
      }
      _respond('model_switched', switched)
      return true
    }

    case 'get_api_key': {
      var key = configStore.credentials.get(provider)
      _respond('api_key', { apiKey: key, hasKey: !!key })
      return true
    }

    case 'delete_api_key': {
      var removed = configStore.credentials.remove(provider)
      if (removed) configStore.setConnected(provider, false)
      _respond('api_key_deleted', { success: removed })
      return true
    }

    case 'refresh_models': {
      if (!configStore.hasProvider(provider)) {
        _respond('models_refreshed', { success: false, error: '未知的模型供应商' })
        return true
      }
      var refreshKey = payload.apiKey && payload.apiKey !== KEEP_API_KEY
        ? payload.apiKey
        : configStore.credentials.get(provider)
      var refreshCfg = configStore.current().Providers[provider]
      modelService.fetchModels({
        provider: provider,
        baseUrl: payload.baseUrl || refreshCfg.BaseUrl,
        apiKey: refreshKey,
      }).then(function (result) {
        if (!result.success) {
          _respond('models_refreshed', result)
          return
        }
        _respond('models_refreshed', {
          success: true,
          models: result.models,
          selected: refreshCfg.Models || [],
          currentModel: configStore.current().CurrentModel,
        })
      }).catch(function (error) {
        console.error('[DeepExcel] refresh_models failed:', error)
        _respond('models_refreshed', { success: false, error: '刷新模型列表失败，请检查 Base URL 和网络连接' })
      })
      return true
    }

    case 'test_api_key': {
      if (!configStore.hasProvider(provider)) {
        _respond('api_test_result', { success: false, error: '未知的模型供应商' })
        return true
      }
      var testKey = payload.apiKey && payload.apiKey !== KEEP_API_KEY
        ? payload.apiKey
        : configStore.credentials.get(provider)
      var testCfg = configStore.current().Providers[provider]
      modelService.testApiKey({
        provider: provider,
        baseUrl: payload.baseUrl || testCfg.BaseUrl,
        apiKey: testKey,
        model: payload.model || testCfg.DefaultModel,
      }).then(function (result) {
        configStore.setConnected(provider, result.success)
        _respond('api_test_result', result)
      }).catch(function (error) {
        console.error('[DeepExcel] test_api_key failed:', error)
        _respond('api_test_result', { success: false, error: '测试失败，请检查网络连接' })
      })
      return true
    }

    default:
      return false
  }
}

var SESSION_MESSAGE_TYPES = [
  'list_conversations', 'get_current_messages', 'new_conversation',
  'continue_conversation', 'delete_conversation',
  'list_attachments', 'upload_attachment', 'delete_attachment',
  'memory_get', 'memory_save', 'memory_clear',
]

// ★ 对话历史 / 附件消息：需要会话状态，但不强制 sidecar 在跑
// （历史列表、附件管理在 AI 进程没起来时也应该能看能删）
function _handleSessionMessage(type, payload) {
  if (SESSION_MESSAGE_TYPES.indexOf(type) < 0) return false

  var session = _session()
  if (!session) {
    _respond('error', { message: '当前 WPS 运行环境无法读写对话历史（缺少 Node 运行时）。' })
    return true
  }

  switch (type) {
    case 'list_conversations':
      _respond('conversations', { list: session.conversation.listConversations() })
      return true

    // ★ 工作簿记忆（侧车维护的 NOTES.md）：面板查看 / 修改 / 清除，文件与 Excel 端共用
    case 'memory_get':
    case 'memory_save':
    case 'memory_clear': {
      var memoryStore = require('./workbook-memory-store')
      var identity = _readWorkbookIdentity()
      var refusal = null
      try {
        if (type === 'memory_save') refusal = memoryStore.save(identity.key, identity.name, payload && payload.notes)
        if (type === 'memory_clear') memoryStore.clear(identity.key)
      } catch (error) {
        refusal = '保存失败：' + String(error.message || error)
      }
      _respond('memory', memoryStore.describe(identity.key, identity.name, refusal))
      return true
    }

    case 'get_current_messages':
      _respond('current_messages', { messages: session.conversation.currentMessages() })
      return true

    case 'new_conversation': {
      var newId = session.conversation.newConversation()
      // 与 Excel 端一致：重启 sidecar 彻底清掉 AI 侧上下文，再补发 config
      if (sidecar) {
        sidecar.restart()
        setTimeout(_pushConfigToSidecar, 800)
      }
      _respond('new_conversation', { conversationId: newId })
      return true
    }

    case 'continue_conversation': {
      var conversationId = payload.conversation_id
      if (!conversationId) {
        _respond('error', { message: 'conversation_id is required' })
        return true
      }
      var messages = session.conversation.continueConversation(conversationId)
      if (!messages) {
        _respond('error', { message: '找不到该对话: ' + conversationId })
        return true
      }
      // 重启 sidecar 并把历史文本回灌，让 AI"记得"之前聊过什么
      if (sidecar) {
        sidecar.restart()
        setTimeout(function () {
          _pushConfigToSidecar()
          try {
            sidecar.sendRestoreHistory(session.conversation.restorableMessages())
          } catch (error) {
            console.error('[DeepExcel] restore history failed:', error)
          }
        }, 800)
      }
      _respond('continue_conversation', { conversationId: conversationId, messages: messages })
      return true
    }

    case 'delete_conversation': {
      var deleteId = payload.conversation_id
      if (!deleteId) {
        _respond('error', { message: 'conversation_id is required' })
        return true
      }
      _respond('delete_conversation', {
        success: session.conversation.deleteConversation(deleteId),
        conversationId: deleteId,
      })
      return true
    }

    case 'list_attachments':
      _respond('attachments', { list: session.attachments.list() })
      return true

    case 'upload_attachment': {
      try {
        var info = session.attachments.add(payload.file_name, payload.file_base64)
        _respond('uploaded', info)
      } catch (error) {
        _respond('error', { message: String((error && error.message) || '上传失败') })
      }
      return true
    }

    case 'delete_attachment': {
      var fileName = payload.file_name
      if (!fileName) {
        _respond('error', { message: '缺少 file_name 参数' })
        return true
      }
      _respond('deleted', { success: session.attachments.remove(fileName), file_name: fileName })
      return true
    }

    default:
      return false
  }
}

// ★ 首次使用：工作簿结构（推荐在面板里算）/ 插入示例数据到新表。不需要 sidecar
function _handleStarterMessage(type, payload) {
  if (type !== 'get_starter' && type !== 'insert_sample') return false
  try {
    var starter = require('./starter-host')
    if (type === 'get_starter') {
      _respond('starter', starter.describe(_application()))
    } else {
      _respond('sample_inserted', { sheet: starter.insertSample(_application(), payload) })
    }
  } catch (error) {
    var prefix = type === 'get_starter' ? '读取工作簿结构失败：' : '插入示例失败：'
    _respond('error', { message: prefix + String((error && error.message) || error) })
  }
  return true
}

function _handleFrontendMessage(message) {
  if (!message || !message.type) return

  // ★ 先处理模型配置消息：这些不需要 sidecar 起来
  if (_handleConfigMessage(message.type, message.payload || {})) return
  // ★ 再处理对话历史 / 附件消息
  if (_handleSessionMessage(message.type, message.payload || {})) return
  if (_handleStarterMessage(message.type, message.payload || {})) return

  if (!_ensureSidecar()) {
    _forwardToTaskpane({
      type: 'error',
      payload: { message: 'WPS 本地 AI 进程尚未启动，请关闭 WPS 后重试。' },
    })
    return
  }

  var payload = message.payload || {}
  switch (message.type) {
    case 'user_message': {
      var content = payload.content || ''
      // ★ 记进当前对话，stream_end 时连同 AI 回复一起落盘
      var userSession = _session()
      if (userSession) userSession.conversation.appendUserMessage(content)
      sidecar.sendUserMessage(content, 'wps-' + Date.now(), _buildContext(), payload.steer === true,
        payload.permission_mode)
      break
    }
    case 'cancel':
      sidecar.sendCancel()
      break
    case 'set_permission_mode':
      sidecar.sendPermissionMode(payload.mode)
      break
    case 'permission_response':
      sidecar.sendPermissionResponse(payload.request_id, payload.decision)
      break
    case 'clarify_answer':
      sidecar.sendClarifyAnswer(payload.answer)
      break
    // new_conversation 由 _handleSessionMessage 处理（要先把当前对话存盘）
    default:
      console.warn('[DeepExcel] unsupported WPS frontend message:', message.type)
  }
}

function _forwardToTaskpane(event) {
  _ensureMessageBridge()
  if (messageChannel) {
    messageChannel.postMessage(event)
    return
  }
  try {
    if (taskpane && typeof taskpane.postMessage === 'function') {
      taskpane.postMessage(JSON.stringify(event))
    }
  } catch (error) {
    console.error('[DeepExcel] taskpane message failed:', error)
  }
}

function _buildContext() {
  try {
    var app = _application()
    var workbook = app && app.ActiveWorkbook
    // ★ 带上附件清单（名称/大小/绝对路径），agent 需要时自己去读文件
    var attachments = []
    try {
      var session = _session()
      if (session) attachments = session.attachments.contextList()
    } catch (error) {
      console.error('[DeepExcel] build attachment context failed:', error)
    }
    return {
      host_type: 'wps',
      workbook: workbook ? workbook.Name : '',
      path: workbook ? workbook.FullName : '',
      activeSheet: workbook && workbook.ActiveSheet ? workbook.ActiveSheet.Name : '',
      // 工作簿记忆按它找目录（与 Excel 端同一口径，同一个文件两边共用一份记忆）
      workbookKey: _workbookIdentity().key,
      attachments: attachments,
    }
  } catch (error) {
    return { host_type: 'wps', error: String(error.message || error) }
  }
}

window.OnRibbonLoad = OnRibbonLoad
window.OnPluginInit = OnPluginInit
window.OnPluginDestroy = OnPluginDestroy
window.OnAction = OnAction
window.OnShowHelp = OnShowHelp
window.GetImage = GetImage

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    OnRibbonLoad: OnRibbonLoad,
    OnPluginInit: OnPluginInit,
    OnPluginDestroy: OnPluginDestroy,
    OnAction: OnAction,
    OnShowHelp: OnShowHelp,
    GetImage: GetImage,
  }
}
