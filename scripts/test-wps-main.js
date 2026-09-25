'use strict'

// WPS 加载项冒烟测试：
//   1. main.js 能在类浏览器环境里加载，Ribbon/面板/图标回调正常
//   2. 模型配置消息（get_model_config / set_provider_models / switch_model …）
//      能走通 main.js → config-store → 回传前端 的整条链路
//
// 运行：node scripts/test-wps-main.js（build-wps.ps1 会自动调用）

const assert = require('assert')
const fs = require('fs')
const os = require('os')
const path = require('path')
const vm = require('vm')

const wpsDir = path.resolve(__dirname, '..', 'src', 'DeepExcel.Wps')
const sourcePath = path.join(wpsDir, 'main.js')
const source = fs.readFileSync(sourcePath, 'utf8')

// ★ 用临时目录当 APPDATA / LOCALAPPDATA，避免测试写坏开发机上真实的
// config.json、对话历史和附件
const tempAppData = fs.mkdtempSync(path.join(os.tmpdir(), 'deepexcel-wps-test-'))
process.env.APPDATA = tempAppData
process.env.LOCALAPPDATA = tempAppData

// ★ 预置一份"Excel 端（C#）写出的" config.json：PascalCase、deepseek 已被用户自定义排序。
// WPS 必须原样读出这个顺序，而不是用内置目录覆盖掉——这是两端共用配置的核心契约。
const excelWrittenConfig = {
  CurrentProvider: 'deepseek',
  CurrentModel: 'deepseek-v4-flash',
  DefaultProvider: 'deepseek',
  ModelCatalogVersion: 2,
  Providers: {
    deepseek: {
      Type: 'anthropic',
      DisplayName: 'DeepSeek',
      ApiKey: '',
      BaseUrl: 'https://api.deepseek.com/anthropic',
      Models: ['deepseek-v4-flash', 'deepseek-v4-pro', 'deepseek-chat'],
      DefaultModel: 'deepseek-v4-flash',
      Headers: {},
      SupportsVision: false,
      LastTestSuccess: false,
      ModelsCustomized: true,
    },
  },
  General: { MaxTurns: 33 },
  UI: { Theme: 'light' },
}
fs.mkdirSync(path.join(tempAppData, 'DeepExcel'), { recursive: true })
fs.writeFileSync(
  path.join(tempAppData, 'DeepExcel', 'config.json'),
  JSON.stringify(excelWrittenConfig, null, 2),
  'utf8',
)

const storage = new Map()
let createdUrl = null
let sidecarStarts = 0
let sidecarRestarts = 0
const sidecarConfigs = []
const sentToTaskpane = []

class FakeBroadcastChannel {
  constructor() { this.onmessage = null }
  addEventListener() {}
  postMessage(event) { sentToTaskpane.push(event) }
  close() {}
}

const userMessages = []
const restoredHistories = []

class FakeSidecar {
  constructor() {
    this.pythonPath = 'python'
    this.onEvent = null
    this.dispatcher = { edits: [], recordUserEdit(sheet, address) { this.edits.push([sheet, address]) } }
    FakeSidecar.lastInstance = this
  }
  start() { sidecarStarts++ }
  stop() {}
  sendUserMessage(text, sessionId, context) { userMessages.push({ text, sessionId, context }) }
  sendCancel() {}
  sendPermissionResponse() {}
  sendClarifyAnswer() {}
  sendRestoreHistory(messages) { restoredHistories.push(messages) }
  restart() { sidecarRestarts++ }
  updateConfig(baseUrl, model, apiKey) { sidecarConfigs.push({ baseUrl, model, apiKey }) }
}

const pane = { ID: 42, Visible: false }
const workbook = {
  Name: 'Book1.xlsx',
  FullName: 'C:\\Users\\tester\\Book1.xlsx',
  ActiveSheet: { Name: 'Sheet1' },
}
const application = {
  PluginStorage: {
    getItem: key => storage.get(key),
    setItem: (key, value) => storage.set(key, value),
  },
  CreateTaskPane: url => {
    createdUrl = url
    return pane
  },
  GetTaskPane: () => pane,
  Enum: { msoCTPDockPositionRight: 2 },
  ActiveWorkbook: workbook,
  ApiEvent: { AddApiEventListener: (name, fn) => { apiListeners[name] = fn } },
}
const apiListeners = {}

const windowObject = {
  Application: application,
  BroadcastChannel: FakeBroadcastChannel,
  addEventListener() {},
}
windowObject.window = windowObject

// 记录 main.js 创建的 BroadcastChannel，测试通过它注入前端消息
let createdChannel = null
class TrackingBroadcastChannel extends FakeBroadcastChannel {
  constructor(name) {
    super(name)
    createdChannel = this
  }
}

const context = {
  window: windowObject,
  document: { location: { toString: () => 'file:///C:/DeepExcel/main.html' } },
  BroadcastChannel: TrackingBroadcastChannel,
  console,
  encodeURIComponent,
  setTimeout,
  // 真实加载 config-store / model-service，只把 sidecar-host 换成假的
  require: request => (request === './sidecar-host'
    ? FakeSidecar
    : require(path.resolve(wpsDir, request))),
  module: { exports: {} },
}

vm.runInNewContext(source, context, { filename: sourcePath })

// ============ 1. 原有冒烟断言 ============
assert.strictEqual(context.window.OnRibbonLoad({}), true)
assert.strictEqual(context.window.OnAction({ Id: 'btnTogglePanel' }), true)
assert.strictEqual(createdUrl, 'file:///C:/DeepExcel/taskpane.html')
assert.strictEqual(pane.Visible, true)
assert.strictEqual(sidecarStarts, 1)
assert.strictEqual(context.window.GetImage({ Id: 'btnTogglePanel' }), 'images/panel.svg')
assert.strictEqual(context.window.GetImage({ Id: 'btnHelp' }), 'images/help.svg')

// 读后被改检测：sidecar 启动时订阅 SheetChange，用户改动（去掉 $）交给调度器的账本
assert.strictEqual(typeof apiListeners.SheetChange, 'function', 'SheetChange listener should be registered')
apiListeners.SheetChange({ Name: 'Sheet1' }, { Address: (rowAbs, colAbs) => (rowAbs || colAbs ? '$B$2:$C$3' : 'B2:C3') })
apiListeners.SheetChange({ Name: 'Sheet1' }, { Address: '$D$4' })
assert.deepStrictEqual(FakeSidecar.lastInstance.dispatcher.edits, [['Sheet1', 'B2:C3'], ['Sheet1', 'D4']])

// sidecar 启动后应立刻收到一份模型配置（否则 sidecar 没有 base_url / model）
assert.ok(sidecarConfigs.length >= 1, 'sidecar should receive config on start')
assert.ok(sidecarConfigs[0].baseUrl, 'sidecar config should carry baseUrl')

// ============ 2. 模型配置消息链路 ============
assert.ok(createdChannel, 'main.js should open a BroadcastChannel')

function send(type, payload) {
  const before = sentToTaskpane.length
  createdChannel.onmessage({ data: { type, payload: payload || {} } })
  return sentToTaskpane.slice(before)
}

function lastOfType(events, type) {
  for (let i = events.length - 1; i >= 0; i--) {
    if (events[i].type === type) return events[i].payload
  }
  return null
}

// 2.1 读取配置：Excel 端写的 PascalCase 配置要能原样读出（含用户排好的模型顺序）
const config = lastOfType(send('get_model_config'), 'model_config')
assert.ok(config, 'get_model_config should answer with model_config')
assert.ok(config.providers.deepseek, 'built-in providers should be present')
assert.strictEqual(config.providers.deepseek.hasApiKey, false, 'temp APPDATA has no credentials')
assert.deepStrictEqual(
  config.providers.deepseek.models,
  ['deepseek-v4-flash', 'deepseek-v4-pro', 'deepseek-chat'],
  'customized order written by the Excel host must survive',
)
assert.strictEqual(config.currentProvider, 'deepseek')
assert.strictEqual(config.defaultProvider, 'deepseek')
assert.strictEqual(config.general.maxTurns, 33, 'general settings must be preserved')
// 未在 config.json 里出现的厂商要用内置目录补齐
assert.ok(config.providers.anthropic && config.providers.anthropic.models.length > 0)

// 2.2 模型优先级排序（本次新功能的核心）：顺序即优先级，第 0 个成为主模型
const reordered = ['deepseek-v4-flash', 'deepseek-v4-pro']
const saved = lastOfType(send('set_provider_models', {
  provider: 'deepseek',
  models: reordered,
}), 'provider_models_saved')
assert.ok(saved && saved.success, 'set_provider_models should succeed')
assert.deepStrictEqual(saved.models, reordered)
assert.strictEqual(saved.defaultModel, 'deepseek-v4-flash')

// 2.3 顺序要落盘，重新读取后仍然保持
const reloaded = lastOfType(send('get_model_config'), 'model_config')
assert.deepStrictEqual(reloaded.providers.deepseek.models, reordered, 'order must persist')
assert.strictEqual(reloaded.providers.deepseek.defaultModel, 'deepseek-v4-flash')

// 2.4 去重 + 过滤非法模型名，且不能清空
const messy = lastOfType(send('set_provider_models', {
  provider: 'deepseek',
  models: ['deepseek-v4-pro', 'deepseek-v4-pro', '  ', 'bad name', 'deepseek-v4-flash'],
}), 'provider_models_saved')
assert.deepStrictEqual(messy.models, ['deepseek-v4-pro', 'deepseek-v4-flash'])
const empty = lastOfType(send('set_provider_models', { provider: 'deepseek', models: [] }), 'provider_models_saved')
assert.strictEqual(empty.success, false, 'empty model list must be rejected')

// 2.5 默认厂商 + 切换模型（切换后要重启 sidecar 并重新下发 config）
const def = lastOfType(send('set_default_provider', { provider: 'deepseek' }), 'default_provider_set')
assert.ok(def.success && def.defaultProvider === 'deepseek')

const restartsBefore = sidecarRestarts
const switched = lastOfType(send('switch_model', {
  provider: 'deepseek',
  model: 'deepseek-v4-flash',
}), 'model_switched')
assert.ok(switched.success, 'switch_model should succeed')
assert.strictEqual(sidecarRestarts, restartsBefore + 1, 'switch_model should restart sidecar')

// 不在列表里的模型要被拒绝
const badSwitch = lastOfType(send('switch_model', { provider: 'deepseek', model: 'not-exist' }), 'error')
assert.ok(badSwitch && /不在/.test(badSwitch.message), 'unknown model must be rejected')

// ============ 3. 对话历史 ============
// 3.1 一轮完整对话：用户消息 → 工具调用 → 流式回复 → stream_end 落盘
function emitFromSidecar(type, payload) {
  const handler = sidecarInstance().onEvent
  assert.ok(handler, 'main.js should hook sidecar.onEvent')
  handler({ type, payload })
}
function sidecarInstance() {
  return FakeSidecar.lastInstance
}

send('user_message', { content: '帮我统计 A 列' })
assert.strictEqual(userMessages.length, 1, 'user_message should reach the sidecar')
assert.ok(Array.isArray(userMessages[0].context.attachments), 'context should carry attachments')

emitFromSidecar('tool_call', { name: 'write_formula' })
emitFromSidecar('stream_delta', { delta: '已在 B1 ' })
emitFromSidecar('stream_delta', { delta: '写入 =SUM(A:A)' })
emitFromSidecar('stream_end', {})

const current = lastOfType(send('get_current_messages'), 'current_messages')
assert.strictEqual(current.messages.length, 3, 'user + tool + assistant')
assert.strictEqual(current.messages[0].role, 'user')
assert.deepStrictEqual(current.messages[1].toolGroup, ['write_formula'])
assert.strictEqual(current.messages[2].content, '已在 B1 写入 =SUM(A:A)', 'stream deltas must merge')
assert.strictEqual(current.messages[2].streaming, false, 'stream_end must clear the streaming flag')

// 3.2 历史列表：标题取首条用户消息
const conversations = lastOfType(send('list_conversations'), 'conversations')
assert.strictEqual(conversations.list.length, 1)
assert.strictEqual(conversations.list[0].title, '帮我统计 A 列')
const firstConversationId = conversations.list[0].id

// 3.3 历史文件要落在与 Excel 端一致的位置（同一个工作簿 → 同一个文件名）
const { fileNameFor } = require(path.join(wpsDir, 'conversation-store.js'))
const expectedHistoryFile = path.join(
  tempAppData, 'DeepExcel', 'history', fileNameFor(workbook.FullName))
assert.ok(fs.existsSync(expectedHistoryFile), 'history file path must match the C# host')
const historyOnDisk = JSON.parse(fs.readFileSync(expectedHistoryFile, 'utf8'))
assert.strictEqual(historyOnDisk.WorkbookKey, workbook.FullName, 'C#-compatible PascalCase keys')
assert.strictEqual(historyOnDisk.Conversations[0].Messages.length, 3)

// 3.4 新建对话：旧对话存盘、内存清空、sidecar 重启
const restartsBeforeNew = sidecarRestarts
const created = lastOfType(send('new_conversation'), 'new_conversation')
assert.ok(created.conversationId, 'new_conversation should return a new id')
assert.strictEqual(sidecarRestarts, restartsBeforeNew + 1)
assert.strictEqual(lastOfType(send('get_current_messages'), 'current_messages').messages.length, 0)

// 3.5 继续历史对话：消息恢复 + 历史回灌给 sidecar
const continued = lastOfType(send('continue_conversation', {
  conversation_id: firstConversationId,
}), 'continue_conversation')
assert.strictEqual(continued.messages.length, 3, 'history messages must come back')
assert.strictEqual(continued.conversationId, firstConversationId)

// 回灌给 AI 的内容：只要 user/assistant 文本，工具调用不回放（避免重复执行）
const { ConversationSession, ConversationHistory } = require(path.join(wpsDir, 'conversation-store.js'))
const probe = new ConversationSession(workbook.FullName, workbook.Name, new ConversationHistory())
probe.continueConversation(firstConversationId)
assert.deepStrictEqual(probe.restorableMessages(), [
  { role: 'user', content: '帮我统计 A 列' },
  { role: 'assistant', content: '已在 B1 写入 =SUM(A:A)' },
])

const missing = lastOfType(send('continue_conversation', { conversation_id: 'nope' }), 'error')
assert.ok(missing && /找不到/.test(missing.message))

// 3.6 删除对话
const deleted = lastOfType(send('delete_conversation', {
  conversation_id: firstConversationId,
}), 'delete_conversation')
assert.ok(deleted.success)
assert.strictEqual(lastOfType(send('list_conversations'), 'conversations').list.length, 0)

// ============ 4. 附件 ============
const uploaded = lastOfType(send('upload_attachment', {
  file_name: 'data.csv',
  file_base64: Buffer.from('a,b\n1,2\n').toString('base64'),
}), 'uploaded')
assert.strictEqual(uploaded.fileName, 'data.csv')
assert.strictEqual(uploaded.size, 8)
assert.ok(fs.existsSync(uploaded.filePath), 'attachment must be written to disk')

// 同名文件自动加序号，不覆盖
const uploadedAgain = lastOfType(send('upload_attachment', {
  file_name: 'data.csv',
  file_base64: Buffer.from('x').toString('base64'),
}), 'uploaded')
assert.strictEqual(uploadedAgain.fileName, 'data_1.csv')

// 可执行文件必须被白名单挡掉
const blocked = lastOfType(send('upload_attachment', {
  file_name: 'evil.exe',
  file_base64: Buffer.from('MZ').toString('base64'),
}), 'error')
assert.ok(blocked && /不支持的文件类型/.test(blocked.message))

const attachmentList = lastOfType(send('list_attachments'), 'attachments')
assert.strictEqual(attachmentList.list.length, 2)

// 附件要出现在发给 sidecar 的上下文里（agent 靠 path 自己读）
send('user_message', { content: '看看这个 CSV' })
const lastContext = userMessages[userMessages.length - 1].context
assert.strictEqual(lastContext.attachments.length, 2)
assert.ok(lastContext.attachments[0].path, 'attachment context needs an absolute path')

const removed = lastOfType(send('delete_attachment', { file_name: 'data.csv' }), 'deleted')
assert.ok(removed.success)
assert.ok(!fs.existsSync(uploaded.filePath), 'deleted attachment must be gone from disk')
assert.strictEqual(lastOfType(send('list_attachments'), 'attachments').list.length, 1)

// ============ 5. 异步消息 ============
// 未配置 key 时，拉取模型列表要给出明确提示而不是静默失败
// （refresh_models / test_api_key 是异步的，要等一个事件循环再看回包）
async function checkAsyncMessages() {
  const before = sentToTaskpane.length
  createdChannel.onmessage({ data: { type: 'refresh_models', payload: { provider: 'deepseek' } } })
  await new Promise(resolve => setTimeout(resolve, 50))
  const refreshed = lastOfType(sentToTaskpane.slice(before), 'models_refreshed')
  assert.ok(refreshed && refreshed.success === false, 'refresh_models should answer even without key')
  assert.ok(/API Key/.test(refreshed.error), 'should ask for API key first')

  const beforeTest = sentToTaskpane.length
  createdChannel.onmessage({ data: { type: 'test_api_key', payload: { provider: 'deepseek' } } })
  await new Promise(resolve => setTimeout(resolve, 50))
  const tested = lastOfType(sentToTaskpane.slice(beforeTest), 'api_test_result')
  assert.ok(tested && tested.success === false, 'test_api_key should answer even without key')

  // continue_conversation 会在 sidecar 重启后（延时 800ms）把历史回灌给 AI，
  // 这一步决定了"继续历史对话"时 AI 是否真的记得之前聊过什么
  await new Promise(resolve => setTimeout(resolve, 900))
  assert.ok(restoredHistories.length >= 1, 'history should be replayed to the restarted sidecar')
  const replayed = restoredHistories[restoredHistories.length - 1]
  assert.ok(replayed.length > 0, 'replayed history should not be empty')
  assert.ok(replayed.every(m => m.role === 'user' || m.role === 'assistant'),
    'only user/assistant text is replayed, tool calls are not')
}

checkAsyncMessages().then(() => {
  // ============ 3. 清理 ============
  fs.rmSync(tempAppData, { recursive: true, force: true })
  console.log('WPS_MAIN_SMOKE=PASS')
}).catch(error => {
  fs.rmSync(tempAppData, { recursive: true, force: true })
  console.error(error)
  process.exit(1)
})
