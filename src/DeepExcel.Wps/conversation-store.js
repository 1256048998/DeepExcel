// src/DeepExcel.Wps/conversation-store.js
// 对话历史（对应 C# 端 Collaboration/ConversationHistory.cs + WorkbookSession 的对话部分）
//
// ★ 存储路径与文件结构和 Excel 加载项完全一致：
//   %LOCALAPPDATA%\DeepExcel\history\<sha256(workbookKey) 前 16 字节 hex>.json
//   { WorkbookKey, WorkbookName, SavedAt, Conversations: [{ Id, Title, CreatedAt, UpdatedAt, WorkbookName, Messages }] }
// 所以同一个工作簿在 Excel 里聊的历史，在 WPS 里能直接接着看、接着聊。
//
// 发给前端的是 camelCase（与 C# MessageBridge 的序列化策略一致），前端组件无需区分宿主。

const crypto = require('crypto')
const fs = require('fs')
const path = require('path')

const MAX_MESSAGES_PER_CONVERSATION = 200
const MAX_CONVERSATIONS = 20
const MAX_TOTAL_SIZE_BYTES = 50 * 1024 * 1024

function historyDir() {
  return path.join(process.env.LOCALAPPDATA || '', 'DeepExcel', 'history')
}

/** 与 C# ConversationHistory.GetFileName 一致：SHA256 取前 16 字节转小写 hex */
function fileNameFor(workbookKey) {
  const digest = crypto.createHash('sha256').update(String(workbookKey), 'utf8').digest()
  return digest.slice(0, 16).toString('hex') + '.json'
}

function toIso(value) {
  if (!value) return null
  const date = value instanceof Date ? value : new Date(value)
  return isNaN(date.getTime()) ? null : date.toISOString()
}

class ConversationHistory {
  constructor(options = {}) {
    this.dir = options.dir || historyDir()
  }

  filePathFor(workbookKey) {
    return path.join(this.dir, fileNameFor(workbookKey))
  }

  /** 加载某工作簿的全部对话；无历史返回空数组 */
  loadAll(workbookKey) {
    try {
      const file = this.filePathFor(workbookKey)
      if (!fs.existsSync(file)) return []
      const data = JSON.parse(fs.readFileSync(file, 'utf8'))
      const list = data && (data.Conversations || data.conversations)
      return Array.isArray(list) ? list : []
    } catch (e) {
      console.warn('[ConversationHistory] loadAll failed:', e.message)
      return []
    }
  }

  /** 覆盖写；同时执行数量/体积上限裁剪（规则与 C# 端一致） */
  saveAll(workbookKey, workbookName, conversations) {
    try {
      fs.mkdirSync(this.dir, { recursive: true })
      let list = Array.isArray(conversations) ? conversations.slice() : []

      if (list.length > MAX_CONVERSATIONS) {
        list.sort((a, b) => new Date(b.UpdatedAt || b.CreatedAt || 0) - new Date(a.UpdatedAt || a.CreatedAt || 0))
        list = list.slice(0, MAX_CONVERSATIONS)
      }
      for (const conv of list) {
        if (Array.isArray(conv.Messages) && conv.Messages.length > MAX_MESSAGES_PER_CONVERSATION) {
          conv.Messages = conv.Messages.slice(conv.Messages.length - MAX_MESSAGES_PER_CONVERSATION)
        }
      }

      fs.writeFileSync(this.filePathFor(workbookKey), JSON.stringify({
        WorkbookKey: workbookKey,
        WorkbookName: workbookName || '',
        SavedAt: new Date().toISOString(),
        Conversations: list,
      }), 'utf8')

      this._cleanupTotalSize()
      return true
    } catch (e) {
      console.warn('[ConversationHistory] saveAll failed:', e.message)
      return false
    }
  }

  deleteConversation(workbookKey, conversationId) {
    try {
      const list = this.loadAll(workbookKey)
      const kept = list.filter(c => c.Id !== conversationId)
      if (kept.length === list.length) return false
      this.saveAll(workbookKey, (kept[0] && kept[0].WorkbookName) || '', kept)
      return true
    } catch (e) {
      console.warn('[ConversationHistory] deleteConversation failed:', e.message)
      return false
    }
  }

  /** 历史总体积超限时按最旧优先删除（与 C# CleanupTotalSize 一致） */
  _cleanupTotalSize() {
    try {
      if (!fs.existsSync(this.dir)) return
      const files = fs.readdirSync(this.dir)
        .filter(name => name.endsWith('.json'))
        .map(name => {
          const full = path.join(this.dir, name)
          const stat = fs.statSync(full)
          return { full, size: stat.size, mtime: stat.mtimeMs }
        })
        .sort((a, b) => a.mtime - b.mtime)

      let total = files.reduce((sum, f) => sum + f.size, 0)
      if (total <= MAX_TOTAL_SIZE_BYTES) return
      for (const f of files) {
        if (total <= MAX_TOTAL_SIZE_BYTES * 0.8) break
        try {
          total -= f.size
          fs.unlinkSync(f.full)
        } catch (e) { /* 删不掉就跳过 */ }
      }
    } catch (e) {
      console.warn('[ConversationHistory] cleanup failed:', e.message)
    }
  }
}

/**
 * 单个工作簿的当前对话状态（对应 C# WorkbookSession 里对话相关的部分）。
 * 消息在内存里累积，stream_end / 新建对话 / 切换对话时落盘。
 */
class ConversationSession {
  constructor(workbookKey, workbookName, history) {
    this.workbookKey = workbookKey
    this.workbookName = workbookName || ''
    this.history = history || new ConversationHistory()
    this.messages = []
    this.currentConversationId = null
  }

  setWorkbookName(name) {
    if (name) this.workbookName = name
  }

  /** 列表只返回元信息，不带 messages（避免一次传大量数据） */
  listConversations() {
    return this.history.loadAll(this.workbookKey).map(c => ({
      id: c.Id,
      title: c.Title,
      createdAt: toIso(c.CreatedAt),
      updatedAt: toIso(c.UpdatedAt),
      workbookName: c.WorkbookName,
    }))
  }

  /** 前端展示用的当前对话消息（camelCase 与 C# 一致，字段名本身就是小写） */
  currentMessages() {
    return this.messages
  }

  newConversation() {
    if (this.messages.length > 0 && this.currentConversationId) this.saveCurrent()
    this.messages = []
    this.currentConversationId = crypto.randomUUID().replace(/-/g, '')
    return this.currentConversationId
  }

  /** 切到历史对话；返回该对话的消息，找不到返回 null */
  continueConversation(conversationId) {
    const all = this.history.loadAll(this.workbookKey)
    const target = all.find(c => c.Id === conversationId)
    if (!target) return null

    if (this.messages.length > 0 && this.currentConversationId &&
        this.currentConversationId !== conversationId) {
      this.saveCurrent()
    }

    this.messages = (target.Messages || []).map(m => Object.assign({}, m, { streaming: false }))
    this.currentConversationId = target.Id
    return this.messages
  }

  deleteConversation(conversationId) {
    const ok = this.history.deleteConversation(this.workbookKey, conversationId)
    if (this.currentConversationId === conversationId) {
      this.messages = []
      this.currentConversationId = null
    }
    return ok
  }

  // ============= 消息追加（与 C# WorkbookSession 的 Append* 对齐）=============

  appendUserMessage(content) {
    if (!this.currentConversationId) {
      this.currentConversationId = crypto.randomUUID().replace(/-/g, '')
    }
    this.messages.push({ role: 'user', content: content })
  }

  appendAssistantDelta(delta) {
    const last = this.messages[this.messages.length - 1]
    if (last && last.role === 'assistant' && last.streaming) {
      last.content += delta
    } else {
      this.messages.push({ role: 'assistant', content: delta, streaming: true })
    }
  }

  appendToolCall(toolName) {
    const last = this.messages[this.messages.length - 1]
    if (last && last.role === 'tool' && Array.isArray(last.toolGroup)) {
      last.toolGroup.push(toolName)
    } else {
      this.messages.push({ role: 'tool', content: '', toolGroup: [toolName] })
    }
  }

  appendClarify(question, options) {
    this.messages.push({
      role: 'assistant',
      content: question,
      type: 'clarify',
      options: options || [],
    })
  }

  /** stream_end：清 streaming 标志并落盘 */
  onStreamEnd() {
    for (const m of this.messages) {
      if (m.streaming) m.streaming = false
    }
    if (this.currentConversationId && this.messages.length > 0) this.saveCurrent()
  }

  saveCurrent() {
    try {
      const all = this.history.loadAll(this.workbookKey)
      let existing = all.find(c => c.Id === this.currentConversationId)
      if (!existing) {
        existing = {
          Id: this.currentConversationId,
          CreatedAt: new Date().toISOString(),
          WorkbookName: this.workbookName,
        }
        all.push(existing)
      }

      const firstUser = this.messages.find(m => m.role === 'user')
      const firstText = (firstUser && firstUser.content) || ''
      existing.Title = firstText.length > 30 ? firstText.slice(0, 30) + '...' : (firstText || '新对话')
      existing.UpdatedAt = new Date().toISOString()
      existing.WorkbookName = this.workbookName
      existing.Messages = this.messages.slice()

      this.history.saveAll(this.workbookKey, this.workbookName, all)
      return true
    } catch (e) {
      console.warn('[ConversationSession] saveCurrent failed:', e.message)
      return false
    }
  }

  /** 供 sidecar 恢复上下文：只要 user/assistant 的纯文本，工具调用不回放 */
  restorableMessages() {
    return this.messages
      .filter(m => (m.role === 'user' || m.role === 'assistant') && m.content && m.type !== 'clarify')
      .map(m => ({ role: m.role, content: m.content }))
  }
}

module.exports = { ConversationHistory, ConversationSession, fileNameFor }
