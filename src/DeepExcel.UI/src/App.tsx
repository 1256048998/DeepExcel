import { useState, useRef, useEffect } from 'react'
import { sendToHost, sendToHostWithResponse, onHostMessage } from './bridge'
import { MessageList } from './components/MessageList'
import { InputArea } from './components/InputArea'
import { HistoryPanel } from './components/HistoryPanel'
import { AttachmentPanel } from './components/AttachmentPanel'
import { ConversationsPanel } from './components/ConversationsPanel'
import { ModelConfigPanel } from './components/ModelConfigPanel'
import { PermissionDrawer } from './components/PermissionDrawer'
import { PromptManager } from './components/PromptManager'
import type { Message } from './types'
import type { PromptTemplate, PromptType } from './utils/prompts'
import { loadPrompts } from './utils/prompts'

// ★ AI Native 权限确认抽屉状态（PreToolUse hook 触发，从输入框上方 slide-up）
interface PermissionState {
  visible: boolean
  requestId?: string
  tool?: string
  args?: Record<string, any>
}

export interface AttachmentInfo {
  fileName: string
  size: number
}

export default function App() {
  const [messages, setMessages] = useState<Message[]>([
    {
      role: 'assistant',
      content: '你好！我是 DeepExcel AI Agent，可以帮你执行Excel任务。\n\n试试说：\n• "在A1写入=SUM(B1:B10)"\n• "把Sheet1的A列数据清洗一下"\n• "根据Sheet1数据创建柱状图"'
    }
  ])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const [isClarifying, setIsClarifying] = useState(false)
  const [historyOpen, setHistoryOpen] = useState(false)
  // ★ 附件面板开关 + 附件列表
  const [attachmentsOpen, setAttachmentsOpen] = useState(false)
  const [attachments, setAttachments] = useState<AttachmentInfo[]>([])
  // ★ 历史对话弹窗
  const [conversationsOpen, setConversationsOpen] = useState(false)
  // ★ 模型配置弹窗（Ribbon 按钮触发）
  const [modelConfigOpen, setModelConfigOpen] = useState(false)
  // ★ AI Native 权限确认抽屉（PreToolUse hook 请求时显示）
  const [permission, setPermission] = useState<PermissionState>({ visible: false })
  // ★ 提示词/技能：localStorage 持久化（用户级，跨工作簿保留），/ 触发下拉 + 管理面板
  const [prompts, setPrompts] = useState<PromptTemplate[]>([])
  const [promptManagerOpen, setPromptManagerOpen] = useState(false)
  // ★ 从历史消息保存时预填内容
  const [promptPrefillContent, setPromptPrefillContent] = useState<string | undefined>(undefined)
  // ★ 管理面板初始类型（从右上角"新建技能"入口传 skill；其余默认 prompt）
  const [promptManagerInitialType, setPromptManagerInitialType] = useState<PromptType>('prompt')

  // ★ "加载历史"开关：默认开启，打开面板时自动恢复最近一次对话
  // 状态持久化到 localStorage，用户可在顶部按钮切换
  const [autoLoadHistory, setAutoLoadHistory] = useState<boolean>(() => {
    try {
      const v = localStorage.getItem('deepexcel_autoload_history')
      // 默认开启：未设置或值为 "1" 都视为开启
      return v === null || v === '1'
    } catch { return true }
  })
  // 防止重复自动加载（连接重建时只加载一次）
  const hasAutoLoadedRef = useRef(false)

  // ★ Loading 超时兜底：如果 5 分钟内没收到 stream_end（消息丢失/前端卡死等），
  // 自动停止 loading，避免输入框永久禁用、停止按钮永久转圈
  const loadingTimerRef = useRef<number | null>(null)
  const LOADING_TIMEOUT_MS = 5 * 60 * 1000  // 5 分钟

  const startLoadingTimeout = () => {
    if (loadingTimerRef.current) window.clearTimeout(loadingTimerRef.current)
    loadingTimerRef.current = window.setTimeout(() => {
      console.warn('[DeepExcel] Loading timeout, auto-stopping')
      setLoading(false)
    }, LOADING_TIMEOUT_MS)
  }
  const clearLoadingTimeout = () => {
    if (loadingTimerRef.current) {
      window.clearTimeout(loadingTimerRef.current)
      loadingTimerRef.current = null
    }
  }

  // ★ 挂载时加载提示词模板
  useEffect(() => {
    setPrompts(loadPrompts())
  }, [])

  // ★ 从历史消息保存为提示词/技能：预填 content 并打开管理面板（默认 prompt 类型）
  const handleSaveAsPrompt = (content: string) => {
    setPromptPrefillContent(content)
    setPromptManagerInitialType('prompt')
    setPromptManagerOpen(true)
  }

  // ★ 从下拉新建：打开管理面板（无预填，默认 prompt）
  const handleCreatePrompt = () => {
    setPromptPrefillContent(undefined)
    setPromptManagerInitialType('prompt')
    setPromptManagerOpen(true)
  }

  // ★ 右上角管理入口：打开管理面板（无预填，默认 prompt）
  const handleOpenPromptManager = () => {
    setPromptPrefillContent(undefined)
    setPromptManagerInitialType('prompt')
    setPromptManagerOpen(true)
  }

  // ★ 关闭管理面板时清掉预填内容
  const handleClosePromptManager = () => {
    setPromptManagerOpen(false)
    setPromptPrefillContent(undefined)
  }

  // 监听来自C#主机的消息
  useEffect(() => {
    const unsubscribe = onHostMessage((data) => {
      if (data.type === 'stream_delta') {
        // ★ 启动 loading 超时兜底（仅首次 stream_delta 时启动）
        if (!loadingTimerRef.current) startLoadingTimeout()
        // 流式响应
        setMessages(prev => {
          const last = prev[prev.length - 1]
          if (last && last.role === 'assistant' && last.streaming) {
            return [
              ...prev.slice(0, -1),
              { ...last, content: last.content + data.payload.delta }
            ]
          }
          return [...prev, { role: 'assistant', content: data.payload.delta, streaming: true }]
        })
      } else if (data.type === 'stream_end') {
        clearLoadingTimeout()
        // ★ 遍历所有消息，把所有 streaming: true 都重置为 false
        // 之前只处理最后一条，如果中间有 tool_call 插入导致 streaming 消息不在末尾，
        // 光标会永久闪烁（streaming 永远不被重置）
        setMessages(prev => prev.map(m =>
          m.streaming ? { ...m, streaming: false } : m
        ))
        setLoading(false)
      } else if (data.type === 'tool_call') {
        // Agent 调用工具：合并连续 tool 消息为单个折叠组
        const toolName = data.payload.name
        setMessages(prev => {
          const last = prev[prev.length - 1]
          // 如果最后一条已是工具组，追加到组内
          if (last && last.role === 'tool' && last.toolGroup) {
            return [
              ...prev.slice(0, -1),
              { ...last, toolGroup: [...last.toolGroup, toolName] }
            ]
          }
          // 否则新建一个工具组
          return [...prev, {
            role: 'tool',
            content: '',
            toolName,
            toolGroup: [toolName],
            expanded: false
          }]
        })
      } else if (data.type === 'tool_result') {
        // 工具结果已不再单独展示（被合并到折叠组中），保留接口避免报错
        // 如果需要展示结果详情，可在此处把 result 写入对应工具组
      } else if (data.type === 'clarify') {
        const { question, options } = data.payload
        const safeQuestion = question ?? ''
        // 不再把选项拼到文本里，选项由按钮渲染
        setMessages(prev => [...prev, {
          role: 'assistant',
          content: safeQuestion,
          type: 'clarify',
          options
        }])
        setIsClarifying(true)
        setLoading(false)
      } else if (data.type === 'permission_request') {
        // ★ AI Native 权限确认：PreToolUse hook 请求用户确认高风险工具
        // 从输入框上方 slide-up 显示抽屉，不阻塞 Excel UI 线程
        const { request_id, tool, args } = data.payload
        setPermission({ visible: true, requestId: request_id, tool, args })
      } else if (data.type === 'compacted') {
        // ★ autocompact 触发：插入压缩提示卡，让用户知道发生了上下文压缩
        setMessages(prev => [...prev, {
          role: 'assistant',
          content: '对话已自动压缩，保留了关键上下文。',
          type: 'compacted'
        }])
      } else if (data.type === 'error') {
        clearLoadingTimeout()
        // ★ 同样重置所有 streaming 状态
        setMessages(prev => [
          ...prev.map(m => m.streaming ? { ...m, streaming: false } : m),
          { role: 'assistant', content: `❌ ${data.payload.message}` }
        ])
        setLoading(false)
      } else if (data.type === 'connection_ok') {
        // ★ 自动加载历史：开关开启时，连接建立后恢复最近一次对话
        if (autoLoadHistory && !hasAutoLoadedRef.current) {
          hasAutoLoadedRef.current = true
          autoLoadLatestConversation()
        }
      }
    })
    return unsubscribe
  }, [])

  const sendMessage = async (text?: string) => {
    const content = (text ?? input).trim()
    if (!content || loading) return

    const userMessage: Message = { role: 'user', content }
    setMessages(prev => [...prev, userMessage])
    setInput('')
    setLoading(true)
    setIsClarifying(false)

    try {
      await sendToHost({
        type: 'user_message',
        payload: { content }
      })
    } catch (err) {
      clearLoadingTimeout()
      setMessages(prev => [...prev, { role: 'assistant', content: `❌ 发送失败: ${err}` }])
      setLoading(false)
    }
  }

  const stopGeneration = () => {
    clearLoadingTimeout()
    sendToHost({ type: 'cancel', payload: {} })
    setLoading(false)
  }

  // 切换工具组的折叠/展开状态
  const toggleToolGroup = (idx: number) => {
    setMessages(prev => prev.map((m, i) =>
      i === idx && m.toolGroup ? { ...m, expanded: !m.expanded } : m
    ))
  }

  // 点击 clarify 选项按钮：直接作为用户回答发送
  const handleClarifyAnswer = (answer: string) => {
    sendMessage(answer)
  }

  // ★ 流式选项卡片点击：把选项作为用户消息发送
  // 复用 sendMessage，不重复 loading 状态判断
  const handleChoiceSelect = (choice: string) => {
    sendMessage(choice)
  }

  // ★ AI Native 权限确认：用户点击"允许并记住"或"拒绝"
  // 发送 permission_response 给 C#，C# 转发到 Python sidecar，sidecar 的 PreToolUse hook 收到后继续/中止
  const handlePermissionAllow = () => {
    if (permission.requestId) {
      sendToHost({
        type: 'permission_response',
        payload: { request_id: permission.requestId, decision: 'allow' }
      })
    }
    setPermission({ visible: false })
  }
  const handlePermissionDeny = () => {
    if (permission.requestId) {
      sendToHost({
        type: 'permission_response',
        payload: { request_id: permission.requestId, decision: 'deny' }
      })
    }
    setPermission({ visible: false })
  }

  // ★ 附件：加载列表
  const loadAttachments = async () => {
    try {
      const resp = await sendToHostWithResponse(
        { type: 'list_attachments', payload: {} },
        'attachments'
      )
      if (resp?.type === 'attachments' && resp.payload?.list) {
        setAttachments(resp.payload.list)
      }
    } catch (e) {
      console.warn('loadAttachments failed', e)
    }
  }

  // ★ 附件：上传文件（base64 发送）
  const uploadAttachment = async (file: File) => {
    return new Promise<void>((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = async () => {
        try {
          const base64 = (reader.result as string).split(',')[1]
          const resp = await sendToHostWithResponse(
            { type: 'upload_attachment', payload: { file_name: file.name, file_base64: base64 } },
            'uploaded'
          )
          if (resp?.type === 'uploaded') {
            await loadAttachments()
            resolve()
          } else {
            reject(resp?.payload?.message || '上传失败')
          }
        } catch (e) { reject(e) }
      }
      reader.onerror = () => reject(reader.error)
      reader.readAsDataURL(file)
    })
  }

  // ★ 附件：删除
  const deleteAttachment = async (fileName: string) => {
    try {
      await sendToHost({ type: 'delete_attachment', payload: { file_name: fileName } })
      await loadAttachments()
    } catch (e) {
      console.warn('deleteAttachment failed', e)
    }
  }

  // ★ 打开附件面板时刷新列表
  const openAttachments = () => {
    setAttachmentsOpen(true)
    loadAttachments()
  }

  // ★ 新建对话：清空前端消息，发 new_conversation 给 C#（会存当前对话+重启 sidecar）
  const handleNewConversation = async () => {
    try {
      await sendToHost({ type: 'new_conversation', payload: {} })
      // 重置前端状态
      setMessages([{
        role: 'assistant',
        content: '新对话已开始，请告诉我你需要做什么。'
      }])
      setInput('')
      setLoading(false)
      setIsClarifying(false)
      clearLoadingTimeout()
    } catch (e) {
      console.warn('new conversation failed', e)
    }
  }

  // ★ 自动加载最近一次对话：开关开启时连接建立后调用
  // 复用 list_conversations + continue_conversation 接口，取列表第一条（最新）
  const autoLoadLatestConversation = async () => {
    try {
      const listResp = await sendToHostWithResponse(
        { type: 'list_conversations', payload: {} },
        'conversations'
      )
      if (listResp?.type === 'conversations' && Array.isArray(listResp.payload?.list)) {
        const list = listResp.payload.list
        if (list.length === 0) return  // 无历史，保持欢迎语
        // 取最新的一条（列表已按更新时间倒序）
        const latest = list[0]
        const contResp = await sendToHostWithResponse(
          { type: 'continue_conversation', payload: { conversation_id: latest.id } },
          'continue_conversation'
        )
        if (contResp?.type === 'continue_conversation' && Array.isArray(contResp.payload?.messages)) {
          const msgs: Message[] = contResp.payload.messages.map((m: any) => ({
            role: m.role,
            content: m.content || '',
            type: m.type,
            options: m.options,
            toolGroup: m.toolGroup,
            streaming: false
          }))
          if (msgs.length > 0) {
            setMessages(msgs)
          }
        }
      }
    } catch (e) {
      console.warn('[DeepExcel] autoLoadLatestConversation failed', e)
    }
  }

  // ★ 切换"加载历史"开关
  const toggleAutoLoadHistory = () => {
    const next = !autoLoadHistory
    setAutoLoadHistory(next)
    try {
      localStorage.setItem('deepexcel_autoload_history', next ? '1' : '0')
    } catch { /* ignore */ }
  }

  // ★ 从历史继续对话：前端用历史 messages 恢复显示
  const handleContinueConversation = (historyMessages: Message[]) => {
    if (historyMessages.length === 0) {
      setMessages([{
        role: 'assistant',
        content: '已恢复历史对话（无历史消息）。'
      }])
    } else {
      setMessages(historyMessages)
    }
    setInput('')
    setLoading(false)
    setIsClarifying(false)
    clearLoadingTimeout()
  }

  return (
    <div className="app">
      <header className="app-header">
        <div className="app-header-actions">
          {/* 左侧：图标 + 文字 */}
          <button
            className="header-btn primary"
            onClick={handleNewConversation}
            title="开始新对话（当前对话会保存到历史）"
            type="button"
          >
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="12" y1="5" x2="12" y2="19"/>
              <line x1="5" y1="12" x2="19" y2="12"/>
            </svg>
            新建
          </button>
          <button
            className="header-btn"
            onClick={() => setConversationsOpen(true)}
            title="查看历史对话"
            type="button"
          >
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M3 12a9 9 0 1 0 9-9 9 9 0 0 0-6.36 2.64L3 8"/>
              <polyline points="3 4 3 8 7 8"/>
              <polyline points="12 7 12 12 15 14"/>
            </svg>
            历史对话
          </button>
        </div>
        <div className="app-header-actions right">
          {/* 右侧：纯图标 + tooltip */}
          {/* ★ 提示词/技能管理入口（用户级，跨工作簿保留） */}
          <button
            className="header-btn icon-only"
            onClick={handleOpenPromptManager}
            title="提示词与技能管理"
            type="button"
          >
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M19 21l-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z"/>
            </svg>
          </button>
          <button
            className={`header-btn icon-only ${autoLoadHistory ? 'on' : ''}`}
            onClick={toggleAutoLoadHistory}
            title={autoLoadHistory ? '加载历史：开启（打开面板自动恢复上次对话）' : '加载历史：关闭（打开面板为新对话）'}
            type="button"
          >
            {autoLoadHistory ? (
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="1 4 1 10 7 10"/>
                <path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/>
              </svg>
            ) : (
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" opacity="0.5">
                <polyline points="23 4 23 10 17 10"/>
                <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>
              </svg>
            )}
          </button>
          <button
            className="header-btn icon-only"
            onClick={() => setHistoryOpen(true)}
            title="历史版本"
            type="button"
          >
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/>
              <circle cx="12" cy="13" r="4"/>
            </svg>
          </button>
          <button
            className="header-btn icon-only"
            onClick={() => setModelConfigOpen(true)}
            title="模型配置"
            type="button"
          >
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="12" cy="12" r="3"/>
              <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>
            </svg>
          </button>
          <button
            className="header-btn icon-only attach-btn"
            onClick={openAttachments}
            title="附件管理"
            type="button"
          >
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48" />
            </svg>
            {attachments.length > 0 && (
              <span className="attach-badge">{attachments.length}</span>
            )}
          </button>
        </div>
      </header>

      <MessageList
        messages={messages}
        loading={loading}
        onToggleToolGroup={toggleToolGroup}
        onClarifyAnswer={handleClarifyAnswer}
        onChoiceSelect={handleChoiceSelect}
        onSaveAsPrompt={handleSaveAsPrompt}
      />

      {/* ★ AI Native 权限确认抽屉：从输入框上方 slide-up 显示，类似 Claude Code/Trae/Codex */}
      <PermissionDrawer
        visible={permission.visible}
        tool={permission.tool || ''}
        args={permission.args || {}}
        onAllow={handlePermissionAllow}
        onDeny={handlePermissionDeny}
      />

      <InputArea
        value={input}
        onChange={setInput}
        onSend={() => sendMessage()}
        onStop={stopGeneration}
        disabled={loading}
        isClarifying={isClarifying}
        onUploadAttachment={uploadAttachment}
        attachmentCount={attachments.length}
        onViewAttachments={openAttachments}
        attachments={attachments}
        onDeleteAttachment={deleteAttachment}
        permissionPending={permission.visible}
        prompts={prompts}
        onCreatePrompt={handleCreatePrompt}
      />

      <HistoryPanel
        open={historyOpen}
        onClose={() => setHistoryOpen(false)}
      />

      <AttachmentPanel
        open={attachmentsOpen}
        onClose={() => setAttachmentsOpen(false)}
        attachments={attachments}
        onDelete={deleteAttachment}
      />

      <ConversationsPanel
        open={conversationsOpen}
        onClose={() => setConversationsOpen(false)}
        onContinue={handleContinueConversation}
      />
      <ModelConfigPanel
        open={modelConfigOpen}
        onClose={() => setModelConfigOpen(false)}
      />
      <PromptManager
        visible={promptManagerOpen}
        prompts={prompts}
        onChange={setPrompts}
        onClose={handleClosePromptManager}
        prefillContent={promptPrefillContent}
        initialType={promptManagerInitialType}
      />
    </div>
  )
}
