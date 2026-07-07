import { useState, useRef, useEffect } from 'react'
import { sendToHost, sendToHostWithResponse, onHostMessage } from './bridge'
import { MessageList } from './components/MessageList'
import { InputArea } from './components/InputArea'
import { StatusBar } from './components/StatusBar'
import { HistoryPanel } from './components/HistoryPanel'
import { AttachmentPanel } from './components/AttachmentPanel'
import { ConversationsPanel } from './components/ConversationsPanel'
import { ModelConfigPanel } from './components/ModelConfigPanel'
import type { Message, ConnectionStatus } from './types'

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
  const [status, setStatus] = useState<ConnectionStatus>('connecting')
  const [isClarifying, setIsClarifying] = useState(false)
  const [historyOpen, setHistoryOpen] = useState(false)
  // ★ 附件面板开关 + 附件列表
  const [attachmentsOpen, setAttachmentsOpen] = useState(false)
  const [attachments, setAttachments] = useState<AttachmentInfo[]>([])
  // ★ 历史对话弹窗
  const [conversationsOpen, setConversationsOpen] = useState(false)
  // ★ 模型配置弹窗（Ribbon 按钮触发）
  const [modelConfigOpen, setModelConfigOpen] = useState(false)

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
      } else if (data.type === 'error') {
        clearLoadingTimeout()
        // ★ 同样重置所有 streaming 状态
        setMessages(prev => [
          ...prev.map(m => m.streaming ? { ...m, streaming: false } : m),
          { role: 'assistant', content: `❌ ${data.payload.message}` }
        ])
        setLoading(false)
      } else if (data.type === 'open_model_config') {
        // ★ Ribbon 按钮触发的打开模型配置弹窗
        setModelConfigOpen(true)
      } else if (data.type === 'connection_ok') {
        setStatus('connected')
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
          <button
            className="new-conv-btn"
            onClick={handleNewConversation}
            title="开始新对话（当前对话会保存到历史）"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="12" y1="5" x2="12" y2="19"></line>
              <line x1="5" y1="12" x2="19" y2="12"></line>
            </svg>
            新对话
          </button>
          <button
            className="history-toggle-btn"
            onClick={() => setConversationsOpen(true)}
            title="查看历史对话"
          >
            历史对话
          </button>
          <button
            className="history-toggle-btn"
            onClick={() => setHistoryOpen(true)}
            title="查看历史版本"
          >
            版本
          </button>
          <button
            className="history-toggle-btn"
            onClick={() => setModelConfigOpen(true)}
            title="模型配置"
            type="button"
          >
            模型
          </button>
          <button
            className="attach-toggle-btn"
            onClick={openAttachments}
            title="附件管理"
            type="button"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48" />
            </svg>
            附件
            {attachments.length > 0 && (
              <span className="attach-toggle-badge">{attachments.length}</span>
            )}
          </button>
          <StatusBar status={status} />
        </div>
      </header>

      <MessageList
        messages={messages}
        loading={loading}
        onToggleToolGroup={toggleToolGroup}
        onClarifyAnswer={handleClarifyAnswer}
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
    </div>
  )
}
