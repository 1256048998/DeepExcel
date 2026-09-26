import { useRef, useState, ChangeEvent, useEffect } from 'react'
import { PromptDropdown } from './PromptDropdown'
import { ModelPicker } from './ModelPicker'
import type { PromptTemplate } from '../utils/prompts'
import type { PermissionMode } from '../types'

export interface AttachmentItem {
  fileName: string
  size: number
}

/**
 * ★ 模型选择下拉单选项：每个已连接 provider 的每个模型作为一个选项。
 * 输入框底部工具栏的 ModelPicker 按 provider 分组显示。
 * value 用 `${provider}::${model}` 格式唯一标识。
 */
export interface ModelOption {
  provider: string             // provider key, e.g. "anthropic"
  providerDisplayName: string  // e.g. "Claude (Anthropic)"
  model: string                // model name, e.g. "claude-sonnet-5"
  isPrimary: boolean           // 是否该 provider 的主模型（模型优先级第 1 项 / DefaultModel）
}

interface Props {
  value: string
  onChange: (val: string) => void
  onSend: () => void
  onStop?: () => void
  disabled: boolean
  // ★ 任务进行中仍可输入：发出的话会在下一个工具结果后交给 AI（Claude Code 的排队消息）
  allowQueue?: boolean
  isClarifying?: boolean
  // ★ 附件上传：点击回形针图标时触发
  onUploadAttachment?: (file: File) => Promise<void>
  // ★ 附件数量（显示徽章）
  attachmentCount?: number
  // ★ 查看附件列表（点击徽章时打开）
  onViewAttachments?: () => void
  // ★ 附件列表（在 input-box 上方显示 chip，支持逐个删除）
  attachments?: AttachmentItem[]
  // ★ 删除附件
  onDeleteAttachment?: (fileName: string) => void
  // ★ 权限抽屉可见时隐藏停止按钮（防止抽屉收起后双击误点停止）
  permissionPending?: boolean
  // ★ 提示词模板列表（/ 触发下拉）
  prompts?: PromptTemplate[]
  // ★ 从下拉新建提示词（打开管理面板）
  onCreatePrompt?: () => void
  // ★ 模型选择下拉单：列出已连接 provider 的所有模型
  modelOptions?: ModelOption[]
  // ★ 当前选中的模型（`${provider}::${model}` 格式）
  selectedModel?: string
  // ★ 切换模型：用户选择后调用，App.tsx 会在 stream_end 后真正切换
  onModelChange?: (provider: string, model: string) => void
  // ★ 模型弹层底部「管理模型与密钥」：打开模型配置
  onManageModels?: () => void
  // 权限模式：点按钮或 Shift+Tab 轮换（任务进行中也能切，立即生效）
  permissionMode?: PermissionMode
  onPermissionModeChange?: (mode: PermissionMode) => void
}

export const PERMISSION_MODE_ORDER: PermissionMode[] = ['default', 'accept_writes', 'plan']

export const PERMISSION_MODE_TEXT: Record<PermissionMode, { label: string; hint: string }> = {
  default: { label: '每步确认', hint: '批量写入、清洗、删除和执行代码前都先让你确认（Shift+Tab 切换）' },
  accept_writes: { label: '自动应用写入', hint: '本次会话里写入不再逐个确认，每步都能回退；删除、清空和执行代码仍会确认。关掉面板就恢复每步确认' },
  plan: { label: '只出方案', hint: '只读不写：AI 先出一份变更方案，你批准后才执行' },
}

export function nextPermissionMode(mode: PermissionMode): PermissionMode {
  const i = PERMISSION_MODE_ORDER.indexOf(mode)
  return PERMISSION_MODE_ORDER[(i + 1) % PERMISSION_MODE_ORDER.length]
}

export function InputArea({
  value, onChange, onSend, onStop, disabled, allowQueue = false, isClarifying,
  onUploadAttachment, attachmentCount = 0, onViewAttachments,
  attachments = [], onDeleteAttachment,
  permissionPending = false,
  prompts = [], onCreatePrompt,
  modelOptions = [], selectedModel, onModelChange, onManageModels,
  permissionMode = 'default', onPermissionModeChange,
}: Props) {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)
  const [uploading, setUploading] = useState(false)
  const [uploadError, setUploadError] = useState<string | null>(null)

  // ★ 斜杠命令：输入以 / 开头时显示 PromptDropdown
  // 仅当 value === '/' 或 '/xxx'（单行、/ 是第一个字符）时触发
  const slashActive = value.startsWith('/') && !value.includes('\n') && !disabled
  const slashQuery = slashActive ? value.slice(1) : ''

  // ★ textarea 自动高度：根据内容行数调整，最小 1 行，最大 40vh
  // 类似 Trae 的行为，不需要用户手动拖拽调节
  useEffect(() => {
    const ta = textareaRef.current
    if (!ta) return
    ta.style.height = 'auto'
    const maxH = Math.floor(window.innerHeight * 0.4)
    ta.style.height = Math.min(ta.scrollHeight, maxH) + 'px'
  }, [value])

  const handleFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files
    if (!files || files.length === 0 || !onUploadAttachment) return
    setUploadError(null)
    setUploading(true)
    try {
      for (let i = 0; i < files.length; i++) {
        await onUploadAttachment(files[i])
      }
    } catch (err: any) {
      setUploadError(typeof err === 'string' ? err : (err?.message || '上传失败'))
    } finally {
      setUploading(false)
      if (fileInputRef.current) fileInputRef.current.value = ''
    }
  }

  // ★ 选中提示词：用 content 替换输入框内容（不自动发送，用户可编辑后 Enter）
  const handleSelectPrompt = (prompt: PromptTemplate) => {
    onChange(prompt.content)
  }

  // ★ 点击"+ 新建提示词"：清掉 / 前缀，打开管理面板
  const handleCreateNewPrompt = () => {
    onChange('')
    onCreatePrompt?.()
  }

  // ★ ESC 关闭下拉：清掉 / 前缀（保留剩余字符作为普通输入）
  const handleCloseDropdown = () => {
    // 仅去掉开头的 /，保留用户已输入的查询字符
    if (value.startsWith('/')) onChange(value.slice(1))
  }

  return (
    <div className="input-area">
      {uploadError && (
        <div className="upload-error">{uploadError}</div>
      )}
      {/* ★ 附件 chip 列表：显示在 input-box 上方，支持点击 x 删除 */}
      {attachments.length > 0 && onDeleteAttachment && (
        <div className="attach-chips">
          {attachments.map(att => (
            <span key={att.fileName} className="attach-chip" title={att.fileName}>
              <span className="attach-chip-name">{att.fileName}</span>
              <button
                className="attach-chip-x"
                onClick={() => onDeleteAttachment(att.fileName)}
                title="移除"
                type="button"
              >
                <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="18" y1="6" x2="6" y2="18" />
                  <line x1="6" y1="6" x2="18" y2="18" />
                </svg>
              </button>
            </span>
          ))}
        </div>
      )}
      {/* ★ input-box 容器：相对定位，PromptDropdown 绝对定位浮在上方 */}
      <div className="input-box-wrapper">
        {slashActive && (
          <PromptDropdown
            query={slashQuery}
            prompts={prompts}
            onSelect={handleSelectPrompt}
            onCreateNew={handleCreateNewPrompt}
            onClose={handleCloseDropdown}
          />
        )}
      {/* ★ input-box：包裹输入框 + 工具栏，外层浅灰色边框 */}
      <div className="input-box">
      <div className="input-row">
        {/* ★ 附件上传按钮（隐藏 file input） */}
        {onUploadAttachment && (
          <input
            ref={fileInputRef}
            type="file"
            multiple
            style={{ display: 'none' }}
            onChange={handleFileChange}
          />
        )}
        {/* ★ textarea 占据完整宽度，图标按钮在底部工具栏 */}
        <textarea
          ref={textareaRef}
          value={value}
          onChange={e => onChange(e.target.value)}
          onKeyDown={e => {
            // Shift+Tab 轮换权限模式（Claude Code 同款）
            if (e.key === 'Tab' && e.shiftKey && onPermissionModeChange) {
              e.preventDefault()
              onPermissionModeChange(nextPermissionMode(permissionMode))
              return
            }
            // 输入法选词时的 Enter 是确认候选词，不是发送
            if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
              e.preventDefault()
              onSend()
            }
          }}
          placeholder={isClarifying ? '输入你的回答...'
            : disabled && allowQueue ? '补充或纠正（会在下一步交给 AI）…' : '描述你的Excel任务...'}
          rows={2}
          disabled={disabled && !allowQueue}
        />
      </div>
      {/* ★ 底部工具栏：左「+」附件与权限模式，右模型选择与发送/停止 */}
      <div className="input-toolbar">
        {/* 左侧工具组：上传按钮 */}
        <div className="toolbar-left">
          {onUploadAttachment && (
            <button
              className={`composer-plus${uploading ? ' uploading' : ''}`}
              onClick={() => fileInputRef.current?.click()}
              disabled={uploading || disabled}
              title={uploading ? '上传中...' : '添加附件（图片、PDF、表格等）'}
              aria-label="添加附件"
              type="button"
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
                <line x1="12" y1="5" x2="12" y2="19" />
                <line x1="5" y1="12" x2="19" y2="12" />
              </svg>
              {attachmentCount > 0 && (
                <span
                  className="attach-badge"
                  title={`${attachmentCount} 个附件，点击查看`}
                  onClick={(e) => {
                    e.stopPropagation()
                    onViewAttachments?.()
                  }}
                >
                  {attachmentCount}
                </span>
              )}
              {uploading && <span className="attach-loading" />}
            </button>
          )}
          {onPermissionModeChange && (
            <button
              type="button"
              className={`toolbar-mode mode-${permissionMode}`}
              onClick={() => onPermissionModeChange(nextPermissionMode(permissionMode))}
              title={PERMISSION_MODE_TEXT[permissionMode].hint}
              aria-label={`权限模式：${PERMISSION_MODE_TEXT[permissionMode].label}，点击切换`}
            >
              {PERMISSION_MODE_TEXT[permissionMode].label}
            </button>
          )}
        </div>
        {/* 右侧工具组：模型选择 + 发送/停止 */}
        <div className="toolbar-right">
          {/* ★ 模型选择下拉单：列出已连接 provider 的所有模型，按 provider 分组。
              默认值 = 默认厂商的主模型（模型优先级第 1 项）。
              选择后不立即切换，等当前对话输出结束（stream_end）后才切换。 */}
          {onModelChange && modelOptions.length > 0 && (
            <ModelPicker
              options={modelOptions}
              value={selectedModel}
              onChange={onModelChange}
              disabled={disabled}
              onManage={onManageModels}
            />
          )}
          {disabled && allowQueue && value.trim() && (
            <button
              onClick={onSend}
              className="toolbar-btn send queue"
              title="补充给 AI（Enter）：在下一步执行前送达"
              aria-label="补充给 AI"
              type="button"
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <line x1="12" y1="19" x2="12" y2="5" />
                <polyline points="5 12 12 5 19 12" />
              </svg>
            </button>
          )}
          {disabled && onStop && !permissionPending ? (
            <button
              onClick={onStop}
              className="toolbar-btn stop"
              title="停止"
              aria-label="停止"
              type="button"
            >
              <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                <rect x="4" y="4" width="16" height="16" rx="3" />
              </svg>
            </button>
          ) : (
            <button
              onClick={onSend}
              disabled={disabled || !value.trim()}
              className="toolbar-btn send"
              title="发送（Enter）"
              aria-label="发送"
              type="button"
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <line x1="12" y1="19" x2="12" y2="5" />
                <polyline points="5 12 12 5 19 12" />
              </svg>
            </button>
          )}
        </div>
      </div>
      </div>
      </div>
    </div>
  )
}
