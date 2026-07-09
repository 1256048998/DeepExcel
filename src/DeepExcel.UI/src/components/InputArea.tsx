import { useRef, useState, ChangeEvent, useEffect } from 'react'

interface Props {
  value: string
  onChange: (val: string) => void
  onSend: () => void
  onStop?: () => void
  disabled: boolean
  isClarifying?: boolean
  // ★ 附件上传：点击回形针图标时触发
  onUploadAttachment?: (file: File) => Promise<void>
  // ★ 附件数量（显示徽章）
  attachmentCount?: number
  // ★ 查看附件列表（点击徽章时打开）
  onViewAttachments?: () => void
}

export function InputArea({
  value, onChange, onSend, onStop, disabled, isClarifying,
  onUploadAttachment, attachmentCount = 0, onViewAttachments,
}: Props) {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const [uploading, setUploading] = useState(false)
  const [uploadError, setUploadError] = useState<string | null>(null)

  // ★ 底部输入区可调整高度：拖拽 resize handle 改变 textarea 高度
  // 限制：最小 60px，最大 50vh（不超过整个面板高度的 50%）
  const [inputHeight, setInputHeight] = useState<number>(120)
  const dragStateRef = useRef<{ startY: number; startH: number } | null>(null)

  const onHandleMouseDown = (e: React.MouseEvent) => {
    e.preventDefault()
    dragStateRef.current = { startY: e.clientY, startH: inputHeight }
    document.body.style.cursor = 'ns-resize'
    document.body.style.userSelect = 'none'
  }

  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      const st = dragStateRef.current
      if (!st) return
      // 向下拖 → 减小高度（handle 在上方，输入区在下方）
      const delta = st.startY - e.clientY
      const maxH = Math.floor(window.innerHeight * 0.5)
      const next = Math.max(60, Math.min(maxH, st.startH + delta))
      setInputHeight(next)
    }
    const onUp = () => {
      dragStateRef.current = null
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [])

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

  return (
    <div className="input-area" style={{ height: inputHeight }}>
      {/* ★ 拖拽手柄：上下调整输入区高度 */}
      <div
        className="input-resize-handle"
        onMouseDown={onHandleMouseDown}
        title="拖拽调整高度"
      />
      {isClarifying && (
        <div className="clarify-hint">
          请回答上面的澄清问题
        </div>
      )}
      {uploadError && (
        <div className="upload-error">{uploadError}</div>
      )}
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
          value={value}
          onChange={e => onChange(e.target.value)}
          onKeyDown={e => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault()
              onSend()
            }
          }}
          placeholder={isClarifying ? '输入你的回答...' : '描述你的Excel任务...'}
          rows={2}
          disabled={disabled}
        />
      </div>
      {/* ★ 底部工具栏：左上传 + 右发送/停止，无额外边框 */}
      <div className="input-toolbar">
        {onUploadAttachment && (
          <button
            className="toolbar-btn"
            onClick={() => fileInputRef.current?.click()}
            disabled={uploading || disabled}
            title={uploading ? '上传中...' : '上传附件'}
            type="button"
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="17 8 12 3 7 8" />
              <line x1="12" y1="3" x2="12" y2="15" />
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
        {disabled && onStop ? (
          <button
            onClick={onStop}
            className="toolbar-btn stop"
            title="停止生成"
            type="button"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
              <rect x="6" y="6" width="12" height="12" rx="2" />
            </svg>
            <span>停止</span>
          </button>
        ) : (
          <button
            onClick={onSend}
            disabled={disabled || !value.trim()}
            className="toolbar-btn send"
            title="发送（Enter）"
            type="button"
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="22" y1="2" x2="11" y2="13" />
              <polygon points="22 2 15 22 11 13 2 9 22 2" />
            </svg>
          </button>
        )}
      </div>
    </div>
  )
}
