import { useRef, useState, ChangeEvent } from 'react'

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
    <div className="input-area">
      {isClarifying && (
        <div className="clarify-hint">
          请回答上面的澄清问题
        </div>
      )}
      {uploadError && (
        <div className="upload-error">{uploadError}</div>
      )}
      <div className="input-row">
        {/* ★ 附件上传按钮（回形针图标 + 数量徽章） */}
        {onUploadAttachment && (
          <>
            <input
              ref={fileInputRef}
              type="file"
              multiple
              style={{ display: 'none' }}
              onChange={handleFileChange}
            />
            <button
              className="attach-btn"
              onClick={() => fileInputRef.current?.click()}
              disabled={uploading || disabled}
              title={uploading ? '上传中...' : '上传附件'}
              type="button"
            >
              {/* 回形针图标 */}
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48" />
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
          </>
        )}
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
        {disabled && onStop ? (
          <button
            onClick={onStop}
            className="stop-button"
          >
            停止
          </button>
        ) : (
          <button
            onClick={onSend}
            disabled={disabled || !value.trim()}
            className="send-button"
          >
            发送
          </button>
        )}
      </div>
    </div>
  )
}
