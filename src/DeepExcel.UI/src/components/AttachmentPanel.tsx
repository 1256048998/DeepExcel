export interface AttachmentInfo {
  fileName: string
  size: number
}

interface Props {
  open: boolean
  onClose: () => void
  attachments: AttachmentInfo[]
  onDelete: (fileName: string) => void
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return bytes + ' B'
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB'
}

export function AttachmentPanel({ open, onClose, attachments, onDelete }: Props) {
  if (!open) return null

  const handleDelete = (fileName: string) => {
    if (!confirm(`确定删除附件「${fileName}」吗？`)) return
    onDelete(fileName)
  }

  return (
    <div className="history-panel-overlay" onClick={onClose}>
      <div className="history-panel" onClick={(e) => e.stopPropagation()}>
        <div className="history-header">
          <h3>附件管理</h3>
          <div className="history-actions">
            <button
              className="history-close-btn"
              onClick={onClose}
              title="关闭"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="18" y1="6" x2="6" y2="18"></line>
                <line x1="6" y1="6" x2="18" y2="18"></line>
              </svg>
            </button>
          </div>
        </div>

        <div className="history-list">
          {attachments.length === 0 && (
            <div className="history-empty">
              暂无附件
              <div className="history-empty-hint">
                点击输入框左侧的回形针图标上传附件
              </div>
            </div>
          )}
          {attachments.map((a) => (
            <div key={a.fileName} className="history-item">
              <div className="history-item-info">
                <div className="history-item-title">
                  <span className="history-wb-icon" aria-hidden="true">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48" />
                    </svg>
                  </span>
                  <span className="history-wb-name">{a.fileName}</span>
                </div>
                <div className="history-item-meta">
                  <span className="history-time">{formatSize(a.size)}</span>
                </div>
              </div>
              <div className="history-item-actions">
                <button
                  className="history-delete-btn"
                  onClick={() => handleDelete(a.fileName)}
                  title="删除"
                >
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <polyline points="3 6 5 6 21 6"></polyline>
                    <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                  </svg>
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}
