interface PermissionDrawerProps {
  visible: boolean
  tool: string
  args: Record<string, any>
  onAllow: () => void
  onDeny: () => void
}

// 高风险工具的中文描述（简洁版，去掉"AI 想执行"前缀，更自然）
const TOOL_DESC: Record<string, string> = {
  execute_vba: '执行 VBA 代码',
  execute_python: '执行 Python 代码',
  rollback: '回滚工作簿到快照',
  clean_data: '清洗数据（会修改单元格）',
  remove_duplicates: '删除重复项（会删除行）',
}

// 参数显示名映射
const ARG_LABELS: Record<string, string> = {
  code: '代码',
  sub: '子过程',
  address: '地址',
  range_address: '范围',
  snapshot_id: '快照 ID',
  operations: '操作',
}

function formatValue(v: any): string {
  if (v == null) return ''
  const s = typeof v === 'string' ? v : JSON.stringify(v)
  // 代码截断显示：保留前 240 字符
  return s.length > 240 ? s.slice(0, 240) + '\n…' : s
}

export function PermissionDrawer({ visible, tool, args, onAllow, onDeny }: PermissionDrawerProps) {
  const desc = TOOL_DESC[tool] || `执行 ${tool}`

  // 筛选要显示的参数（最多 5 个，跳过 null/undefined）
  const argEntries = Object.entries(args || {})
    .filter(([, v]) => v != null)
    .slice(0, 5)

  return (
    <div className={`permission-drawer ${visible ? 'visible' : ''}`}>
      <div className="permission-drawer-inner">
        {/* 顶部：图标 + 请求描述 + 工具名 chip */}
        <div className="permission-drawer-header">
          <svg className="permission-icon" width="14" height="14" viewBox="0 0 16 16" fill="none">
            {/* 中性的菱形请求图标，替代警告三角形（去 alert 感） */}
            <path
              d="M8 1.5L14.5 8L8 14.5L1.5 8L8 1.5Z"
              stroke="currentColor"
              strokeWidth="1.25"
              strokeLinejoin="round"
            />
            <circle cx="8" cy="8" r="1.5" fill="currentColor" />
          </svg>
          <span className="permission-action-text">请求</span>
          <span className="permission-tool-chip">{desc}</span>
        </div>

        {/* 参数预览：代码/参数用深色背景 monospace */}
        {argEntries.length > 0 && (
          <div className="permission-args">
            {argEntries.map(([k, v]) => (
              <div key={k} className="permission-arg">
                <span className="arg-key">{ARG_LABELS[k] || k}</span>
                <pre className="arg-val">{formatValue(v)}</pre>
              </div>
            ))}
          </div>
        )}

        {/* 底部：操作按钮 */}
        <div className="permission-actions">
          <span className="permission-hint">本次会话内允许后不再询问</span>
          <div className="permission-btns">
            <button className="perm-btn perm-deny" onClick={onDeny} type="button">
              拒绝
            </button>
            <button className="perm-btn perm-allow" onClick={onAllow} type="button">
              允许
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
