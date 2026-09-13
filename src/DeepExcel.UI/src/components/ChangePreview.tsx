/**
 * 执行前的变更预览。
 *
 * 这个组件唯一重要的职责，是把两种情况在视觉上分开：
 *
 *   previewable = true  —— 这些就是将要发生的改动，逐格列出
 *   previewable = false —— 我们无法预知（VBA/Python），只能先快照
 *
 * 把两者显示成同一种样子，用户很快就会对前者也不再细看——而前者恰恰是唯一
 * 值得细看的。
 */

export interface CellChangeView {
  address: string
  before: string
  after: string
  kind: 'add' | 'overwrite' | 'clear' | 'format' | 'delete'
  overwrites_formula: boolean
}

export interface ChangePreviewData {
  previewable: boolean
  reason?: string | null
  summary: string
  affected_cells: number
  truncated: boolean
  structural: boolean
  formulas_overwritten: number
  deleted_rows: number[]
  deleted_columns: string[]
  warnings: string[]
  changes: CellChangeView[]
}

interface Props {
  preview: ChangePreviewData
}

/** 行号列表折叠显示，避免删 200 行时铺满面板 */
function formatRows(rows: number[]): string {
  if (rows.length === 0) return ''
  if (rows.length <= 8) return rows.join(', ')
  return `${rows.slice(0, 6).join(', ')} … 共 ${rows.length} 行`
}

export function ChangePreview({ preview }: Props) {
  if (!preview.previewable) {
    return (
      <div className="preview-block preview-unknown">
        <div className="preview-summary">
          <span className="preview-badge preview-badge-unknown">无法预览</span>
          {preview.reason || '此操作的影响无法预先推算'}
        </div>
        <p className="preview-note">
          已自动创建快照。执行后如果结果不对，可以在历史里回滚。
        </p>
      </div>
    )
  }

  if (preview.affected_cells === 0 && !preview.structural) {
    return (
      <div className="preview-block preview-noop">
        <div className="preview-summary">
          <span className="preview-badge preview-badge-noop">无变化</span>
          此操作不会改变任何内容
        </div>
        <p className="preview-note">
          通常说明目标区域或条件不对，建议确认后再执行。
        </p>
      </div>
    )
  }

  return (
    <div className="preview-block preview-exact">
      <div className="preview-summary">
        <span className="preview-badge preview-badge-exact">变更预览</span>
        {preview.summary}
      </div>

      {preview.deleted_rows.length > 0 && (
        <div className="preview-structural">删除行：{formatRows(preview.deleted_rows)}</div>
      )}
      {preview.deleted_columns.length > 0 && (
        <div className="preview-structural">删除列：{preview.deleted_columns.join(', ')}</div>
      )}

      {preview.changes.length > 0 && (
        <table className="preview-table">
          <tbody>
            {preview.changes.map((change, index) => (
              <tr
                key={`${change.address}-${index}`}
                className={change.overwrites_formula ? 'preview-row-formula' : undefined}
              >
                <td className="preview-addr">{change.address}</td>
                <td className="preview-before">{change.before || <em>空</em>}</td>
                <td className="preview-arrow">→</td>
                <td className="preview-after">
                  {change.kind === 'clear' ? <em>清空</em> : change.after || <em>空</em>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {preview.truncated && (
        <div className="preview-note">
          仅列出前 {preview.changes.length} 处，实际将修改 {preview.affected_cells} 个单元格。
        </div>
      )}

      {preview.warnings.map((warning, index) => (
        <div key={index} className="preview-warning">
          ⚠ {warning}
        </div>
      ))}
    </div>
  )
}
