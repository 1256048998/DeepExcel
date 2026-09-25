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

/** 副本试跑：代码在工作簿副本上真实跑过一遍的结果（只有 VBA 有） */
export interface TrialData {
  success: boolean
  error?: string | null
  timed_out: boolean
  duration_ms: number
  errors_before: number
  errors_after: number
  /** 执行器的提示：静态检查警告、被改回的全局开关、自动点掉的弹窗 */
  note?: string | null
  /** 副本代表不了真实环境的地方 */
  not_representative: string[]
  /** 试跑之后用户改到了它涉及的格 */
  stale: boolean
  stale_message?: string
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
  trial?: TrialData | null
}

interface Props {
  preview: ChangePreviewData
  /** 试跑结果过期时「重新试跑」 */
  onRerunTrial?: () => void
}

/** 行号列表折叠显示，避免删 200 行时铺满面板 */
function formatRows(rows: number[]): string {
  if (rows.length === 0) return ''
  if (rows.length <= 8) return rows.join(', ')
  return `${rows.slice(0, 6).join(', ')} … 共 ${rows.length} 行`
}

function ChangeTable({ changes }: { changes: CellChangeView[] }) {
  if (changes.length === 0) return null
  return (
    <table className="preview-table">
      <tbody>
        {changes.map((change, index) => (
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
  )
}

/** 试跑结果的一句话标题 */
export function trialHeadline(preview: ChangePreviewData, trial: TrialData): string {
  const seconds = (trial.duration_ms / 1000).toFixed(1)
  if (trial.timed_out) return trial.error || '试跑超时，已终止'
  if (!trial.success) return `试跑出错：${trial.error || '未知错误'}`
  if (preview.affected_cells === 0) return `试跑完成（${seconds} 秒），没有改动任何单元格`
  return `试跑完成（${seconds} 秒），改动了 ${preview.affected_cells} 个单元格`
}

/**
 * 副本试跑：比「无法预览」多知道很多，但仍不是精确预告——真实执行时环境可能不同。
 * 所以用单独的样式，并把「副本代表不了的地方」逐条列出来。
 */
function TrialPreview({ preview, trial, onRerunTrial }: { preview: ChangePreviewData; trial: TrialData; onRerunTrial?: () => void }) {
  const failed = trial.timed_out || !trial.success
  return (
    <div className={`preview-block preview-trial${failed ? ' preview-trial-failed' : ''}`}>
      <div className="preview-summary">
        <span className="preview-badge preview-badge-trial">副本试跑</span>
        {trialHeadline(preview, trial)}
      </div>

      {trial.stale && (
        <div className="preview-stale">
          <span>⚠ {trial.stale_message || '试跑之后工作簿有改动，结果可能已不准确'}</span>
          {onRerunTrial && (
            <button type="button" className="preview-rerun" onClick={onRerunTrial}>重新试跑</button>
          )}
        </div>
      )}

      <ChangeTable changes={preview.changes} />

      {preview.truncated && (
        <div className="preview-note">
          仅列出前 {preview.changes.length} 处，共改动 {preview.affected_cells} 个单元格。
        </div>
      )}

      {preview.warnings.map((warning, index) => (
        <div key={index} className="preview-warning">
          ⚠ {warning}
        </div>
      ))}

      {trial.not_representative.length > 0 && (
        <div className="preview-caveats">
          <div className="preview-caveats-title">副本代表不了的地方</div>
          <ul>
            {trial.not_representative.map((item, index) => <li key={index}>{item}</li>)}
          </ul>
        </div>
      )}

      {trial.note && <div className="preview-note">{trial.note}</div>}

      <p className="preview-note">
        以上是代码在工作簿副本上的实际结果。允许后会在真实工作簿上重新执行同一段代码，执行前自动快照。
      </p>
    </div>
  )
}

export function ChangePreview({ preview, onRerunTrial }: Props) {
  if (preview.trial) {
    return <TrialPreview preview={preview} trial={preview.trial} onRerunTrial={onRerunTrial} />
  }

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

      <ChangeTable changes={preview.changes} />

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
