import type { StarterSuggestion } from '../utils/starter'

/** 欢迎语下面的「首次使用」卡片的状态，App 维护 */
export interface StarterView {
  /** 还在读工作簿结构 */
  loading: boolean
  /** empty：工作簿是空的，给示例；recommend：按结构推荐；unknown：宿主不支持，给通用示例 */
  mode: 'empty' | 'recommend' | 'sample' | 'unknown'
  suggestions: StarterSuggestion[]
  workbookName?: string
  /** 已插入的示例表名 */
  sampleSheet?: string
  inserting?: boolean
  error?: string
}

/** 宿主不支持 get_starter 时（老版本 / 读结构失败）的通用示例 */
export const GENERIC_STARTERS: StarterSuggestion[] = [
  { title: '在 A1 写入求和公式', prompt: '在A1写入=SUM(B1:B10)' },
  { title: '清洗 Sheet1 的 A 列', prompt: '把Sheet1的A列数据清洗一下' },
  { title: '根据 Sheet1 画柱状图', prompt: '根据Sheet1数据创建柱状图' },
]

interface Props {
  starter: StarterView
  busy: boolean
  onPick: (prompt: string) => void
  onInsertSample: () => void
}

export function StarterCard({ starter, busy, onPick, onInsertSample }: Props) {
  if (starter.loading) {
    return <div className="starter-card starter-loading">正在看看这个工作簿…</div>
  }

  const title = starter.mode === 'recommend'
    ? `为你的文件推荐${starter.workbookName ? `（${starter.workbookName}）` : ''}`
    : starter.mode === 'sample'
      ? `示例已放在新表「${starter.sampleSheet}」，试试这几个问题`
      : starter.mode === 'empty'
        ? '这个工作簿还是空的'
        : '试试说'
  const suggestions = starter.mode === 'empty' ? [] : starter.suggestions

  return (
    <div className="starter-card">
      <div className="starter-title">{title}</div>

      {starter.mode === 'empty' && (
        <div className="starter-empty">
          <p className="starter-note">
            可以先插入一份合成的中文销售明细（40 行，带几处常见的数据问题），用它试试 DeepExcel 能做什么。示例放在一张新表里，不会改动你已有的内容。
          </p>
          <button
            type="button"
            className="starter-sample-btn"
            onClick={onInsertSample}
            disabled={busy || starter.inserting}
          >
            {starter.inserting ? '正在插入…' : '插入示例数据'}
          </button>
          {starter.error && <div className="starter-error">{starter.error}</div>}
        </div>
      )}

      {suggestions.length > 0 && (
        <div className="starter-list">
          {suggestions.map((s, i) => (
            <button
              key={i}
              type="button"
              className="starter-item"
              title={s.prompt}
              onClick={() => onPick(s.prompt)}
              disabled={busy}
            >
              {s.title}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
