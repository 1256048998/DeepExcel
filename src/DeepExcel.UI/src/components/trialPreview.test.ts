import { describe, expect, it } from 'vitest'
import { trialHeadline } from './ChangePreview'
import type { ChangePreviewData, TrialData } from './ChangePreview'

const preview = (affected: number): ChangePreviewData => ({
  previewable: true, summary: '', affected_cells: affected, truncated: false, structural: false,
  formulas_overwritten: 0, deleted_rows: [], deleted_columns: [], warnings: [], changes: [],
})
const trial = (over: Partial<TrialData>): TrialData => ({
  success: true, timed_out: false, duration_ms: 1234, errors_before: 0, errors_after: 0,
  not_representative: [], stale: false, ...over,
})

describe('trialHeadline', () => {
  it('reports how many cells the trial changed', () => {
    expect(trialHeadline(preview(3), trial({}))).toBe('试跑完成（1.2 秒），改动了 3 个单元格')
  })

  it('says plainly when nothing changed', () => {
    expect(trialHeadline(preview(0), trial({}))).toContain('没有改动任何单元格')
  })

  it('leads with the error when the trial failed or timed out', () => {
    expect(trialHeadline(preview(0), trial({ success: false, error: '下标越界' }))).toBe('试跑出错：下标越界')
    expect(trialHeadline(preview(0), trial({ success: false, timed_out: true, error: '副本上运行超过 20 秒仍未结束，已终止' })))
      .toContain('20 秒')
  })
})
