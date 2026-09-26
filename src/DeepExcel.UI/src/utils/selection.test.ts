import { describe, expect, it } from 'vitest'
import { cellCountText, sameSelection, selectionLabel, shouldShowSelection } from './selection'

describe('selection bar', () => {
  it('shows only real ranges, not a lone cursor cell', () => {
    expect(shouldShowSelection(null)).toBe(false)
    expect(shouldShowSelection({ address: '' })).toBe(false)
    expect(shouldShowSelection({ sheet: 'S', address: 'B2', cells: 1 })).toBe(false)
    expect(shouldShowSelection({ sheet: 'S', address: 'A1:D20', cells: 80 })).toBe(true)
  })

  it('labels the range the way Excel writes it', () => {
    expect(selectionLabel({ sheet: 'Sheet1', address: 'A1:D20' })).toBe('Sheet1!A1:D20')
    expect(selectionLabel({ sheet: '销售明细', address: 'A1:D20' })).toBe('销售明细!A1:D20')
    expect(selectionLabel({ sheet: "Q1 老板's 表", address: 'C:C' })).toBe("'Q1 老板''s 表'!C:C")
    expect(selectionLabel({ address: 'A1:B2' })).toBe('A1:B2')
  })

  it('shortens big cell counts', () => {
    expect(cellCountText(80)).toBe('80 格')
    expect(cellCountText(12000)).toBe('1.2 万格')
    expect(cellCountText(2 * 1048576)).toBe('209.7 万格')
    expect(cellCountText(17179869184)).toBe('171.8 亿格')
  })

  it('treats a re-sent identical selection as unchanged', () => {
    expect(sameSelection({ sheet: 'S', address: 'A1:B2' }, { sheet: 'S', address: 'A1:B2', cells: 4 })).toBe(true)
    expect(sameSelection({ sheet: 'S', address: 'A1:B2' }, { sheet: 'T', address: 'A1:B2' })).toBe(false)
    expect(sameSelection(null, { sheet: 'S', address: 'A1' })).toBe(false)
  })
})
