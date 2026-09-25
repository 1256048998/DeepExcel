import { describe, expect, it } from 'vitest'
import {
  profileSheet, recommendStarters, sampleQuestions, sampleRows,
  SAMPLE_HEADERS, SAMPLE_SHEET_NAME,
} from './starter'
import type { Cell, OutlineSheet } from './starter'

const sheet = (name: string, grid: Cell[][], extra: Partial<OutlineSheet> = {}): OutlineSheet =>
  ({ name, rows: grid.length, columns: grid[0]?.length ?? 0, grid, ...extra })

describe('profileSheet', () => {
  it('classifies columns and counts dirty cells', () => {
    const p = profileSheet(sheet('明细', [
      ['日期', '部门', '金额'],
      [45474, '财务部', 100],
      [45475, ' 市场部', '1,280'],
      [45476, '财务部', 300],
      [45476, '财务部', 300],
    ], { date_columns: [true, false, false] }))

    expect(p.columns.map(c => c.kind)).toEqual(['date', 'text', 'number'])
    expect(p.columns[2].textNumbers).toBe(1)
    expect(p.columns[1].padded).toBe(1)
    expect(p.duplicateRows).toBe(1)
    expect(p.dataRows).toBe(4)
  })

  it('counts error values and treats blank columns as empty', () => {
    const p = profileSheet(sheet('S', [['A', 'B'], ['#N/A', null], ['#REF!', '']]))
    expect(p.errors).toBe(2)
    expect(p.columns[1].kind).toBe('empty')
  })
})

describe('recommendStarters', () => {
  it('offers the sample when every sheet is empty', () => {
    expect(recommendStarters({ workbook_name: '工作簿1', sheets: [sheet('Sheet1', [])] }).mode).toBe('empty')
    expect(recommendStarters(null).mode).toBe('empty')
  })

  it('leads with cleaning, then a grouped summary, at most three', () => {
    const plan = recommendStarters({ workbook_name: 'a.xlsx', sheets: [sheet('销售', [
      ['下单日期', '区域', '金额'],
      [45474, '华东', 100], [45480, '华南', '200'], [45490, '华东 ', 50], [45500, '华北', 80],
    ], { date_columns: [true, false, false] })] })

    expect(plan.mode).toBe('recommend')
    expect(plan.suggestions).toHaveLength(3)
    expect(plan.suggestions[0].title).toContain('清洗「销售」')
    expect(plan.suggestions[0].title).toContain('「金额」有文本型数字')
    expect(plan.suggestions[1].prompt).toContain('按「区域」汇总「销售」的「金额」')
    expect(plan.suggestions[2].prompt).toContain('按「下单日期」的月份')
  })

  it('suggests fixing formula errors and merging same-shaped sheets', () => {
    const month = (name: string): OutlineSheet => sheet(name, [['姓名', '工资'], ['张三', 5000], ['李四', '#DIV/0!']])
    const plan = recommendStarters({ workbook_name: 'b.xlsx', sheets: [month('1月'), month('2月')] })
    expect(plan.suggestions.some(s => s.title.includes('公式错误'))).toBe(true)
    expect(plan.suggestions.some(s => s.title.includes('2 张结构相同的表'))).toBe(true)
  })

  it('falls back to a tour of the workbook', () => {
    const plan = recommendStarters({ workbook_name: 'c.xlsx', sheets: [sheet('备注', [['说明'], ['这是一张备注表']])] })
    expect(plan.suggestions[0].title).toBe('先看看这个工作簿里都有什么')
  })
})

describe('sample workbook', () => {
  it('is deterministic and carries the planted problems', () => {
    const rows = sampleRows()
    expect(rows).toEqual(sampleRows())
    expect(rows[0]).toEqual(SAMPLE_HEADERS)
    expect(rows).toHaveLength(42)
    // 撇号前缀让宿主把金额按文本写入
    expect(rows.filter(r => typeof r[6] === 'string' && (r[6] as string).startsWith("'"))).toHaveLength(2)
    expect(rows.filter(r => typeof r[2] === 'string' && r[2] !== (r[2] as string).trim())).toHaveLength(2)
    expect(rows.filter(r => r[1] === '')).toHaveLength(1)
  })

  it('once inserted, is recognised by the recommender as dirty and summarisable', () => {
    // 宿主写入后，文本金额读回来没有撇号
    const grid = sampleRows().map(r => r.map(v => (typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v)))
    const plan = recommendStarters({ workbook_name: 'x', sheets: [sheet(SAMPLE_SHEET_NAME, grid, { date_columns: [true] })] })
    expect(plan.suggestions[0].title).toContain('文本型数字')
    expect(plan.suggestions[0].title).toContain('多余空格')
    expect(plan.suggestions[0].title).toContain('重复行')
    expect(plan.suggestions.some(s => s.prompt.includes('按「区域」汇总'))).toBe(true)
  })

  it('has three demo questions naming the sheet', () => {
    const qs = sampleQuestions('示例-销售明细(2)')
    expect(qs).toHaveLength(3)
    qs.forEach(q => expect(q.prompt).toContain('「示例-销售明细(2)」'))
  })
})
