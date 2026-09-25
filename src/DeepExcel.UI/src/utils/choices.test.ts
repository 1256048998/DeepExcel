import { describe, expect, it } from 'vitest'
import { detectChoices, stripMarkdown } from './choices'

describe('detectChoices', () => {
  it('finds explicitly labelled options', () => {
    const out = detectChoices('有两种做法：\n\n方案 A：用 SUMIFS 汇总\n方案 B：插入透视表\n\n你选哪个？')
    expect(out.map(c => c.label)).toEqual(['方案 A', '方案 B'])
    expect(out[1].value).toBe('插入透视表')
  })

  it('finds lettered options', () => {
    expect(detectChoices('A. 覆盖原列\nB. 写到新列').map(c => c.value)).toEqual(['覆盖原列', '写到新列'])
  })

  it('treats a numbered list as options only when the user is asked to choose', () => {
    expect(detectChoices('你想怎么处理空值？\n1. 填 0\n2. 删除整行')).toHaveLength(2)
    expect(detectChoices('可以这样处理：\n\n1. 填 0\n2. 删除整行')).toHaveLength(2)
  })

  it('ignores a numbered summary of what was done', () => {
    const summary = '已完成：\n\n1. **C 列**：37 个金额从文本转成了数字。\n2. **D 列**：填入 `=C2/B2` 计算单价。\n\n要不要改成 IF？'
    expect(detectChoices(summary)).toEqual([])
  })

  it('needs at least two consecutive options', () => {
    expect(detectChoices('方案 A：先备份\n\n然后再说')).toEqual([])
  })

  it('strips markdown from option text', () => {
    const out = detectChoices('选一个：\n1. **按月**汇总\n2. 按 `部门` 汇总')
    expect(out.map(c => c.value)).toEqual(['按月汇总', '按 部门 汇总'])
  })

  it('accepts bullets in front of the labels', () => {
    expect(detectChoices('- A. 保留\n- B. 删除')).toHaveLength(2)
  })
})

describe('stripMarkdown', () => {
  it('removes emphasis, code and links', () => {
    expect(stripMarkdown('**粗** _斜_ `码` [链接](http://x)')).toBe('粗 斜 码 链接')
  })
})
