import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import type { Message, UiEvent } from '../types'
import { applyUiEvent, runSummaryText } from './uiEvents'
import { TOOL_LABELS, toolLabel } from './toolCatalog'

const start = (id: string, name: string, args: Record<string, unknown> = {}): UiEvent =>
  ({ v: 1, kind: 'tool_start', id, name, args })
const end = (id: string, ok = true, extra: Record<string, unknown> = {}): UiEvent =>
  ({ v: 1, kind: 'tool_end', id, name: '', ok, duration_ms: 320, ...extra }) as UiEvent

function run(events: UiEvent[], initial: Message[] = []): Message[] {
  return events.reduce(applyUiEvent, initial)
}

describe('applyUiEvent', () => {
  it('pairs tool_start and tool_end by id inside one group', () => {
    const out = run([
      start('a', 'read_range', { address: 'A1:D20' }),
      start('b', 'write_formula', { address: 'E2', formula: '=SUM(A2:D2)' }),
      end('a', true, { summary: '20 行 × 4 列' }),
    ])
    expect(out).toHaveLength(1)
    const steps = out[0].toolSteps!
    expect(steps.map(s => s.status)).toEqual(['ok', 'running'])
    expect(steps[0].label).toBe('读取 A1:D20')
    expect(steps[0].summary).toBe('20 行 × 4 列')
    expect(steps[0].durationMs).toBe(320)
  })

  it('starts a new group after assistant text and seals the streaming text', () => {
    const initial: Message[] = [{ role: 'assistant', content: '我先看一下', streaming: true }]
    const out = run([start('a', 'read_workbook')], initial)
    expect(out).toHaveLength(2)
    expect(out[0].streaming).toBe(false)
    expect(out[1].role).toBe('tool')
  })

  it('keeps the error message and hint of a failed step', () => {
    const out = run([
      start('a', 'write_value', { address: 'A1', value: 1 }),
      end('a', false, { error: { code: 'tool_failed', message: '工作表受保护', hint: '先取消保护' } }),
    ])
    const step = out[0].toolSteps![0]
    expect(step.status).toBe('error')
    expect(step.error?.message).toBe('工作表受保护')
    expect(step.error?.hint).toBe('先取消保护')
  })

  it('ignores tool_end for an unknown id', () => {
    const before = run([start('a', 'read_range')])
    expect(applyUiEvent(before, end('zzz'))).toBe(before)
  })

  it('renders a classified error as a card', () => {
    const out = run([{ v: 1, kind: 'error', code: 'auth', message: '模型服务拒绝了凭据。', hint: '检查 API Key' }])
    expect(out[0].type).toBe('error')
    expect(out[0].error?.hint).toBe('检查 API Key')
  })

  it('shows run_summary only when it says something', () => {
    const quiet: UiEvent = { v: 1, kind: 'run_summary', outcome: 'success', tool_calls: 1, failed_calls: 0, duration_ms: 900 }
    expect(run([quiet])).toHaveLength(0)
    const maxTurns: UiEvent = { v: 1, kind: 'run_summary', outcome: 'max_turns', tool_calls: 20, failed_calls: 2, duration_ms: 95_000 }
    const out = run([maxTurns])
    expect(out[0].type).toBe('run_summary')
    expect(out[0].content).toContain('继续')
    expect(runSummaryText(maxTurns as Extract<UiEvent, { kind: 'run_summary' }>)).toContain('20 步（2 步失败）')
  })

  it('reports compaction with the pre-compaction size', () => {
    const out = run([{ v: 1, kind: 'compaction', trigger: 'auto', pre_tokens: 152_000 }])
    expect(out[0].type).toBe('compacted')
    expect(out[0].content).toContain('152k')
  })
})

describe('toolCatalog', () => {
  // 侧车注册表是真相来源：新增工具没写中文文案，面板就会显示英文函数名
  const py = readFileSync(resolve(__dirname, '../../../DeepExcel.Sidecar/excel_tools.py'), 'utf-8')
  const registered = [...py.matchAll(/@tool\("(\w+)"/g)].map(m => m[1])

  it('parses the sidecar registry', () => {
    expect(registered.length).toBeGreaterThan(40)
  })

  it('has a Chinese label for every registered tool', () => {
    const missing = registered.filter(name => !TOOL_LABELS[name])
    expect(missing).toEqual([])
  })

  it('has no labels for tools that no longer exist', () => {
    const stale = Object.keys(TOOL_LABELS).filter(name => !registered.includes(name))
    expect(stale).toEqual([])
  })

  it('never throws on missing or odd arguments', () => {
    for (const name of registered) {
      expect(() => toolLabel(name, {})).not.toThrow()
      expect(() => toolLabel(name, { address: 42, values: { __shape: [3, 2] } })).not.toThrow()
      expect(toolLabel(name, {})).not.toMatch(/undefined|null|\[object/)
    }
  })

  it('describes common calls in plain Chinese', () => {
    expect(toolLabel('mcp__excel__read_range', { address: 'Sheet1!A1:D20' })).toBe('读取 Sheet1!A1:D20')
    expect(toolLabel('write_range', { address: 'A1', values: { __shape: [500, 3], head: [] } })).toBe('批量写入 A1 500 行 × 3 列')
    expect(toolLabel('execute_vba', { code: 'Sub A()\nEnd Sub' })).toBe('运行 VBA（2 行）')
    expect(toolLabel('unknown_tool')).toBe('unknown_tool')
  })
})
