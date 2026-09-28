// 「数据与隐私」页的守卫：页面上说收集什么，必须和服务端白名单、客户端实际上报的事件一致。
// 服务端加了字段、客户端开始上报新事件，这里会失败，逼着同步改页面文案。
import { readFileSync, readdirSync, statSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { NEVER_COLLECTED, TELEMETRY_EVENTS, TELEMETRY_NOT_SENT } from './privacy'

const repo = resolve(__dirname, '../../../..')

// 解析 server/app/routers/telemetry.py 里的 ALLOWLIST：事件名 → 字段名
function serverAllowlist(): Record<string, string[]> {
  const src = readFileSync(join(repo, 'server/app/routers/telemetry.py'), 'utf-8')
  const start = src.indexOf('ALLOWLIST: dict')
  const end = src.indexOf('\n}\n', start)
  expect(start).toBeGreaterThan(0)
  const block = src.slice(start, end)
  const result: Record<string, string[]> = {}
  let current: string | null = null
  for (const line of block.split('\n')) {
    const event = line.match(/^ {4}"([a-z_]+)": \{/)
    if (event) { current = event[1]; result[current] = []; continue }
    const field = line.match(/^ {8}"([a-z_]+)":/)
    if (field && current) result[current].push(field[1])
  }
  return result
}

function filesUnder(dir: string, ext: string[]): string[] {
  const out: string[] = []
  for (const name of readdirSync(dir)) {
    if (name === 'node_modules' || name === 'bin' || name === 'obj' || name === 'dist') continue
    const p = join(dir, name)
    if (statSync(p).isDirectory()) out.push(...filesUnder(p, ext))
    else if (ext.some(e => name.endsWith(e))) out.push(p)
  }
  return out
}

// 客户端（Excel 插件 + WPS）真正上报的事件名
function clientEvents(): Set<string> {
  const files = [
    ...filesUnder(join(repo, 'src/DeepExcel.AddIn'), ['.cs']),
    ...filesUnder(join(repo, 'src/DeepExcel.Wps'), ['.js']),
  ]
  const events = new Set<string>()
  for (const f of files) {
    const src = readFileSync(f, 'utf-8')
    for (const m of src.matchAll(/(?:Record|AppendToOutbox)\(\s*"([a-z_]+)"/g)) events.add(m[1])
  }
  return events
}

describe('privacy page matches what is actually collected', () => {
  const allowlist = serverAllowlist()

  it('parses the server allowlist', () => {
    expect(Object.keys(allowlist)).toContain('task_complete')
    expect(allowlist.task_complete).toContain('tool_sequence')
  })

  it('describes every event and field the server accepts', () => {
    for (const [event, fields] of Object.entries(allowlist)) {
      if (TELEMETRY_NOT_SENT.includes(event)) continue
      expect(TELEMETRY_EVENTS[event], `页面缺少事件 ${event}`).toBeDefined()
      for (const field of fields) {
        expect(TELEMETRY_EVENTS[event].fields[field], `页面缺少字段 ${event}.${field}`).toBeTruthy()
      }
    }
  })

  it('does not describe fields the server would drop anyway', () => {
    for (const [event, spec] of Object.entries(TELEMETRY_EVENTS)) {
      expect(allowlist[event], `服务端没有事件 ${event}`).toBeDefined()
      for (const field of Object.keys(spec.fields)) expect(allowlist[event]).toContain(field)
    }
  })

  it('every event the client emits is on the page, and the "not sent" ones really are not sent', () => {
    const emitted = clientEvents()
    expect(emitted.size).toBeGreaterThan(0)
    for (const event of emitted) expect(TELEMETRY_EVENTS[event], `客户端上报了页面没写的事件 ${event}`).toBeDefined()
    for (const event of TELEMETRY_NOT_SENT) expect(emitted.has(event), `${event} 已被客户端上报，页面要改`).toBe(false)
  })

  it('lists the same never-collected categories the server publishes', () => {
    const src = readFileSync(join(repo, 'server/app/routers/telemetry.py'), 'utf-8')
    const published = src.slice(src.indexOf('"never_collected"'), src.indexOf('],', src.indexOf('"never_collected"')))
    const count = (published.match(/"[^"]+",/g) ?? []).length
    expect(NEVER_COLLECTED.length).toBe(count)
  })
})
