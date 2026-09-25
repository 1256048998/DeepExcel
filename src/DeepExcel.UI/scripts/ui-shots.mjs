// 截图回归：npm run ui:shots
//
// 用本机的 Edge（和 Excel 里的 WebView2 同一个内核）打开 vite dev 下的面板，
// 通过 window.__deepexcelDevHost 注入宿主消息，把面板摆成几个主要场景，
// 在 320 / 400 / 480 三个宽度下各截一张图，并自动检查：
//   - 横向溢出（任何可见元素的右边缘超出面板宽度）
//   - 控制台报错
// 有问题时退出码为 1。截图写到 shots/（不入库），供每次界面改动后人工过一眼。
//
//   npm run ui:shots                 全部场景
//   npm run ui:shots -- conversation 只跑名字里含 conversation 的场景

import { createServer } from 'vite'
import { chromium } from 'playwright-core'
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const outDir = join(root, 'shots')
const WIDTHS = [320, 400, 480]
const HEIGHT = 760
const filter = process.argv[2] || ''

// ---------------- 场景 ----------------
// 每个场景是在页面里执行的函数：用 emit(type, payload) 模拟宿主消息，
// send(text) 模拟用户在输入框里发一条消息。

const ev = (kind, fields) => ({ v: 1, kind, ts: Date.now(), ...fields })

const scenes = {
  welcome: async () => {},

  conversation: async ({ emit, send }) => {
    await send('把 C 列金额清洗成数字，再在 D 列算单价')
    emit('ui_event', ev('tool_start', { id: 't1', name: 'read_range', args: { address: 'A1:C40' } }))
    emit('ui_event', ev('tool_end', { id: 't1', name: 'read_range', ok: true, duration_ms: 180, summary: '40 行 × 3 列' }))
    emit('ui_event', ev('tool_start', { id: 't2', name: 'clean_amount', args: { range_address: 'C2:C40' } }))
    emit('ui_event', ev('tool_end', {
      id: 't2', name: 'clean_amount', ok: true, duration_ms: 640, checkpoint_id: 'cp-2', summary: '清洗了 37 个金额',
      changes: { changed: 37, sheet: 'Sheet1', samples: [
        { address: 'C2', before: '¥1,200.00', after: '1200' },
        { address: 'C3', before: '3,450元', after: '3450' },
        { address: 'C4', before: '', after: '0' },
      ] },
    }))
    emit('ui_event', ev('tool_start', { id: 't3', name: 'fill_formula_down', args: { from_address: 'D2', row_count: 38 } }))
    emit('ui_event', ev('tool_end', {
      id: 't3', name: 'fill_formula_down', ok: true, duration_ms: 420, checkpoint_id: 'cp-3',
      check: { ok: false, summary: '体检发现问题：新增 2 个公式错误：Sheet1!D17 #DIV/0!、Sheet1!D23 #DIV/0!。先查明原因并修正，不要直接宣布完成' },
    }))
    emit('ui_event', ev('tool_start', { id: 't4', name: 'write_value', args: { address: 'Z1', value: 'x' } }))
    emit('ui_event', ev('tool_end', {
      id: 't4', name: 'write_value', ok: false, duration_ms: 90,
      error: { code: 'tool_failed', message: '目标区域 Sheet1!Z1 已有内容，但你还没有读过它，本次未写入。', hint: '先用 read_range 读取 Sheet1!Z1' },
    }))
    const answer = [
      '已完成：',
      '',
      '1. **C 列**：37 个金额从文本转成了数字（去掉了 ¥、逗号和「元」）。',
      '2. **D 列**：填入 `=C2/B2` 计算单价，共 38 行。',
      '',
      '| 行 | 数量 | 金额 | 单价 |',
      '|---|---|---|---|',
      '| 2 | 10 | 1200 | 120 |',
      '| 17 | 0 | 560 | #DIV/0! |',
      '',
      '第 17、23 行数量为 0，单价出现 #DIV/0!，要不要改成 `=IF(B2=0,"",C2/B2)`？',
    ].join('\n')
    emit('stream_delta', { delta: answer })
    emit('ui_event', ev('run_summary', { outcome: 'success', tool_calls: 4, failed_calls: 1, duration_ms: 38000, num_turns: 5 }))
    emit('stream_end', {})
  },

  running: async ({ emit, send }) => {
    await send('按部门汇总销售额，生成透视表和柱状图')
    emit('ui_event', ev('plan', { items: [
      { content: '读取数据范围', status: 'completed' },
      { content: '创建透视表', status: 'in_progress' },
      { content: '插入柱状图', status: 'pending' },
    ] }))
    emit('ui_event', ev('tool_start', { id: 'r1', name: 'read_workbook', args: {} }))
    emit('ui_event', ev('tool_end', { id: 'r1', name: 'read_workbook', ok: true, duration_ms: 120 }))
    const code = 'Sub BuildPivot()\n    Dim ws As Worksheet\n    Set ws = Worksheets("数据")\n    Dim pc As PivotCache\n    Set pc = ActiveWorkbook.PivotCaches.Create( _\n        SourceType:=xlDatabase, SourceData:=ws.Range("A1:F500"))\n'
    emit('ui_event', ev('tool_gen', { id: 'r2', name: 'execute_vba', chars: code.length, lines: 6, preview: code }))
    emit('ui_event', ev('status', { text: '仍在等待模型响应（25 秒）…' }))
  },

  error: async ({ emit, send }) => {
    await send('帮我把这张表按日期排序')
    emit('ui_event', ev('error', { code: 'auth', message: '模型服务拒绝了凭据。', hint: '到「模型设置」检查 API Key 是否正确、是否过期。', retryable: false }))
    emit('ui_event', ev('run_summary', { outcome: 'error', tool_calls: 0, failed_calls: 0, duration_ms: 2100 }))
    emit('stream_end', {})
  },

  permission: async ({ emit, send }) => {
    await send('删除所有空行')
    emit('permission_request', {
      request_id: 'p1', tool: 'delete_rows', args: { start_row: 5, count: 12 },
      preview: null,
    })
  },

  setup: { mock: 'nokey', run: async () => {} },

  menu: async ({ page }) => {
    await page.click('[aria-label="更多"]')
  },
}

// ---------------- 执行 ----------------

function sceneEntries() {
  return Object.entries(scenes)
    .filter(([name]) => name.includes(filter))
    .map(([name, s]) => (typeof s === 'function' ? { name, mock: 'default', run: s } : { name, ...s }))
}

async function runScene(page, scene) {
  await page.evaluate(() => { window.__deepexcelDevHost.silent = true })
  const send = async text => {
    await page.fill('textarea', text)
    await page.keyboard.press('Enter')
    await page.waitForTimeout(120)
  }
  const emit = (type, payload) => page.evaluate(([t, p]) => window.__deepexcelDevHost.emit(t, p), [type, payload])
  // 场景函数在 Node 里跑，逐条把消息送进页面，和真实宿主一样是一条一条到达的
  await scene.run({ emit: (t, p) => emit(t, p), send, page })
  await page.waitForTimeout(300)
}

async function findOverflow(page) {
  return page.evaluate(() => {
    const width = document.documentElement.clientWidth
    const problems = []
    for (const el of document.querySelectorAll('body *')) {
      const style = getComputedStyle(el)
      if (style.display === 'none' || style.visibility === 'hidden' || style.position === 'fixed') continue
      const rect = el.getBoundingClientRect()
      if (rect.width === 0 || rect.height === 0) continue
      // 被祖先裁掉（overflow hidden / auto）的不算
      let clipped = false
      for (let p = el.parentElement; p && p !== document.body; p = p.parentElement) {
        const ps = getComputedStyle(p)
        if (ps.overflowX !== 'visible') {
          const pr = p.getBoundingClientRect()
          if (pr.right <= width + 0.5) { clipped = true; break }
        }
      }
      if (!clipped && rect.right > width + 0.5) {
        const name = el.tagName.toLowerCase() + (el.className && typeof el.className === 'string' ? '.' + el.className.trim().split(/\s+/).join('.') : '')
        problems.push(`${name} 右边缘 ${Math.round(rect.right)}px > ${width}px`)
      }
    }
    // 按钮 / 输入框被祖先裁掉一部分：看得见却点不全（顶栏按钮挤出面板就是这种）
    for (const el of document.querySelectorAll('button, select, input, textarea, [role="button"]')) {
      const style = getComputedStyle(el)
      if (style.display === 'none' || style.visibility === 'hidden') continue
      const rect = el.getBoundingClientRect()
      if (rect.width === 0 || rect.height === 0) continue
      if (el.closest('pre, table, .message-list, [class*="overlay"], [class*="panel"]')) continue
      if (rect.left < -0.5 || rect.right > width + 0.5) {
        const label = (el.getAttribute('title') || el.getAttribute('aria-label') || el.textContent || el.tagName).trim().slice(0, 12)
        problems.push(`控件「${label}」超出面板（${Math.round(rect.left)}–${Math.round(rect.right)}px）`)
      }
    }
    if (document.documentElement.scrollWidth > width) problems.push(`页面可横向滚动（scrollWidth ${document.documentElement.scrollWidth}）`)
    return [...new Set(problems)].slice(0, 8)
  })
}

async function main() {
  mkdirSync(outDir, { recursive: true })
  const server = await createServer({ root, server: { port: 5199, strictPort: false }, logLevel: 'error' })
  await server.listen()
  const base = server.resolvedUrls.local[0]
  const browser = await chromium.launch({ channel: 'msedge', headless: true })
  const report = []
  let failed = 0
  try {
    for (const scene of sceneEntries()) {
      for (const width of WIDTHS) {
        const page = await browser.newPage({ viewport: { width, height: HEIGHT }, deviceScaleFactor: 1, reducedMotion: 'reduce' })
        const errors = []
        // 资源 404 按 URL 报（控制台那条不带地址）；favicon 不算
        page.on('console', msg => { if (msg.type() === 'error' && !msg.text().startsWith('Failed to load resource')) errors.push(msg.text()) })
        page.on('response', res => { if (res.status() >= 400 && !res.url().includes('favicon')) errors.push(`${res.status()} ${res.url()}`) })
        page.on('pageerror', err => errors.push(String(err)))
        await page.goto(`${base}?mock=${scene.mock}`)
        await page.waitForSelector('textarea')
        await page.addStyleTag({ content: '*, *::before, *::after { animation: none !important; transition: none !important; caret-color: transparent !important; }' })
        await page.waitForTimeout(300)
        await runScene(page, scene)
        const file = join(outDir, `${scene.name}-${width}.png`)
        await page.screenshot({ path: file })
        const overflow = await findOverflow(page)
        const ok = overflow.length === 0 && errors.length === 0
        if (!ok) failed++
        report.push({ scene: scene.name, width, ok, overflow, errors })
        console.log(`${ok ? 'ok  ' : 'FAIL'} ${scene.name} @${width}px${overflow.length ? '  溢出: ' + overflow.join('；') : ''}${errors.length ? '  报错: ' + errors.join('；') : ''}`)
        await page.close()
      }
    }
  } finally {
    await browser.close()
    await server.close()
  }
  writeFileSync(join(outDir, 'report.json'), JSON.stringify(report, null, 2))
  console.log(`\n${report.length - failed}/${report.length} 通过，截图在 ${outDir}`)
  process.exit(failed > 0 ? 1 : 0)
}

main().catch(err => {
  console.error(err)
  process.exit(1)
})
