// scripts/test-wps-memory-store.js
// WPS 端工作簿记忆存取：目录口径必须与侧车（workbook_memory.py）和 C#（WorkbookMemoryStore）一致，
// 否则同一个文件在两个宿主里各记各的。期望值由侧车 key_hash 算出。
const assert = require('assert')
const fs = require('fs')
const os = require('os')
const path = require('path')

const root = fs.mkdtempSync(path.join(os.tmpdir(), 'deepexcel-memory-'))
process.env.DEEPEXCEL_MEMORY_DIR = root
const store = require('../src/DeepExcel.Wps/workbook-memory-store')

const KEY = 'C:\\财务\\2026 预算.xlsx'
assert.strictEqual(store.keyHash(KEY), '745f66b0b53be193d2368fd7504c8f6a')

// 没保存过的工作簿没有记忆
assert.strictEqual(store.describe('工作簿1', '工作簿1').available, false)
assert.ok(store.save('工作簿1', '工作簿1', 'x'))

// 保存 / 读取 / 上限
assert.strictEqual(store.save(KEY, '2026 预算.xlsx', '## 禁区\r\n- 汇总\r\n'), null)
let state = store.describe(KEY, '2026 预算.xlsx')
assert.strictEqual(state.available, true)
assert.strictEqual(state.notes, '## 禁区\n- 汇总\n')
assert.ok(store.save(KEY, 'x', 'y'.repeat(store.MAX_NOTES_CHARS + 1)).includes('上限'))

// 操作记录计数（侧车写的 history.jsonl）
fs.writeFileSync(path.join(store.dirFor(KEY), 'history.jsonl'), '{"tool":"a"}\n{"tool":"b"}\n', 'utf8')
assert.strictEqual(store.describe(KEY, 'x').historyCount, 2)

// 清除
store.clear(KEY)
state = store.describe(KEY, 'x')
assert.strictEqual(state.notes, '')
assert.strictEqual(state.historyCount, 0)
assert.ok(!fs.existsSync(store.dirFor(KEY)))

fs.rmSync(root, { recursive: true, force: true })
console.log('WPS workbook memory store tests passed')
