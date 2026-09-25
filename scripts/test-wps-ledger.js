'use strict'

// WPS 先读后写 + 读后被改检测（对应 C# ReadLedgerTests / ToolDispatcher 的账本用法）
// 运行：node scripts/test-wps-ledger.js（build-wps.ps1 会自动调用）

const assert = require('assert')
const path = require('path')
const wpsDir = path.resolve(__dirname, '..', 'src', 'DeepExcel.Wps')
const Ledger = require(path.join(wpsDir, 'read-ledger.js'))

// ---- 1. 地址解析与格式：与 C# CellRect 同口径 ----
{
  const r = Ledger.parseRect("'Q1 销售'!$B$2:C4", 'Data')
  assert.deepStrictEqual(r, { sheet: 'Q1 销售', row1: 2, col1: 2, row2: 4, col2: 3 })
  assert.strictEqual(Ledger.toA1(r), "'Q1 销售'!B2:C4")
  assert.strictEqual(Ledger.toA1(Ledger.parseRect('A1', 'Data')), 'Data!A1')
  assert.deepStrictEqual(Ledger.parseRect('B:C', 'S'), { sheet: 'S', row1: 1, col1: 2, row2: Ledger.MAX_ROW, col2: 3 })
  assert.deepStrictEqual(Ledger.parseRect('2:5', 'S'), { sheet: 'S', row1: 2, col1: 1, row2: 5, col2: Ledger.MAX_COLUMN })
  assert.strictEqual(Ledger.parseRect('A1,B2', 'S'), null)
  assert.strictEqual(Ledger.parseRect('销售额', 'S'), null)
  assert.ok(Ledger.overlaps(Ledger.parseRect('A1:C3', 'data'), Ledger.parseRect('C3', 'Data')), 'sheet names compare case-insensitively')
  assert.ok(!Ledger.overlaps(Ledger.parseRect('A1:C3', 'Data'), Ledger.parseRect('D1', 'Data')))
}

// ---- 2. 账本规则 ----
{
  const ledger = new Ledger.ReadLedger()
  const target = Ledger.parseRect('A1:B2', 'S')
  assert.strictEqual(ledger.check(target, () => true).verdict, 'not_read')
  assert.strictEqual(ledger.check(target, () => false).verdict, 'allowed', 'blank area may be written without reading')
  ledger.recordRead(Ledger.parseRect('A1:C10', 'S'))
  assert.strictEqual(ledger.check(target, () => true).verdict, 'allowed')
  ledger.recordUserEdit(Ledger.parseRect('B2', 'S'))
  const stale = ledger.check(target, () => true)
  assert.strictEqual(stale.verdict, 'stale_read')
  assert.deepStrictEqual(stale.changed, ['S!B2'])
  assert.deepStrictEqual(ledger.takeUnreportedUserEdits(), ['S!B2'])
  assert.deepStrictEqual(ledger.takeUnreportedUserEdits(), [], 'each edit is reported once')
  ledger.recordRead(Ledger.parseRect('A1:B2', 'S'))
  assert.strictEqual(ledger.check(target, () => true).verdict, 'allowed', 're-reading clears the staleness')
}

// ---- 3. 调度器：写入前把关、读写记账、用户改动提醒 ----
const content = new Set(['Data!A1', 'Data!B2'])  // 有内容的格
global.wps = {
  Application: {
    ActiveSheet: { Name: 'Data' },
    ActiveWorkbook: { FullName: 'C:\\t\\Book1.xlsx', Name: 'Book1.xlsx' },
  },
}
const WpsActions = require(path.join(wpsDir, 'wps-actions.js'))
const writes = []
WpsActions.rangeHasContent = address => {
  const rect = Ledger.parseRect(address, 'Data')
  for (const cell of content) if (Ledger.overlaps(rect, Ledger.parseRect(cell, 'Data'))) return true
  return false
}
WpsActions.readRangePage = address => ({ address: address.replace(/^.*!/, ''), values: [[1]], formulas: [[1]] })
WpsActions.writeValue = (address, value) => writes.push(['value', address, value])
WpsActions.writeRange = (address, values) => writes.push(['range', address, values])
WpsActions.writeFormula = (address, formula) => writes.push(['formula', address, formula])

const ToolDispatcher = require(path.join(wpsDir, 'tool-dispatcher.js'))
const dispatcher = new ToolDispatcher()
const quiet = console.log
console.log = () => {}

;(async () => {
  // 3.1 有内容但没读过：拒绝，什么都不写
  let r = await dispatcher.execute('write_value', { address: 'A1', value: 'x' })
  assert.strictEqual(r.success, false)
  assert.ok(r.error.includes('还没有读过'), r.error)
  assert.strictEqual(writes.length, 0)

  // 3.2 空白区域：直接写
  r = await dispatcher.execute('write_value', { address: 'H9', value: 'x' })
  assert.strictEqual(r.success, true)

  // 3.3 读过再写：放行；write_range 的目标按二维数组扩展
  r = await dispatcher.execute('read_range', { address: 'A1:C4' })
  assert.strictEqual(r.success, true)
  r = await dispatcher.execute('write_range', { address: 'A1', values: [[1, 2], [3, 4]] })
  assert.strictEqual(r.success, true, r.error)

  // 3.4 用户手动改了 B2：写到那里要先重读；提醒随结果带给模型一次
  dispatcher.recordUserEdit('Data', 'B2')
  r = await dispatcher.execute('write_range', { address: 'A1', values: [[1, 2], [3, 4]] })
  assert.strictEqual(r.success, false)
  assert.ok(r.error.includes('Data!B2'), r.error)
  assert.ok(r.warning && r.warning.includes('用户刚刚手动修改了 Data!B2'), r.warning)
  r = await dispatcher.execute('write_value', { address: 'H9', value: 'y' })
  assert.strictEqual(r.warning, undefined, 'the edit notice is sent once')

  // 3.5 重读之后放行
  await dispatcher.execute('read_range', { address: 'Data!A1:B2' })
  r = await dispatcher.execute('write_range', { address: 'A1', values: [[1, 2], [3, 4]] })
  assert.strictEqual(r.success, true, r.error)

  // 3.6 我们自己的写入期间触发的 SheetChange 不算用户改动
  WpsActions.writeFormula = (address) => { dispatcher.recordUserEdit('Data', address) }
  await dispatcher.execute('write_formula', { address: 'A1', formula: '=1' })
  r = await dispatcher.execute('write_value', { address: 'A1', value: 'z' })
  assert.strictEqual(r.success, true, r.error)
  assert.strictEqual(r.warning, undefined)

  console.log = quiet
  console.log('WPS ledger tests passed')
})().catch(error => {
  console.log = quiet
  console.error(error)
  process.exit(1)
})
