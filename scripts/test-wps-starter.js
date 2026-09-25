'use strict'

// WPS 首次使用：get_starter 读出的结构、insert_sample 写入的新表（对应 C# StarterTests）
// 运行：node scripts/test-wps-starter.js（build-wps.ps1 会自动调用）

const assert = require('assert')
const path = require('path')
const starter = require(path.resolve(__dirname, '..', 'src', 'DeepExcel.Wps', 'starter-host.js'))

// ---- 极简的 WPS 表格对象模型：够 starter-host 用就行 ----
function makeSheet(name, cells) {
  const ws = { Name: name, Visible: -1, cells: cells || {}, formats: {}, bold: false, activated: false }
  const key = (r, c) => r + ',' + c
  ws.Cells = (r, c) => ({
    r, c,
    get NumberFormat() { return ws.formats[key(r, c)] || 'General' },
  })
  ws.Range = (a, b) => ({
    set Value2(data) {
      data.forEach((row, i) => row.forEach((v, j) => {
        // 与 Excel 一致：撇号开头按文本写入，撇号本身不进值
        const stored = typeof v === 'string' && v.charAt(0) === "'" ? v.slice(1) : v
        ws.cells[key(a.r + i, a.c + j)] = stored === '' ? null : stored
      }))
    },
    get Value2() {
      if (a.r === b.r && a.c === b.c) return ws.cells[key(a.r, a.c)] ?? null
      const out = []
      for (let r = a.r; r <= b.r; r++) {
        const row = []
        for (let c = a.c; c <= b.c; c++) row.push(ws.cells[key(r, c)] ?? null)
        out.push(row)
      }
      return out
    },
    set NumberFormat(f) { for (let r = a.r; r <= b.r; r++) for (let c = a.c; c <= b.c; c++) ws.formats[key(r, c)] = f },
    Font: { set Bold(v) { ws.bold = v } },
    Columns: { AutoFit() {} },
  })
  Object.defineProperty(ws, 'UsedRange', {
    get() {
      const keys = Object.keys(ws.cells).filter(k => ws.cells[k] !== null && ws.cells[k] !== undefined)
      if (keys.length === 0) return { Row: 1, Column: 1, Rows: { Count: 1 }, Columns: { Count: 1 } }
      const rs = keys.map(k => +k.split(',')[0]), cs = keys.map(k => +k.split(',')[1])
      const r1 = Math.min(...rs), c1 = Math.min(...cs)
      return { Row: r1, Column: c1, Rows: { Count: Math.max(...rs) - r1 + 1 }, Columns: { Count: Math.max(...cs) - c1 + 1 } }
    },
  })
  ws.Activate = () => { ws.activated = true }
  return ws
}

function makeWorkbook(sheets) {
  const list = sheets.slice()
  const wb = {
    Name: 'Book1.xlsx',
    ProtectStructure: false,
    list,
    Worksheets: {
      get Count() { return list.length },
      Item: i => list[i - 1],
      Add(before, after) {
        const ws = makeSheet('Sheet' + (list.length + 1))
        const at = after ? list.indexOf(after) + 1 : 0
        list.splice(at, 0, ws)
        return ws
      },
    },
  }
  return wb
}

// 1. 空工作簿：一张空表
{
  const app = { ActiveWorkbook: makeWorkbook([makeSheet('Sheet1')]) }
  const outline = starter.describe(app)
  assert.strictEqual(outline.workbook_name, 'Book1.xlsx')
  assert.strictEqual(outline.sheets.length, 1)
  assert.deepStrictEqual(outline.sheets[0].grid, [[null]])
}

// 2. 插入示例：放在最后，重名加序号，文本金额留在文本，空区域是空格
{
  const user = makeSheet('Data', { '1,1': '用户的数据' })
  const wb = makeWorkbook([user, makeSheet('示例-销售明细')])
  const app = { ActiveWorkbook: wb }
  const rows = [['订单日期', '区域', '金额'], [45474, '华东', 100], [45475, '华南', "'1,280"], [45476, '', '=1+1']]
  const name = starter.insertSample(app, { sheet_name: '示例-销售明细', rows, number_formats: { 0: 'yyyy-mm-dd' } })
  assert.strictEqual(name, '示例-销售明细(2)')
  assert.strictEqual(wb.list[wb.list.length - 1].Name, name, 'sample goes after the last sheet')
  assert.strictEqual(user.cells['1,1'], '用户的数据', 'user sheet untouched')
  const sample = wb.list[wb.list.length - 1]
  assert.strictEqual(sample.cells['3,3'], '1,280')
  assert.strictEqual(sample.cells['4,2'], null)
  assert.strictEqual(sample.cells['4,3'], '=1+1', 'formula text is written as text, not a formula')
  assert.ok(sample.bold && sample.activated)

  const outline = starter.describe(app)
  const sheet = outline.sheets.find(s => s.name === name)
  assert.strictEqual(sheet.rows, 4)
  assert.deepStrictEqual(sheet.date_columns, [true, false, false])
  assert.strictEqual(sheet.grid[2][2], '1,280')
}

// 3. 受保护的结构、超大示例都拒绝
{
  const wb = makeWorkbook([makeSheet('Sheet1')])
  wb.ProtectStructure = true
  assert.throws(() => starter.insertSample({ ActiveWorkbook: wb }, { sheet_name: 'x', rows: [['a']] }), /受保护/)
  const wide = [Array.from({ length: 31 }, (_, i) => i)]
  assert.throws(() => starter.parseRows(wide), /超出范围/)
  assert.throws(() => starter.parseRows([]), /为空/)
}

// 4. 小工具与 C# 同口径
assert.strictEqual(starter.toCell(-2146826246), '#N/A')
assert.strictEqual(starter.toCell(-2146826281), '#DIV/0!')
assert.strictEqual(starter.toCell(''), null)
assert.strictEqual(starter.looksLikeDateFormat('yyyy/m/d'), true)
assert.strictEqual(starter.looksLikeDateFormat('0.00'), false)
assert.strictEqual(starter.looksLikeDateFormat('"日"0'), false)
assert.strictEqual(starter.uniqueSheetName('a/b', []), 'ab')
assert.ok(starter.uniqueSheetName('表'.repeat(40), ['表'.repeat(31)]).length <= 31)

console.log('WPS starter tests passed')
