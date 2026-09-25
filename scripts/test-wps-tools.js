'use strict'

// WPS 工具调度：侧车发来的参数（与 C# 端同名）落到正确的 WPS 对象模型调用上。
// 参数名是否一致由 src/DeepExcel.Sidecar/tests/test_wps_tool_args.py 静态守卫；
// 这里看调用本身：插删行列的块、填充的目标区域、清除类型、表格、排序的表头参数。
// 运行：node scripts/test-wps-tools.js（build-wps.ps1 会自动调用）

const assert = require('assert')
const path = require('path')

const calls = []
function rec(what) { calls.push(what) }

function makeRange(label) {
  const range = {
    label,
    Worksheet: null,
    Columns: { Count: 2 },
    EntireRow: { Insert: shift => rec(`${label}.EntireRow.Insert(${shift})`), Delete: () => rec(`${label}.EntireRow.Delete`) },
    EntireColumn: {
      Insert: shift => rec(`${label}.EntireColumn.Insert(${shift})`),
      Delete: () => rec(`${label}.EntireColumn.Delete`),
      AutoFit: () => rec(`${label}.EntireColumn.AutoFit`),
    },
    Resize: (r, c) => makeRange(`${label}.Resize(${r},${c})`),
    AutoFill: (dest, type) => rec(`${label}.AutoFill(${dest.label},${type})`),
    Clear: () => rec(`${label}.Clear`),
    ClearContents: () => rec(`${label}.ClearContents`),
    ClearFormats: () => rec(`${label}.ClearFormats`),
    Cells: (r, c) => makeRange(`${label}.Cells(${r},${c})`),
    Sort: (...args) => rec(`${label}.Sort(${args.map(a => (a && a.label) || String(a)).join(',')})`),
    AutoFilter: (field, criteria) => rec(`${label}.AutoFilter(${field},${criteria})`),
    Copy: dest => rec(`${label}.Copy(${dest.label})`),
    Font: {},
    set ColumnWidth(w) { rec(`${label}.ColumnWidth=${w}`) },
    set HorizontalAlignment(v) { rec(`${label}.HorizontalAlignment=${v}`) },
    set VerticalAlignment(v) { rec(`${label}.VerticalAlignment=${v}`) },
    set WrapText(v) { rec(`${label}.WrapText=${v}`) },
    Address: '$A$1:$B$3',
  }
  Object.defineProperty(range.Font, 'Name', { set(v) { rec(`${label}.Font.Name=${v}`) } })
  range.Worksheet = sheet
  return range
}

const sheet = {
  Name: 'Data',
  Cells: (r, c) => ({ label: `Cells(${r},${c})` }),
  Range: (a, b) => makeRange(typeof a === 'string' ? a : `Range(${a.label},${b.label})`),
  ListObjects: {
    Add: (source, range, linkSource, header) => {
      rec(`ListObjects.Add(${source},${range.label},${linkSource},${header})`)
      const table = { Name: 'Table1', Range: makeRange('tableRange') }
      Object.defineProperty(table, 'Name', { get() { return this._n || 'Table1' }, set(v) { this._n = v; rec(`table.Name=${v}`) } })
      return table
    },
  },
}

global.wps = {
  Application: {
    ActiveSheet: sheet,
    Range: address => makeRange(address),
    ActiveWorkbook: { Worksheets: () => sheet },
  },
}

const ToolDispatcher = require(path.resolve(__dirname, '..', 'src', 'DeepExcel.Wps', 'tool-dispatcher.js'))
const dispatcher = new ToolDispatcher()

async function run(tool, args) {
  calls.length = 0
  const result = await dispatcher.execute(tool, args)
  assert.ok(result.success, `${tool} failed: ${result.error}`)
  return { result, calls: calls.slice() }
}

;(async () => {
  let r = await run('insert_rows', { row: 3, count: 2 })
  assert.deepStrictEqual(r.calls, ['Range(Cells(3,1),Cells(4,1)).EntireRow.Insert(-4121)'])

  r = await run('delete_rows', { row: 5, count: 1 })
  assert.deepStrictEqual(r.calls, ['Range(Cells(5,1),Cells(5,1)).EntireRow.Delete'])

  r = await run('insert_columns', { column: 2, count: 3 })
  assert.deepStrictEqual(r.calls, ['Range(Cells(1,2),Cells(1,4)).EntireColumn.Insert(-4161)'])

  r = await run('delete_columns', { column: 4, count: 1 })
  assert.deepStrictEqual(r.calls, ['Range(Cells(1,4),Cells(1,4)).EntireColumn.Delete'])

  // 起始行 + row_count 行（与 C# 端的写入区域一致）
  r = await run('fill_formula_down', { from_address: 'C2', row_count: 10 })
  assert.deepStrictEqual(r.calls, ['C2.AutoFill(C2.Resize(11,2),0)'])

  r = await run('copy_range', { source_address: 'A1:B3', dest_address: 'D1' })
  assert.deepStrictEqual(r.calls, ['A1:B3.Copy(D1)'])

  for (const [type, expected] of [['all', 'A1.Clear'], ['contents', 'A1.ClearContents'], ['formats', 'A1.ClearFormats']]) {
    r = await run('clear_range', { address: 'A1', clear_type: type })
    assert.deepStrictEqual(r.calls, [expected])
  }

  r = await run('set_column_width', { address: 'B:B', width: 12, auto_fit: true })
  assert.deepStrictEqual(r.calls, ['B:B.EntireColumn.AutoFit'])
  r = await run('set_column_width', { address: 'B:B', width: 12, auto_fit: false })
  assert.deepStrictEqual(r.calls, ['B:B.ColumnWidth=12'])

  r = await run('set_cell_style', { address: 'A1', font_name: '微软雅黑', h_align: 'center', v_align: 'top', wrap_text: true })
  assert.deepStrictEqual(r.calls, ['A1.Font.Name=微软雅黑', 'A1.HorizontalAlignment=-4108', 'A1.VerticalAlignment=-4160', 'A1.WrapText=true'])

  // write_table = 转成表格（ListObject），不是写数据
  r = await run('write_table', { address: 'A1:B3', table_name: '销售表' })
  assert.deepStrictEqual(r.calls, ['ListObjects.Add(1,A1:B3,undefined,1)', 'table.Name=销售表'])
  assert.strictEqual(r.result.data.name, '销售表')

  r = await run('filter_data', { range_address: 'A1:D20', column_index: 2, criteria: '>100' })
  assert.deepStrictEqual(r.calls, ['A1:D20.AutoFilter(2,>100)'])

  // 有表头：Header=xlYes(1)、Key 指向数据第一行；没有表头：xlNo(2)
  r = await run('sort_data', { range_address: 'A1:C9', sort_column: 'B', descending: true, has_header: true })
  assert.deepStrictEqual(r.calls, ['A1:C9.Sort(A1:C9.Cells(2,2),2,undefined,undefined,undefined,undefined,undefined,1,1,false,1,1)'])
  r = await run('sort_data', { range_address: 'A2:C9', sort_column: '1', descending: false, has_header: false })
  assert.deepStrictEqual(r.calls, ['A2:C9.Sort(A2:C9.Cells(1,1),1,undefined,undefined,undefined,undefined,undefined,2,1,false,1,1)'])
  r = await run('sort_data', { range_address: 'A1:C9', sort_column: 'c', descending: false, has_header: true })
  assert.ok(r.calls[0].startsWith('A1:C9.Sort(A1:C9.Cells(2,3),'), r.calls[0])

  console.log('WPS tool dispatch tests passed')
})().catch(error => {
  console.error(error)
  process.exit(1)
})
