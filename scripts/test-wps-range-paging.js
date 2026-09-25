'use strict'

// range-paging.js 与 read_range / find / list 在 WPS 端的测试：
// 纯逻辑部分与 C# 端 RangePagingTests 断言一致；宿主部分用一个极简的假 WPS 对象模型
// 跑通 tool-dispatcher → wps-actions，确认分页、查找、列举的返回形状与 Excel 端相同。
//
// 运行：node scripts/test-wps-range-paging.js（build-wps.ps1 会自动调用）

const assert = require('assert')
const path = require('path')

const wpsDir = path.resolve(__dirname, '..', 'src', 'DeepExcel.Wps')
const Paging = require(path.join(wpsDir, 'range-paging'))

// ============ 纯逻辑 ============
assert.deepStrictEqual(
  Paging.clip({ row1: 1, col1: 1, row2: 1048576, col2: 1 }, { row1: 1, col1: 1, row2: 450, col2: 3 }),
  { row1: 1, col1: 1, row2: 450, col2: 1 },
)
assert.strictEqual(Paging.clip({ row1: 900, col1: 24, row2: 950, col2: 26 }, { row1: 1, col1: 1, row2: 450, col2: 3 }), null)

let p = Paging.plan(450, 3, 0, null)
assert.deepStrictEqual([p.offset, p.rows, p.nextOffset], [0, 200, 200])
p = Paging.plan(450, 3, 400, null)
assert.deepStrictEqual([p.rows, p.nextOffset], [50, null])
assert.strictEqual(Paging.plan(1000, 3, 0, 9999).rows, 500, 'limit 上限 500')
assert.strictEqual(Paging.plan(10000, 100, 0, 500).rows, 100, '宽表按 1 万格封顶')
p = Paging.plan(10, 300, 0, null)
assert.strictEqual(p.columns, 256)
assert.strictEqual(p.columnsTruncated, true)
assert.ok(Paging.plan(10, 3, 10, null).error.includes('最大 9'))
assert.strictEqual(Paging.plan(10, 3, -5, null).offset, 0)

assert.ok(Paging.hint('Data!A1:C450', 450, 3, Paging.plan(450, 3, 0, null)).includes('read_range(address="Data!A1:C450", offset=200)'))
assert.strictEqual(Paging.hint('A1:C10', 10, 3, Paging.plan(10, 3, 0, null)), null)
assert.ok(Paging.hint('A1:KZ10', 10, 312, Paging.plan(10, 312, 0, null)).includes('只返回了前 256 列'))

assert.strictEqual(Paging.qualifiedAddress('Data', 'A1'), 'Data!A1')
assert.strictEqual(Paging.qualifiedAddress('Q1 销售', 'A1'), "'Q1 销售'!A1")
assert.strictEqual(Paging.qualifiedAddress("Bob's", 'A1'), "'Bob''s'!A1")

// collectMatches：转一圈停、跨表累计、超过上限只计数
const ring = ['A1', 'A5', 'B2']
const state = { matches: [], total: 0, truncated: false, maxResults: 2 }
const counted = Paging.collectMatches('A1', c => ring[(ring.indexOf(c) + 1) % ring.length], c => c, c => c, state)
assert.deepStrictEqual(counted, { count: 3, complete: true })
assert.deepStrictEqual(state.matches, ['A1', 'A5'])
assert.strictEqual(state.total, 3)
assert.strictEqual(state.truncated, true)
assert.deepStrictEqual(Paging.collectMatches(null, null, null, null, state), { count: 0, complete: true })

assert.ok(Paging.findHint(0, 0, false, true).includes('scope=values'))
assert.ok(Paging.findHint(30, 5, true, false).includes('共找到 30 处'))
assert.strictEqual(Paging.findHint(3, 3, false, false), null)
assert.strictEqual(Paging.clipText(null), '')
assert.strictEqual(Paging.clipText('x'.repeat(130)).length, 121)

// ============ 假 WPS 对象模型 ============
function colLetters(n) {
  let s = ''
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) }
  return s
}
function colNumber(letters) {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}
function parseA1(text) {
  const a = text.replace(/\$/g, '').toUpperCase()
  let m = /^([A-Z]+):([A-Z]+)$/.exec(a)
  if (m) return { row1: 1, col1: colNumber(m[1]), row2: 1048576, col2: colNumber(m[2]) }
  m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a)
  if (!m) throw new Error('bad address ' + text)
  const row1 = +m[2]; const col1 = colNumber(m[1])
  return { row1, col1, row2: m[4] ? +m[4] : row1, col2: m[3] ? colNumber(m[3]) : col1 }
}

function makeSheet(name, index, cells) {
  // cells: { 'A1': { value, formula } }
  const sheet = { Name: name, Index: index, Visible: -1, ProtectContents: false }
  const cellAt = (r, c) => cells[colLetters(c) + r] || null
  const makeRange = (box) => {
    const rows = box.row2 - box.row1 + 1
    const cols = box.col2 - box.col1 + 1
    const grid = (pick) => {
      const out = []
      for (let r = box.row1; r <= box.row2; r++) {
        const row = []
        for (let c = box.col1; c <= box.col2; c++) row.push(pick(cellAt(r, c)))
        out.push(row)
      }
      return rows === 1 && cols === 1 ? out[0][0] : out
    }
    const address = (abs) => {
      const one = (r, c) => abs ? `$${colLetters(c)}$${r}` : `${colLetters(c)}${r}`
      return rows === 1 && cols === 1 ? one(box.row1, box.col1) : one(box.row1, box.col1) + ':' + one(box.row2, box.col2)
    }
    // 只在区域内按行序查找：返回所有匹配的格子
    const hits = (what, inFormulas, whole) => {
      const out = []
      for (let r = box.row1; r <= box.row2; r++) {
        for (let c = box.col1; c <= box.col2; c++) {
          const cell = cellAt(r, c)
          if (!cell) continue
          const text = String(inFormulas && cell.formula ? cell.formula : cell.value).toLowerCase()
          const q = String(what).toLowerCase()
          if (whole ? text === q : text.includes(q)) out.push([r, c])
        }
      }
      return out
    }
    let lastHits = []
    const range = {
      Row: box.row1, Column: box.col1,
      Rows: { Count: rows }, Columns: Object.assign((c) => makeRange({ row1: box.row1, col1: box.col1 + c - 1, row2: box.row2, col2: box.col1 + c - 1 }), { Count: cols }),
      Areas: { Count: 1 },
      Worksheet: sheet,
      Address: (ra, ca) => address(ra !== false),
      get Value2() { return grid(cell => (cell ? cell.value : null)) },
      get Formula() { return grid(cell => (cell ? (cell.formula || String(cell.value)) : '')) },
      NumberFormat: 'General',
      get Text() { const cell = cellAt(box.row1, box.col1); return cell ? String(cell.value) : '' },
      get HasFormula() { const cell = cellAt(box.row1, box.col1); return !!(cell && cell.formula) },
      Cells: (r, c) => makeRange({ row1: box.row1 + r - 1, col1: box.col1 + c - 1, row2: box.row1 + r - 1, col2: box.col1 + c - 1 }),
      Find(what, after, lookIn, lookAt) {
        lastHits = hits(what, lookIn === -4123, lookAt === 1)
        if (!lastHits.length) return null
        const [r, c] = lastHits[0]
        return makeRange({ row1: r, col1: c, row2: r, col2: c })
      },
      FindNext(after) {
        const i = lastHits.findIndex(([r, c]) => r === after.Row && c === after.Column)
        const [r, c] = lastHits[(i + 1) % lastHits.length]
        return makeRange({ row1: r, col1: c, row2: r, col2: c })
      },
    }
    return range
  }
  const keys = Object.keys(cells).map(parseA1)
  const usedBox = keys.length
    ? { row1: Math.min(...keys.map(k => k.row1)), col1: Math.min(...keys.map(k => k.col1)), row2: Math.max(...keys.map(k => k.row1)), col2: Math.max(...keys.map(k => k.col1)) }
    : { row1: 1, col1: 1, row2: 1, col2: 1 }
  sheet.UsedRange = makeRange(usedBox)
  sheet.Cells = (r, c) => makeRange({ row1: r, col1: c, row2: r, col2: c })
  sheet.Range = (a, b) => {
    if (b) return makeRange({ row1: a.Row, col1: a.Column, row2: b.Row, col2: b.Column })
    return makeRange(parseA1(a))
  }
  sheet.ListObjects = { Count: 0 }
  sheet.PivotTables = () => ({ Count: 0 })
  sheet.ChartObjects = () => ({ Count: 0 })
  return sheet
}

const dataCells = {}
for (let r = 1; r <= 450; r++) {
  dataCells['A' + r] = { value: r }
  dataCells['B' + r] = { value: r === 3 ? '应收账款' : r === 9 ? '其他应收款' : 'x' + r }
  dataCells['C' + r] = { value: r * 2, formula: '=A' + r + '*2' }
}
const data = makeSheet('Data', 1, dataCells)
const sum = makeSheet('Q1 汇总', 2, { C2: { value: 42, formula: '=SUMIF(Data!B:B,"应收账款",Data!C:C)' } })
sum.Visible = 0
const sheets = [data, sum]
const names = [{ Name: '税率', RefersTo: "='Q1 汇总'!$B$1", Visible: true }, { Name: '坏名', RefersTo: '=#REF!$A$1', Visible: true }]
const workbook = {
  Name: 'book.xlsx',
  Worksheets: Object.assign((key) => (typeof key === 'number' ? sheets[key - 1] : sheets.find(s => s.Name === key)), { Count: 2 }),
  Names: { Count: names.length, Item: (i) => names[i - 1] },
}
global.wps = { Application: { ActiveWorkbook: workbook, Range: (a) => data.Range(a) } }

const ToolDispatcher = require(path.join(wpsDir, 'tool-dispatcher'))
const dispatcher = new ToolDispatcher()

;(async () => {
  // read_range：整列收到已用区域、分页、提示下一页
  let r = await dispatcher.execute('read_range', { address: 'Data!A:C' })
  assert.strictEqual(r.success, true)
  assert.strictEqual(r.data.rowCount, 200)
  assert.strictEqual(r.data.values.length, 200)
  assert.strictEqual(r.data.paging.total_rows, 450)
  assert.strictEqual(r.data.paging.next_offset, 200)
  assert.strictEqual(r.data.paging.clipped_to_used_range, true)
  assert.ok(r.data.hint.includes('offset=200'))
  assert.strictEqual(r.data.numberFormats[0][0], 'General')

  r = await dispatcher.execute('read_range', { address: 'Data!A:C', offset: '400' })
  assert.strictEqual(r.data.rowCount, 50)
  assert.strictEqual(r.data.paging.next_offset, null)
  assert.strictEqual(r.data.address, '$A$401:$C$450')
  assert.strictEqual(r.data.values[0][0], 401)

  r = await dispatcher.execute('read_range', { address: 'Data!A1:C450', offset: 450 })
  assert.strictEqual(r.success, false)
  assert.ok(r.error.includes('超出范围'))

  r = await dispatcher.execute('read_range', { address: 'Data!X900:Z950' })
  assert.ok(r.data.hint.includes('全是空白'))

  r = await dispatcher.execute('read_range', { address: "'Q1 汇总'!C2" })
  assert.strictEqual(r.success, true)
  assert.deepStrictEqual(r.data.values, [[42]], '单格也是二维数组')

  // find：跨表、值 / 公式、整格匹配、限定表、截断
  r = await dispatcher.execute('find', { query: '应收' })
  assert.strictEqual(r.success, true)
  assert.strictEqual(r.data.total, 2)
  assert.deepStrictEqual(r.data.matches.map(m => m.sheet + '!' + m.address), ['Data!B3', 'Data!B9'])

  r = await dispatcher.execute('find', { query: '应收账款', match: 'exact' })
  assert.strictEqual(r.data.total, 1)

  r = await dispatcher.execute('find', { query: 'SUMIF', scope: 'formulas' })
  assert.strictEqual(r.data.matches[0].sheet, 'Q1 汇总')
  assert.ok(r.data.matches[0].formula.includes('SUMIF'))

  r = await dispatcher.execute('find', { query: '应收', sheets: 'Q1 汇总' })
  assert.strictEqual(r.data.total, 0)
  assert.ok(r.data.hint.includes('没有找到'))

  r = await dispatcher.execute('find', { query: 'x1', max_results: 5 })
  assert.strictEqual(r.data.returned, 5)
  assert.ok(r.data.total > 5)
  assert.strictEqual(r.data.truncated, true)

  r = await dispatcher.execute('find', { query: '应收', sheets: ['不存在'] })
  assert.strictEqual(r.success, false)
  assert.ok(r.suggestion.includes('list(kind=sheets)'))

  r = await dispatcher.execute('find', { query: '' })
  assert.strictEqual(r.success, false)

  // list
  r = await dispatcher.execute('list', { kind: 'sheets' })
  assert.deepStrictEqual(r.data.items.map(s => [s.name, s.visibility]), [['Data', 'visible'], ['Q1 汇总', 'hidden']])
  assert.strictEqual(r.data.items[0].used_range, 'A1:C450')

  r = await dispatcher.execute('list', { kind: 'names' })
  assert.deepStrictEqual(r.data.items.map(n => [n.name, n.broken]), [['税率', false], ['坏名', true]])

  r = await dispatcher.execute('list', {})
  assert.strictEqual(r.data.kind, 'sheets')

  r = await dispatcher.execute('list', { kind: 'widgets' })
  assert.strictEqual(r.success, false)

  console.log('WPS range-paging / find / list tests passed')
})().catch(err => {
  console.error(err)
  process.exit(1)
})
