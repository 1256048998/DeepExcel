// src/DeepExcel.Wps/starter-host.js
// 首次使用的 WPS 一侧（对应 C# Bridge/StarterHandlers.cs）：
//   get_starter   → 每张可见表的前 51 行 × 30 列 + 哪些列是日期，推荐在面板里算（UI utils/starter.ts）
//   insert_sample → 把面板给的示例数据写进最后一张表之后的新表，不动用户已有的表

'use strict'

const MAX_SHEETS = 12
const MAX_ROWS = 51
const MAX_COLUMNS = 30
const MAX_SAMPLE_ROWS = 500
const MAX_SAMPLE_COLUMNS = 30
const MAX_SAMPLE_TEXT = 200

// CVErr：-2146826288 + (xlErr 常量 - 2000)，与 C# SnapshotEncoder.ErrorText 一致
const ERROR_TEXT = {
  2000: '#NULL!', 2007: '#DIV/0!', 2015: '#VALUE!', 2023: '#REF!', 2029: '#NAME?',
  2036: '#NUM!', 2042: '#N/A', 2045: '#SPILL!', 2050: '#CALC!',
}

function toCell(value) {
  if (value === null || value === undefined) return null
  if (typeof value === 'number') {
    if (!isFinite(value)) return null
    // 错误值在 COM 里是 Int32 的 CVErr，落在这个区间的整数不会是正常数据
    const code = value + 2146828288
    if (Number.isInteger(value) && code >= 2000 && code <= 2100) return ERROR_TEXT[code] || '#ERROR'
    return value
  }
  if (typeof value === 'boolean') return value
  if (value instanceof Date) return value.getTime() / 86400000 + 25569
  const text = String(value)
  if (text.length === 0) return null
  return text.length > MAX_SAMPLE_TEXT ? text.slice(0, MAX_SAMPLE_TEXT) : text
}

/** 与 C# SheetProfiler.LooksLikeDateFormat 同一口径 */
function looksLikeDateFormat(format) {
  if (!format || format === 'General' || format === 'G/通用格式') return false
  const lower = String(format).replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').toLowerCase()
  return lower.indexOf('yy') >= 0 || lower.indexOf('mmm') >= 0 ||
    (lower.indexOf('d') >= 0 && lower.indexOf('m') >= 0) ||
    lower.indexOf('年') >= 0 || lower.indexOf('月') >= 0
}

function as2d(value) {
  if (Array.isArray(value)) return value.map(row => (Array.isArray(row) ? row : [row]))
  return [[value]]
}

function isVisible(ws) {
  try { return ws.Visible === -1 || ws.Visible === true } catch (e) { return true }
}

function eachSheet(wb, fn) {
  const sheets = wb.Worksheets
  const count = sheets.Count
  for (let i = 1; i <= count; i++) fn(sheets.Item(i))
}

function outlineOf(ws) {
  const used = ws.UsedRange
  const rows = used.Rows.Count
  const cols = used.Columns.Count
  const r = Math.min(rows, MAX_ROWS)
  const c = Math.min(cols, MAX_COLUMNS)
  const row1 = used.Row
  const col1 = used.Column
  const block = ws.Range(ws.Cells(row1, col1), ws.Cells(row1 + r - 1, col1 + c - 1))
  const grid = as2d(block.Value2).slice(0, r).map(row => row.slice(0, c).map(toCell))
  const dateColumns = []
  for (let j = 0; j < c; j++) {
    let isDate = false
    if (r >= 2) {
      try { isDate = looksLikeDateFormat(ws.Cells(row1 + 1, col1 + j).NumberFormat) } catch (e) { /* 读不到格式就当不是日期 */ }
    }
    dateColumns.push(isDate)
  }
  return { name: ws.Name, rows, columns: cols, grid, date_columns: dateColumns }
}

function describe(app) {
  const wb = app && app.ActiveWorkbook
  if (!wb) return { workbook_name: '', sheets: [] }
  const sheets = []
  eachSheet(wb, ws => {
    if (sheets.length >= MAX_SHEETS || !isVisible(ws)) return
    sheets.push(outlineOf(ws))
  })
  return { workbook_name: wb.Name || '', sheets }
}

function sanitize(name) {
  let cleaned = String(name || '').replace(/[\\/?*[\]:]/g, '').replace(/^[' ]+|[' ]+$/g, '')
  if (!cleaned) cleaned = '示例'
  return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned
}

function uniqueSheetName(wanted, existing) {
  const taken = new Set((existing || []).map(n => String(n).toLowerCase()))
  const base = sanitize(wanted)
  if (!taken.has(base.toLowerCase())) return base
  for (let i = 2; ; i++) {
    const suffix = '(' + i + ')'
    const candidate = (base.length + suffix.length > 31 ? base.slice(0, 31 - suffix.length) : base) + suffix
    if (!taken.has(candidate.toLowerCase())) return candidate
  }
}

/** 面板给的示例行 → 写入用的二维数组；以 = 开头的文本加撇号，不当公式写入 */
function parseRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) throw new Error('示例数据为空')
  const width = Math.max.apply(null, rows.map(r => (Array.isArray(r) ? r.length : 0)))
  if (rows.length > MAX_SAMPLE_ROWS || width === 0 || width > MAX_SAMPLE_COLUMNS) {
    throw new Error('示例数据的行列数超出范围')
  }
  return rows.map(row => {
    const out = []
    for (let c = 0; c < width; c++) {
      const v = Array.isArray(row) ? row[c] : undefined
      if (typeof v === 'number' || typeof v === 'boolean') out.push(v)
      else if (typeof v === 'string') {
        const text = v.charAt(0) === '=' ? "'" + v : v
        out.push(text.length > MAX_SAMPLE_TEXT ? text.slice(0, MAX_SAMPLE_TEXT) : text)
      } else out.push(null)
    }
    return out
  })
}

function insertSample(app, payload) {
  const wb = app && app.ActiveWorkbook
  if (!wb) throw new Error('没有打开的工作簿')
  if (wb.ProtectStructure) throw new Error('工作簿结构受保护，无法新建工作表')
  const data = parseRows(payload && payload.rows)
  const existing = []
  eachSheet(wb, ws => existing.push(ws.Name))
  const name = uniqueSheetName(payload && payload.sheet_name, existing)

  // 放在最后一张表之后：不改变用户已有表的顺序
  const last = wb.Worksheets.Item(wb.Worksheets.Count)
  const ws = wb.Worksheets.Add(undefined, last)
  ws.Name = name
  const height = data.length
  const width = data[0].length
  ws.Range(ws.Cells(1, 1), ws.Cells(height, width)).Value2 = data
  const formats = (payload && payload.number_formats) || {}
  Object.keys(formats).forEach(key => {
    const col = parseInt(key, 10)
    if (!(col >= 0 && col < width) || typeof formats[key] !== 'string' || height < 2) return
    ws.Range(ws.Cells(2, col + 1), ws.Cells(height, col + 1)).NumberFormat = formats[key]
  })
  ws.Range(ws.Cells(1, 1), ws.Cells(1, width)).Font.Bold = true
  try { ws.Range(ws.Cells(1, 1), ws.Cells(height, width)).Columns.AutoFit() } catch (e) { /* 列宽不影响结果 */ }
  try { ws.Activate() } catch (e) { /* 激活失败不影响结果 */ }
  return name
}

module.exports = { describe, insertSample, uniqueSheetName, parseRows, toCell, looksLikeDateFormat }
