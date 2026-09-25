// src/DeepExcel.Wps/range-paging.js
// read_range 分页、find 汇总、sheet_snapshot 编码的纯逻辑（对应 C# 端 Perception/RangePaging.cs、
// ExcelActionsImpl.FindCells 的计数规则、Perception/SnapshotEncoder.cs）。不碰 wps 对象，便于 node 直接测试：
// node scripts/test-wps-range-paging.js（build-wps.ps1 会自动调用）

'use strict'

const DEFAULT_ROWS = 200
const MAX_ROWS = 500
// 一页最多这么多格：值 + 公式 + 数字格式序列化后要远低于 SDK 的 1MB 消息上限
const MAX_CELLS = 10000
const MAX_COLUMNS = 256
// 每张表最多数这么多个匹配：FindNext 在 WPS 里是同步调用，数一百万个会卡住界面
const MAX_FIND_SCAN_PER_SHEET = 5000

/** 两个矩形（{row1,col1,row2,col2}，1 起始含两端）的交集；没有交集返回 null */
function clip(requested, used) {
  const row1 = Math.max(requested.row1, used.row1)
  const col1 = Math.max(requested.col1, used.col1)
  const row2 = Math.min(requested.row2, used.row2)
  const col2 = Math.min(requested.col2, used.col2)
  if (row1 > row2 || col1 > col2) return null
  return { row1, col1, row2, col2 }
}

function boxRows(box) { return box.row2 - box.row1 + 1 }
function boxColumns(box) { return box.col2 - box.col1 + 1 }

/** 一页读取计划：从第 offset 行起读 rows 行、columns 列 */
function plan(totalRows, totalColumns, offset, limit) {
  offset = offset > 0 ? Math.floor(offset) : 0
  if (totalRows <= 0 || totalColumns <= 0) {
    return { offset: 0, rows: 0, columns: 0, nextOffset: null, columnsTruncated: false }
  }
  if (offset >= totalRows) {
    return {
      error: `offset=${offset} 超出范围：这块区域一共 ${totalRows} 行（offset 从 0 开始，最大 ${totalRows - 1}）`,
    }
  }
  const columns = Math.min(totalColumns, MAX_COLUMNS)
  const wanted = limit > 0 ? Math.min(Math.floor(limit), MAX_ROWS) : DEFAULT_ROWS
  const byCells = Math.max(1, Math.floor(MAX_CELLS / columns))
  const rows = Math.min(wanted, byCells, totalRows - offset)
  const next = offset + rows
  return {
    offset,
    rows,
    columns,
    nextOffset: next < totalRows ? next : null,
    columnsTruncated: columns < totalColumns,
  }
}

/** 给模型的一句话：读到了哪里、下一页怎么读。读完了且没截列时返回 null */
function hint(address, totalRows, totalColumns, p) {
  if (!p || !p.rows) return null
  let text = null
  if (p.nextOffset != null) {
    text = `共 ${totalRows} 行，本页是第 ${p.offset + 1}–${p.offset + p.rows} 行，还有 ${totalRows - p.nextOffset} 行未读。` +
      `下一页：read_range(address="${address}", offset=${p.nextOffset})`
  }
  if (p.columnsTruncated) {
    const cols = `这块区域有 ${totalColumns} 列，只返回了前 ${p.columns} 列；其余列请用更窄的地址读取。`
    text = text == null ? cols : text + ' ' + cols
  }
  return text
}

/** 表名里有空格、连字符、叹号或单引号时加单引号（与 C# ErrorCell.QualifiedAddress 一致） */
function qualifiedAddress(sheet, address) {
  if (!sheet) return address
  const quoted = /[ \-!']/.test(sheet) ? `'${sheet.replace(/'/g, "''")}'` : sheet
  return quoted + '!' + address
}

/**
 * 按 Find / FindNext 的约定数一张表上的匹配：从 first 开始，每次 next(cell) 取下一个，
 * 地址回到 first 就是转完了一圈。toMatch(cell) 把单元格变成返回给模型的条目。
 * state 跨表累计：{ matches, total, truncated, maxResults }
 */
function collectMatches(first, next, addressOf, toMatch, state) {
  if (!first) return { count: 0, complete: true }
  const firstAddress = addressOf(first)
  let cell = first
  let count = 0
  let complete = true
  do {
    count++
    if (state.matches.length < state.maxResults) state.matches.push(toMatch(cell))
    else state.truncated = true
    if (count >= MAX_FIND_SCAN_PER_SHEET) { complete = false; state.truncated = true; break }
    cell = next(cell)
  } while (cell && addressOf(cell) !== firstAddress)
  state.total += count
  return { count, complete }
}

function findHint(total, returned, truncated, inFormulas) {
  if (total === 0) {
    return inFormulas
      ? '公式里没有找到。要找显示值请用 scope=values'
      : '没有找到。筛选隐藏的行里的值搜不到；要搜公式文本请用 scope=formulas'
  }
  if (truncated) return `共找到 ${total} 处，只列出了前 ${returned} 处；缩小 sheets 范围或换更具体的 query`
  return null
}

function clipText(value, max) {
  const text = value == null ? '' : String(value)
  const limit = max || 120
  return text.length > limit ? text.slice(0, limit) + '…' : text
}

// ============ sheet_snapshot（对应 C# Perception/SnapshotEncoder.cs） ============

const SNAPSHOT_DEFAULT_CELLS = 60000
const SNAPSHOT_HARD_MAX_CELLS = 100000
const SNAPSHOT_MAX_COLUMNS = 100
const SNAPSHOT_MAX_FORMULAS = 20000
const SNAPSHOT_MAX_MERGES = 500
const SNAPSHOT_MAX_TEXT = 60
const SNAPSHOT_MAX_FORMULA = 300

// CVErr 的数值 = -2146826288 + (xlErr 常量 - 2000)
const ERROR_TEXT = {
  2000: '#NULL!', 2007: '#DIV/0!', 2015: '#VALUE!', 2023: '#REF!', 2029: '#NAME?', 2036: '#NUM!',
  2042: '#N/A', 2043: '#GETTING_DATA', 2045: '#SPILL!', 2046: '#CONNECT!', 2047: '#BLOCKED!',
  2048: '#UNKNOWN!', 2049: '#FIELD!', 2050: '#CALC!',
}
const ERROR_LITERAL = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CONNECT!|BLOCKED!|UNKNOWN!|FIELD!|CALC!)$/

function planSnapshotWindow(totalRows, totalColumns, maxCells) {
  if (totalRows <= 0 || totalColumns <= 0) return { rows: 0, columns: 0 }
  let budget = maxCells > 0 ? maxCells : SNAPSHOT_DEFAULT_CELLS
  budget = Math.min(budget, SNAPSHOT_HARD_MAX_CELLS)
  const columns = Math.min(totalColumns, SNAPSHOT_MAX_COLUMNS)
  const rows = Math.min(totalRows, Math.max(1, Math.floor(budget / columns)))
  return { rows, columns }
}

/** 数字格式是不是日期（去掉引号里的字面量和 [..] 段后看有没有 y / d，或不带 h、s 的 m） */
function isDateFormat(format) {
  if (typeof format !== 'string' || !format) return false
  const bare = format.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '')
  if (/general|@/i.test(bare) && !/[yd]/i.test(bare)) return false
  return /[yd]/i.test(bare) || (/m/i.test(bare) && !/[hs]/i.test(bare))
}

/** Excel 日期序列号 → 'yyyy-MM-dd'（有时间时 'yyyy-MM-dd HH:mm'） */
function serialToIso(serial) {
  const ms = Math.round((serial - 25569) * 86400000)  // 25569 = 1970-01-01 的序列号
  const d = new Date(ms)
  const pad = n => String(n).padStart(2, '0')
  const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`
  const minutes = d.getUTCHours() * 60 + d.getUTCMinutes()
  return minutes ? `${date} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}` : date
}

/** Value2 的一格 → null / 数字 / 布尔 / 截断文本 / {d} / {e}。asDate：这一列是日期格式 */
function encodeValue(value, asDate) {
  if (value === null || value === undefined) return null
  if (value instanceof Date) return { d: serialToIso(value.getTime() / 86400000 + 25569) }
  if (typeof value === 'boolean') return value
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return null
    const code = value + 2146828288
    if (Number.isInteger(value) && ERROR_TEXT[code]) return { e: ERROR_TEXT[code] }
    return asDate && value > 0 && value < 2958466 ? { d: serialToIso(value) } : value
  }
  const text = String(value)
  if (!text) return null
  if (ERROR_LITERAL.test(text)) return { e: text }
  return text.length > SNAPSHOT_MAX_TEXT ? text.slice(0, SNAPSHOT_MAX_TEXT) + '…' : text
}

function clipFormula(formula) {
  return formula.length > SNAPSHOT_MAX_FORMULA ? formula.slice(0, SNAPSHOT_MAX_FORMULA) + '…' : formula
}

module.exports = {
  DEFAULT_ROWS, MAX_ROWS, MAX_CELLS, MAX_COLUMNS, MAX_FIND_SCAN_PER_SHEET,
  clip, boxRows, boxColumns, plan, hint, qualifiedAddress, collectMatches, findHint, clipText,
  SNAPSHOT_MAX_FORMULAS, SNAPSHOT_MAX_MERGES,
  planSnapshotWindow, isDateFormat, serialToIso, encodeValue, clipFormula,
}
