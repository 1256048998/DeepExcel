// src/DeepExcel.Wps/range-paging.js
// read_range 分页与 find 汇总的纯逻辑（对应 C# 端 Perception/RangePaging.cs 与
// ExcelActionsImpl.FindCells 的计数规则）。不碰 wps 对象，便于 node 直接测试：
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

module.exports = {
  DEFAULT_ROWS, MAX_ROWS, MAX_CELLS, MAX_COLUMNS, MAX_FIND_SCAN_PER_SHEET,
  clip, boxRows, boxColumns, plan, hint, qualifiedAddress, collectMatches, findHint, clipText,
}
