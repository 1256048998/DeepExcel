// src/DeepExcel.Wps/read-ledger.js
// 先读后写 + 读后被改检测（对应 C# Sidecar/ReadLedger.cs，规则逐条一致）
//
// 模型看不到工作簿本身，只看得到它读过的部分。没读过就覆盖一块有内容的区域、
// 用户在它工作时手动改了单元格它又照旧值写回去——这里记下模型读过 / 写过的区域和
// 用户手动改过的区域，写入前据此把关。纯逻辑，不碰 WPS 对象模型。

'use strict'

const MAX_ROW = 1048576
const MAX_COLUMN = 16384
const MAX_ENTRIES = 300

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d+)$/
const COLUMN = /^\$?([A-Za-z]{1,3})$/
const ROW = /^\$?(\d+)$/

function columnIndex(letters) {
  let index = 0
  for (const ch of letters.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64)
  return index
}

function columnName(column) {
  let name = ''
  let n = column
  while (n > 0) {
    const rem = (n - 1) % 26
    name = String.fromCharCode(65 + rem) + name
    n = Math.floor((n - 1) / 26)
  }
  return name
}

function makeRect(sheet, row1, col1, row2, col2) {
  return {
    sheet: sheet || '',
    row1: Math.min(row1, row2), col1: Math.min(col1, col2),
    row2: Math.max(row1, row2), col2: Math.max(col1, col2),
  }
}

/**
 * Sheet1!A1:C4、'My Sheet'!B2、$A$1、A:C（整列）、2:5（整行）。没写表名用 defaultSheet。
 * 命名区域、R1C1、多区域（逗号）不解析，返回 null。
 */
function parseRect(address, defaultSheet) {
  if (!address || !String(address).trim()) return null
  let text = String(address).trim()
  let sheet = defaultSheet || ''
  const bang = text.lastIndexOf('!')
  if (bang >= 0) {
    sheet = text.slice(0, bang).trim()
    if (sheet.length >= 2 && sheet[0] === "'" && sheet[sheet.length - 1] === "'") {
      sheet = sheet.slice(1, -1).replace(/''/g, "'")
    }
    text = text.slice(bang + 1).trim()
  }
  if (!text || text.indexOf(',') >= 0) return null
  const parts = text.split(':')
  if (parts.length > 2) return null
  const a = parts[0]
  const b = parts.length === 2 ? parts[1] : parts[0]

  const ca = CELL.exec(a)
  const cb = CELL.exec(b)
  if (ca && cb) {
    const r1 = parseInt(ca[2], 10), r2 = parseInt(cb[2], 10)
    const c1 = columnIndex(ca[1]), c2 = columnIndex(cb[1])
    if ([r1, r2].some(r => r < 1 || r > MAX_ROW) || [c1, c2].some(c => c < 1 || c > MAX_COLUMN)) return null
    return makeRect(sheet, r1, c1, r2, c2)
  }
  const la = COLUMN.exec(a), lb = COLUMN.exec(b)
  if (parts.length === 2 && la && lb) return makeRect(sheet, 1, columnIndex(la[1]), MAX_ROW, columnIndex(lb[1]))
  const ra = ROW.exec(a), rb = ROW.exec(b)
  if (parts.length === 2 && ra && rb) return makeRect(sheet, parseInt(ra[1], 10), 1, parseInt(rb[1], 10), MAX_COLUMN)
  return null
}

function overlaps(x, y) {
  return String(x.sheet).toLowerCase() === String(y.sheet).toLowerCase() &&
    x.row1 <= y.row2 && y.row1 <= x.row2 && x.col1 <= y.col2 && y.col1 <= x.col2
}

function resize(rect, rows, columns) {
  return makeRect(rect.sheet, rect.row1, rect.col1,
    Math.min(MAX_ROW, rect.row1 + Math.max(1, rows) - 1),
    Math.min(MAX_COLUMN, rect.col1 + Math.max(1, columns) - 1))
}

/** 与 C# CellRect.ToA1 同一写法：表名含空格、-、!、' 时加引号 */
function toA1(rect) {
  const sheet = /[ \-!']/.test(rect.sheet) ? "'" + String(rect.sheet).replace(/'/g, "''") + "'" : rect.sheet
  const first = columnName(rect.col1) + rect.row1
  const last = columnName(rect.col2) + rect.row2
  return sheet + '!' + (first === last ? first : first + ':' + last)
}

const Verdict = { ALLOWED: 'allowed', STALE_READ: 'stale_read', NOT_READ: 'not_read' }

class ReadLedger {
  constructor() {
    this._known = []
    this._userEdits = []
    this._notices = []
    this._seq = 0
  }

  /** 模型读到了这块区域的内容 */
  recordRead(rect) { this._add(this._known, rect) }

  /** 模型自己写了这块区域：它知道这里现在是什么，等同于读过 */
  recordOwnWrite(rect) { this._add(this._known, rect) }

  /** 用户（或别的程序）在 WPS 里改了这块区域，不是我们的工具改的 */
  recordUserEdit(rect) { this._add(this._userEdits, rect) }

  /** 回滚之类整体换掉内容的操作之后：之前读到的都作废 */
  forgetReads() { this._known = [] }

  /**
   * @param {object} target 目标区域
   * @param {() => boolean} targetHasContent 从没读过时才问：目标区域有没有内容
   * @returns {{verdict: string, changed: string[]}}
   */
  check(target, targetHasContent) {
    const known = this._known.filter(e => overlaps(e.rect, target))
    if (known.length > 0) {
      const lastKnown = Math.max(...known.map(e => e.seq))
      const changed = [...new Set(this._userEdits
        .filter(e => e.seq > lastKnown && overlaps(e.rect, target))
        .map(e => toA1(e.rect)))].slice(0, 10)
      return { verdict: changed.length > 0 ? Verdict.STALE_READ : Verdict.ALLOWED, changed }
    }
    let hasContent
    try { hasContent = !targetHasContent || !!targetHasContent() } catch (e) { hasContent = false }
    return { verdict: hasContent ? Verdict.NOT_READ : Verdict.ALLOWED, changed: [] }
  }

  /** 还没告诉模型的用户改动（每条只报一次） */
  takeUnreportedUserEdits() {
    const pending = this._userEdits.filter(e => !e.reported)
    pending.forEach(e => { e.reported = true })
    return [...new Set(pending.map(e => toA1(e.rect)))].slice(0, 20)
  }

  addNotice(notice) {
    if (!notice || !String(notice).trim()) return
    this._notices.push(String(notice))
    if (this._notices.length > 5) this._notices.shift()
  }

  takeNotices() {
    const pending = this._notices.slice()
    this._notices = []
    return pending
  }

  _add(list, rect) {
    list.push({ rect, seq: ++this._seq, reported: false })
    if (list.length > MAX_ENTRIES) list.splice(0, list.length - MAX_ENTRIES)
  }
}

module.exports = { ReadLedger, Verdict, parseRect, overlaps, resize, toA1, columnName, MAX_ROW, MAX_COLUMN }
