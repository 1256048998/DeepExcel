// src/DeepExcel.Wps/tool-dispatcher.js
// 工具调度（对应 C# 端 ToolDispatcher.cs）
// 职责：
// 1. 接收 sidecar 的 tool_call 消息
// 2. 路由到 wps-actions.js（对应 IExcelActions）或 jsa-executor.js
// 3. 构建 Excel 快照返回给 sidecar（context 字段）
// 4. 工具执行结果包装为 { success, data, error, suggestion }

const WpsActions = require('./wps-actions')
const JsaExecutor = require('./jsa-executor')
const Ledger = require('./read-ledger')

class ToolDispatcher {
  constructor() {
    this.jsaExecutor = new JsaExecutor()
    // 先读后写：每个工作簿一本账（与 C# 端每个会话一本一致）
    this._ledgers = new Map()
    // 正在执行工具：这期间的 SheetChange 是我们自己的写入，不算用户改动
    this.isExecuting = false
  }

  /**
   * 执行工具（异步，对应 ToolDispatcher.Execute）：先读后写把关 → 执行 → 记账 → 附上用户改动提醒
   * @param {string} toolName
   * @param {object} args
   * @returns {Promise<{success: boolean, data: any, error: string, suggestion: string, warning?: string}>}
   */
  async execute(toolName, args) {
    const target = this._writeTarget(toolName, args || {})
    if (target) {
      const refusal = this._checkLedger(toolName, target)
      if (refusal) return this._withNotices(refusal)
    }
    let result
    this.isExecuting = true
    try {
      result = await this._executeCore(toolName, args)
    } finally {
      this.isExecuting = false
    }
    this._recordLedger(toolName, args || {}, result, target)
    return this._withNotices(result)
  }

  /** main.js 的 SheetChange 监听调用：不是我们的工具引起的改动，记为用户改动 */
  recordUserEdit(sheetName, address) {
    if (this.isExecuting || !address) return
    for (const part of String(address).split(',')) {
      const rect = Ledger.parseRect(part, sheetName)
      if (rect) this._ledger().recordUserEdit(rect)
    }
  }

  _ledger() {
    let key = 'workbook_unknown'
    try {
      const wb = wps.Application.ActiveWorkbook
      key = (wb && (wb.FullName || wb.Name)) || key
    } catch (e) { /* 读不到就用同一本账 */ }
    if (!this._ledgers.has(key)) this._ledgers.set(key, new Ledger.ReadLedger())
    return this._ledgers.get(key)
  }

  _activeSheetName() {
    try { return wps.Application.ActiveSheet.Name } catch (e) { return '' }
  }

  /**
   * 会改单元格内容、且目标区域明确的工具 → 目标区域（与 C# TryResolveWriteTarget 一致，
   * 只列 WPS 实现了的工具）。格式类、结构类、代码类不检查；地址解析不了也不检查。
   */
  _writeTarget(toolName, args) {
    const sheet = this._activeSheetName()
    const parse = key => Ledger.parseRect(this._getArg(args, key, ''), sheet)
    switch (toolName) {
      case 'write_value':
      case 'write_formula':
      case 'merge_cells':
        return parse('address')
      case 'clear_range':
        return String(this._getArg(args, 'clear_type', 'all')).toLowerCase() === 'formats' ? null : parse('address')
      case 'write_range': {
        const rect = parse('address')
        if (!rect) return null
        const values = this._getArg(args, 'values', null)
        const rows = Array.isArray(values) ? values.length : 1
        const cols = Array.isArray(values) && values.length
          ? Math.max(...values.map(r => (Array.isArray(r) ? r.length : 1)))
          : 1
        return Ledger.resize(rect, rows, cols)
      }
      case 'fill_formula_down': {
        const rect = parse('from_address')
        if (!rect) return null
        const count = Math.max(1, this._getInt(args, 'row_count') || 1)
        return Ledger.resize(rect, count + 1, rect.col2 - rect.col1 + 1)
      }
      case 'copy_range': {
        const source = parse('source_address')
        const dest = parse('dest_address')
        if (!source || !dest) return null
        return Ledger.resize(dest, source.row2 - source.row1 + 1, source.col2 - source.col1 + 1)
      }
      case 'sort_data':
        return parse('range_address')
      default:
        return null
    }
  }

  _checkLedger(toolName, target) {
    const address = Ledger.toA1(target)
    const { verdict, changed } = this._ledger().check(target, () => WpsActions.rangeHasContent(address))
    if (verdict === Ledger.Verdict.STALE_READ) {
      return this._refuse(toolName,
        `你上次读取之后，用户手动改动了 ${changed.join('、')}。你手里的是旧内容，本次未写入。`,
        `先重新 read_range 读取 ${address}，确认最新内容后再决定怎么写；不要把用户刚改的内容覆盖掉。`)
    }
    if (verdict === Ledger.Verdict.NOT_READ) {
      return this._refuse(toolName,
        `目标区域 ${address} 已有内容，但你还没有读过它，本次未写入。`,
        `先用 read_range 读取 ${address}（或包含它的区域），确认可以覆盖后再写。`)
    }
    return null
  }

  _recordLedger(toolName, args, result, target) {
    if (!result || !result.success) return
    try {
      if (toolName === 'read_range') {
        // 返回的地址不带表名：表名取参数里写的，没写就是活动表
        const requested = Ledger.parseRect(this._getArg(args, 'address', ''), this._activeSheetName())
        const sheet = requested ? requested.sheet : this._activeSheetName()
        const rect = Ledger.parseRect(result.data && result.data.address, sheet)
        if (rect) this._ledger().recordRead(rect)
      } else if (toolName === 'read_selection') {
        const data = result.data || {}
        const rect = Ledger.parseRect(data.address, data.worksheet || this._activeSheetName())
        if (rect) this._ledger().recordRead(rect)
      } else if (target) {
        this._ledger().recordOwnWrite(target)
      }
    } catch (e) {
      console.warn('[ToolDispatcher] ledger record failed:', e && e.message)
    }
  }

  /** 用户在模型工作期间手动改了单元格：附在这次工具结果里告诉模型（每处只说一次） */
  _withNotices(result) {
    if (!result) return result
    const ledger = this._ledger()
    const parts = ledger.takeNotices()
    const edits = ledger.takeUnreportedUserEdits()
    if (edits.length > 0) {
      parts.push(`用户刚刚手动修改了 ${edits.join('、')}（不是你的操作）。和这些单元格相关的数据请重新读取，不要用之前读到的旧值覆盖。`)
    }
    if (parts.length === 0) return result
    const notice = parts.join(' ')
    result.warning = result.warning ? result.warning + ' ' + notice : notice
    return result
  }

  _refuse(toolName, error, suggestion) {
    console.warn(`[ToolDispatcher] Write refused: tool=${toolName}, reason=${error}`)
    return { success: false, data: null, error, suggestion }
  }

  async _executeCore(toolName, args) {
    console.log(`[ToolDispatcher] Execute: ${toolName}, args keys=${Object.keys(args || {}).join(',')}`)
    try {
      switch (toolName) {
        case 'read_range': {
          // ★ 分页读取：先裁到已用区域、一次只读一页（与 C# 端 ReadRangePage 一致）
          const address = this._getArg(args, 'address', '')
          const page = WpsActions.readRangePage(address, this._getInt(args, 'offset') || 0, this._getInt(args, 'limit'))
          if (page && page.error) return { success: false, data: null, error: page.error, suggestion: page.suggestion || '' }
          return { success: true, data: page, suggestion: this._generateRangeSuggestion(page), error: '' }
        }

        case 'find': {
          const query = String(this._getArg(args, 'query', '') || '')
          if (!query) return { success: false, data: null, error: 'find 需要 query（要找的文本）', suggestion: '例如 find(query="应收账款")' }
          const inFormulas = String(this._getArg(args, 'scope', 'values')).toLowerCase() === 'formulas'
          const wholeCell = String(this._getArg(args, 'match', 'contains')).toLowerCase() === 'exact'
          const max = Math.min(200, Math.max(1, this._getInt(args, 'max_results') || 50))
          return this._wrapRead(WpsActions.findCells(query, inFormulas, this._getStringList(args, 'sheets'), wholeCell, max))
        }

        case 'list':
          return this._wrapRead(WpsActions.listObjects(this._getArg(args, 'kind', 'sheets')))

        // 侧车 inspect_sheet 的取数原语：不对模型开放，分析在侧车 perception 包里做
        case 'sheet_snapshot':
          return this._wrapRead(WpsActions.sheetSnapshot(this._getArg(args, 'sheet', ''), this._getInt(args, 'max_cells') || 0))

        case 'write_formula': {
          const addr = this._getArg(args, 'address', '')
          const formula = this._getArg(args, 'formula', '')
          WpsActions.writeFormula(addr, formula)
          return { success: true, data: { written: true, address: addr }, error: '' }
        }

        case 'write_value': {
          const addr = this._getArg(args, 'address', '')
          const value = this._getArg(args, 'value', '')
          WpsActions.writeValue(addr, value)
          return { success: true, data: { written: true, address: addr }, error: '' }
        }

        case 'write_range': {
          const addr = this._getArg(args, 'address', '')
          const values = this._getArg(args, 'values', [])
          WpsActions.writeRange(addr, values)
          return { success: true, data: { written: true, address: addr }, error: '' }
        }

        case 'read_workbook':
          return { success: true, data: WpsActions.readWorkbook(), error: '' }

        case 'read_selection':
          return { success: true, data: WpsActions.getSelection(), error: '' }

        case 'sort_data': {
          const rangeAddr = this._getArg(args, 'range_address', '')
          const sortColumn = this._getArg(args, 'sort_column', 'A')
          const descending = this._getArg(args, 'descending', false)
          const hasHeader = this._getArg(args, 'has_header', true)
          WpsActions.sortData(rangeAddr, sortColumn, descending, hasHeader)
          return { success: true, data: { sorted: true }, error: '' }
        }

        case 'filter_data': {
          const rangeAddr = this._getArg(args, 'range_address', '')
          const field = this._getInt(args, 'column_index') || 1
          const criteria = this._getArg(args, 'criteria', '')
          WpsActions.filterData(rangeAddr, field, criteria)
          return { success: true, data: { filtered: true }, error: '' }
        }

        case 'merge_cells': {
          WpsActions.mergeCells(this._getArg(args, 'address', ''))
          return { success: true, data: { merged: true }, error: '' }
        }

        case 'unmerge_cells': {
          WpsActions.unmergeCells(this._getArg(args, 'address', ''))
          return { success: true, data: { unmerged: true }, error: '' }
        }

        case 'add_sheet': {
          const name = this._getArg(args, 'name', '')
          const actualName = WpsActions.addSheet(name)
          return { success: true, data: { name: actualName }, error: '' }
        }

        case 'delete_sheet': {
          WpsActions.deleteSheet(this._getArg(args, 'name', ''))
          return { success: true, data: { deleted: true }, error: '' }
        }

        case 'rename_sheet': {
          WpsActions.renameSheet(this._getArg(args, 'old_name', ''), this._getArg(args, 'new_name', ''))
          return { success: true, data: { renamed: true }, error: '' }
        }

        case 'set_number_format': {
          WpsActions.setNumberFormat(this._getArg(args, 'address', ''), this._getArg(args, 'format', 'General'))
          return { success: true, data: { formatted: true }, error: '' }
        }

        case 'set_column_width': {
          WpsActions.setColumnWidth(this._getArg(args, 'address', ''), this._getArg(args, 'width', 10),
            !!this._getArg(args, 'auto_fit', false))
          return { success: true, data: { set: true }, error: '' }
        }

        case 'freeze_panes': {
          WpsActions.freezePanes(this._getArg(args, 'address', 'A1'))
          return { success: true, data: { frozen: true }, error: '' }
        }

        case 'fill_formula_down': {
          WpsActions.fillFormulaDown(this._getArg(args, 'from_address', ''), this._getInt(args, 'row_count') || 1)
          return { success: true, data: { filled: true }, error: '' }
        }

        case 'copy_range': {
          WpsActions.copyRange(this._getArg(args, 'source_address', ''), this._getArg(args, 'dest_address', ''))
          return { success: true, data: { copied: true }, error: '' }
        }

        case 'clear_range': {
          WpsActions.clearRange(this._getArg(args, 'address', ''), this._getArg(args, 'clear_type', 'all'))
          return { success: true, data: { cleared: true }, error: '' }
        }

        case 'insert_rows': {
          WpsActions.insertRows(this._getInt(args, 'row'), this._getInt(args, 'count') || 1)
          return { success: true, data: { inserted: true }, error: '' }
        }

        case 'delete_rows': {
          WpsActions.deleteRows(this._getInt(args, 'row'), this._getInt(args, 'count') || 1)
          return { success: true, data: { deleted: true }, error: '' }
        }

        case 'insert_columns': {
          WpsActions.insertColumns(this._getInt(args, 'column'), this._getInt(args, 'count') || 1)
          return { success: true, data: { inserted: true }, error: '' }
        }

        case 'delete_columns': {
          WpsActions.deleteColumns(this._getInt(args, 'column'), this._getInt(args, 'count') || 1)
          return { success: true, data: { deleted: true }, error: '' }
        }

        case 'set_cell_style': {
          WpsActions.setCellStyle(this._getArg(args, 'address', ''), {
            bold: this._getArg(args, 'bold', undefined),
            italic: this._getArg(args, 'italic', undefined),
            fontSize: this._getArg(args, 'font_size', undefined),
            fontColor: this._getArg(args, 'font_color', undefined),
            bgColor: this._getArg(args, 'bg_color', undefined),
            fontName: this._getArg(args, 'font_name', undefined),
            horizontalAlignment: this._getArg(args, 'h_align', undefined),
            verticalAlignment: this._getArg(args, 'v_align', undefined),
            wrapText: this._getArg(args, 'wrap_text', undefined),
          })
          return { success: true, data: { styled: true }, error: '' }
        }

        case 'write_table': {
          const table = WpsActions.createTable(this._getArg(args, 'address', ''), this._getArg(args, 'table_name', ''))
          return { success: true, data: table, error: '' }
        }

        // ★ JSA 宏执行（WPS 替代 VBA）
        case 'execute_jsa': {
          const code = this._getArg(args, 'code', '')
          const result = await this.jsaExecutor.execute(code)
          return { success: result.success, data: result.data, error: result.error, suggestion: '' }
        }

        default:
          return {
            success: false,
            data: {},
            error: `未知工具: ${toolName}`,
            suggestion: `WPS 端暂未实现工具 ${toolName}，请使用其他工具或 execute_jsa`,
          }
      }
    } catch (err) {
      console.error(`[ToolDispatcher] Execute FAILED: ${toolName}`, err)
      return {
        success: false,
        data: {},
        error: err.message,
        suggestion: '',
      }
    }
  }

  /**
   * 构建 Excel 快照（对应 ToolDispatcher.BuildExcelSnapshot）
   * 返回给 sidecar 作为 context 字段
   */
  buildExcelSnapshot() {
    try {
      const workbook = WpsActions.readWorkbook()
      const selection = WpsActions.getSelection()
      return {
        workbook,
        selection,
        timestamp: new Date().toISOString(),
        // ★ 标识当前宿主是 WPS（sidecar 据此调整 AI 提示）
        host_type: 'wps',
      }
    } catch (err) {
      console.error('[ToolDispatcher] buildExcelSnapshot failed:', err)
      return { error: err.message, host_type: 'wps' }
    }
  }

  // ============= 工具方法 =============

  _getArg(args, key, defaultValue) {
    if (!args || args[key] === undefined) return defaultValue
    return args[key]
  }

  /** 可选整数参数：没传、空串、不是数字都返回 null */
  _getInt(args, key) {
    const value = this._getArg(args, key, null)
    if (value === null || value === '') return null
    const n = parseInt(value, 10)
    return Number.isNaN(n) ? null : n
  }

  /** 字符串列表参数：数组，或逗号（含全角逗号）分隔的文本 */
  _getStringList(args, key) {
    const value = this._getArg(args, key, null)
    const list = Array.isArray(value) ? value : (typeof value === 'string' ? value.split(/[,，]/) : [])
    return list.map(v => String(v).trim()).filter(Boolean)
  }

  /** 读取类结果：宿主返回 { error } 时算失败（与 C# WrapReadResult 一致） */
  _wrapRead(data) {
    if (data && data.error) return { success: false, data: null, error: data.error, suggestion: data.suggestion || '' }
    return { success: true, data, error: '' }
  }

  /**
   * 生成范围建议（用于 AI 后续推理）
   */
  _generateRangeSuggestion(page) {
    if (!page) return ''
    if (page.hint) return page.hint
    return `数据范围 ${page.rowCount || 0} 行 × ${page.columnCount || 0} 列`
  }
}

module.exports = ToolDispatcher
