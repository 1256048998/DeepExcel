// src/DeepExcel.Wps/tool-dispatcher.js
// 工具调度（对应 C# 端 ToolDispatcher.cs）
// 职责：
// 1. 接收 sidecar 的 tool_call 消息
// 2. 路由到 wps-actions.js（对应 IExcelActions）或 jsa-executor.js
// 3. 构建 Excel 快照返回给 sidecar（context 字段）
// 4. 工具执行结果包装为 { success, data, error, suggestion }

const WpsActions = require('./wps-actions')
const JsaExecutor = require('./jsa-executor')

// ★ read_range 返回的最大行数限制（与 C# 端 MaxReadRangeRows=200 一致）
const MAX_READ_RANGE_ROWS = 200

class ToolDispatcher {
  constructor() {
    this.jsaExecutor = new JsaExecutor()
  }

  /**
   * 执行工具（异步，对应 ToolDispatcher.Execute）
   * @param {string} toolName
   * @param {object} args
   * @returns {Promise<{success: boolean, data: any, error: string, suggestion: string}>}
   */
  async execute(toolName, args) {
    console.log(`[ToolDispatcher] Execute: ${toolName}, args keys=${Object.keys(args || {}).join(',')}`)
    try {
      switch (toolName) {
        case 'echo':
          return { success: true, data: { echo: this._getArg(args, 'text', '') } }

        case 'read_range': {
          const address = this._getArg(args, 'address', '')
          const rangeData = WpsActions.readRange(address)
          // ★ 截断超过 MAX_READ_RANGE_ROWS 的数据（与 C# 端一致）
          const truncated = this._truncateRangeData(rangeData, MAX_READ_RANGE_ROWS)
          return { success: true, data: truncated, suggestion: this._generateRangeSuggestion(rangeData), error: '' }
        }

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
          const rangeAddr = this._getArg(args, 'address', '')
          const sortColumn = this._getArg(args, 'sort_column', 'A')
          const descending = this._getArg(args, 'descending', false)
          const hasHeader = this._getArg(args, 'has_header', true)
          WpsActions.sortData(rangeAddr, sortColumn, descending, hasHeader)
          return { success: true, data: { sorted: true }, error: '' }
        }

        case 'filter_data': {
          const rangeAddr = this._getArg(args, 'address', '')
          const field = this._getArg(args, 'field', 1)
          const criteria1 = this._getArg(args, 'criteria1', '')
          WpsActions.filterData(rangeAddr, field, criteria1)
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
          WpsActions.setColumnWidth(this._getArg(args, 'address', ''), this._getArg(args, 'width', 10))
          return { success: true, data: { set: true }, error: '' }
        }

        case 'freeze_panes': {
          WpsActions.freezePanes(this._getArg(args, 'address', 'A1'))
          return { success: true, data: { frozen: true }, error: '' }
        }

        case 'fill_formula_down': {
          WpsActions.fillFormulaDown(this._getArg(args, 'from', ''), this._getArg(args, 'to', ''))
          return { success: true, data: { filled: true }, error: '' }
        }

        case 'copy_range': {
          WpsActions.copyRange(this._getArg(args, 'source', ''), this._getArg(args, 'destination', ''))
          return { success: true, data: { copied: true }, error: '' }
        }

        case 'clear_range': {
          WpsActions.clearRange(this._getArg(args, 'address', ''))
          return { success: true, data: { cleared: true }, error: '' }
        }

        case 'insert_rows': {
          WpsActions.insertRows(this._getArg(args, 'address', ''), this._getArg(args, 'count', 1))
          return { success: true, data: { inserted: true }, error: '' }
        }

        case 'delete_rows': {
          WpsActions.deleteRows(this._getArg(args, 'address', ''), this._getArg(args, 'count', 1))
          return { success: true, data: { deleted: true }, error: '' }
        }

        case 'insert_columns': {
          WpsActions.insertColumns(this._getArg(args, 'address', ''), this._getArg(args, 'count', 1))
          return { success: true, data: { inserted: true }, error: '' }
        }

        case 'delete_columns': {
          WpsActions.deleteColumns(this._getArg(args, 'address', ''), this._getArg(args, 'count', 1))
          return { success: true, data: { deleted: true }, error: '' }
        }

        case 'set_cell_style': {
          WpsActions.setCellStyle(this._getArg(args, 'address', ''), {
            bold: this._getArg(args, 'bold', undefined),
            italic: this._getArg(args, 'italic', undefined),
            fontSize: this._getArg(args, 'font_size', undefined),
            fontColor: this._getArg(args, 'font_color', undefined),
            bgColor: this._getArg(args, 'bg_color', undefined),
            horizontalAlignment: this._getArg(args, 'horizontal_alignment', undefined),
          })
          return { success: true, data: { styled: true }, error: '' }
        }

        case 'write_table': {
          WpsActions.writeTable(
            this._getArg(args, 'address', ''),
            this._getArg(args, 'headers', []),
            this._getArg(args, 'rows', [])
          )
          return { success: true, data: { written: true }, error: '' }
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

  /**
   * 截断超过 maxRows 的范围数据（与 C# 端逻辑一致）
   */
  _truncateRangeData(rangeData, maxRows) {
    if (!rangeData || !rangeData.values) return rangeData
    const rowCount = rangeData.rowCount || 0
    if (rowCount <= maxRows) return rangeData

    const truncatedValues = rangeData.values.slice(0, maxRows)
    const truncatedFormulas = rangeData.formulas ? rangeData.formulas.slice(0, maxRows) : null
    const truncatedFormats = rangeData.numberFormats ? rangeData.numberFormats.slice(0, maxRows) : null

    return {
      ...rangeData,
      values: truncatedValues,
      formulas: truncatedFormulas,
      numberFormats: truncatedFormats,
      rowCount: maxRows,
      truncated: true,
      original_row_count: rowCount,
      truncation_hint: `数据超过 ${maxRows} 行已截断，请用 read_range 指定更小范围查看剩余数据`,
    }
  }

  /**
   * 生成范围建议（用于 AI 后续推理）
   */
  _generateRangeSuggestion(rangeData) {
    if (!rangeData) return ''
    const rowCount = rangeData.rowCount || 0
    const colCount = rangeData.columnCount || 0
    if (rowCount > 100) {
      return `数据范围 ${rowCount} 行 × ${colCount} 列，已截断显示前 ${MAX_READ_RANGE_ROWS} 行。如需查看全部数据请分段读取。`
    }
    return `数据范围 ${rowCount} 行 × ${colCount} 列`
  }
}

module.exports = ToolDispatcher
