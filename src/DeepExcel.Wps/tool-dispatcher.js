// src/DeepExcel.Wps/tool-dispatcher.js
// 工具调度（对应 C# 端 ToolDispatcher.cs）
// 职责：
// 1. 接收 sidecar 的 tool_call 消息
// 2. 路由到 wps-actions.js（对应 IExcelActions）或 jsa-executor.js
// 3. 构建 Excel 快照返回给 sidecar（context 字段）
// 4. 工具执行结果包装为 { success, data, error, suggestion }

const WpsActions = require('./wps-actions')
const JsaExecutor = require('./jsa-executor')

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
