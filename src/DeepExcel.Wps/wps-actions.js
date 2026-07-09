// src/DeepExcel.Wps/wps-actions.js
// WPS JS API 实现（对应 C# 端 ExcelActionsImpl）
// 对齐 IExcelActions 接口，用 wps.Application JS API 操作 WPS 表格
//
// ★ WPS JS API 全局对象说明：
// - wps.Application           表格应用入口（对应 Excel.Application）
// - wps.Application.Workbooks 工作簿集合
// - wps.Application.ActiveWorkbook  活动工作簿
// - wps.Application.ActiveSheet     活动工作表
// - wps.Application.Range(addr)     获取范围
// - wps.Application.Selection       当前选中
// - wps.Application.ActiveCell      活动单元格
//
// ★ WPS JS API 与 Excel COM 的主要差异：
// 1. JS API 是同步的（不需要 STA 线程调度）
// 2. Value2 返回的是 JS 数组而非 2D 数组对象（WPS 自动转换）
// 3. 枚举值用数字常量（xlSortColumns=1 等）

const WpsActions = {
  // ============= 读取类 =============

  /**
   * 读取指定范围的单元格数据
   * @param {string} address A1 格式地址，如 "A1:B10" 或 "Sheet1!A1:B10"
   */
  readRange(address) {
    const range = this._getRange(address)
    const values = range.Value2
    const formulas = range.Formula
    return {
      address: range.Address,
      values: values,
      formulas: formulas,
      rowCount: range.Rows.Count,
      columnCount: range.Columns.Count,
      numberFormats: this._getNumberFormats(range),
    }
  },

  /**
   * 读取整个工作簿结构（所有 sheet + UsedRange 信息）
   */
  readWorkbook() {
    const app = wps.Application
    const wb = app.ActiveWorkbook
    if (!wb) return { worksheets: [], activeSheet: '' }

    const sheets = []
    for (let i = 1; i <= wb.Worksheets.Count; i++) {
      const ws = wb.Worksheets(i)
      let used = { address: '', rowCount: 0, columnCount: 0 }
      try {
        const usedRange = ws.UsedRange
        used = {
          address: usedRange.Address,
          rowCount: usedRange.Rows.Count,
          columnCount: usedRange.Columns.Count,
        }
      } catch (e) { /* 空表 */ }
      sheets.push({
        name: ws.Name,
        index: i,
        visible: ws.Visible,
        usedRange: used,
      })
    }

    return {
      name: wb.Name,
      path: wb.FullName,
      worksheets: sheets,
      activeSheet: (wb.ActiveSheet && wb.ActiveSheet.Name) || '',
    }
  },

  /**
   * 获取当前选中区域信息
   */
  getSelection() {
    const app = wps.Application
    const sel = app.Selection
    if (!sel) return { address: '', worksheet: '', rowCount: 0, columnCount: 0 }
    try {
      const wsName = sel.Worksheet ? sel.Worksheet.Name : ''
      return {
        address: sel.Address,
        worksheet: wsName,
        rowCount: sel.Rows.Count,
        columnCount: sel.Columns.Count,
      }
    } catch (e) {
      return { address: '', worksheet: '', rowCount: 0, columnCount: 0, error: e.message }
    }
  },

  // ============= 写入类 =============

  writeFormula(address, formula) {
    const range = this._getRange(address)
    range.Formula = formula
  },

  writeValue(address, value) {
    const range = this._getRange(address)
    // ★ WPS JS API 直接赋值即可，类型自动转换
    range.Value2 = value
  },

  /**
   * 批量写入范围（values 是 2D 数组）
   */
  writeRange(address, values) {
    const range = this._getRange(address)
    range.Value2 = values
  },

  /**
   * 写入表格数据（带表头）
   */
  writeTable(address, headers, rows) {
    const range = this._getRange(address)
    const data = [headers, ...rows]
    range.Value2 = data
    // 表头加粗
    const headerRange = range.Resize(1, headers.length)
    headerRange.Font.Bold = true
  },

  // ============= Sheet 操作 =============

  addSheet(name) {
    const app = wps.Application
    const wb = app.ActiveWorkbook
    const ws = wb.Worksheets.Add()
    if (name) ws.Name = name
    return ws.Name
  },

  deleteSheet(name) {
    const app = wps.Application
    const wb = app.ActiveWorkbook
    // ★ WPS 删除 sheet 需 DisplayAlerts=false 避免弹确认对话框
    const oldAlerts = app.DisplayAlerts
    app.DisplayAlerts = false
    try {
      wb.Worksheets(name).Delete()
    } finally {
      app.DisplayAlerts = oldAlerts
    }
  },

  renameSheet(oldName, newName) {
    const app = wps.Application
    app.ActiveWorkbook.Worksheets(oldName).Name = newName
  },

  // ============= 格式类 =============

  setNumberFormat(address, format) {
    this._getRange(address).NumberFormat = format
  },

  setColumnWidth(address, width) {
    this._getRange(address).ColumnWidth = width
  },

  setCellStyle(address, style) {
    const range = this._getRange(address)
    if (style.bold !== undefined) range.Font.Bold = style.bold
    if (style.italic !== undefined) range.Font.Italic = style.italic
    if (style.fontSize !== undefined) range.Font.Size = style.fontSize
    if (style.fontColor !== undefined) range.Font.Color = this._parseColor(style.fontColor)
    if (style.bgColor !== undefined) range.Interior.Color = this._parseColor(style.bgColor)
    if (style.horizontalAlignment !== undefined) {
      range.HorizontalAlignment = this._parseHAlign(style.horizontalAlignment)
    }
  },

  // ============= 数据操作类 =============

  /**
   * 排序数据
   * @param {string} rangeAddress 排序范围
   * @param {string|number} sortColumn 排序列（字母如 "A" 或索引）
   * @param {boolean} descending 是否降序
   * @param {boolean} hasHeader 是否有表头
   */
  sortData(rangeAddress, sortColumn, descending, hasHeader) {
    const range = this._getRange(rangeAddress)
    // ★ 排序 Key 指向数据行（hasHeader=true 时用第2行），与 C# 端逻辑一致
    const colIdx = typeof sortColumn === 'string'
      ? this._columnLetterToIndex(sortColumn)
      : sortColumn
    const keyRange = hasHeader
      ? range.Cells(2, colIdx)
      : range.Cells(1, colIdx)
    // ★ Header=xlNo(0) 让所有行参与排序（xlYes=1 会跳过首行）
    // ★ SortMethod=xlPinYin(1) 拼音排序
    range.Sort(keyRange, descending ? 2 : 1, null, null, null, null, null, hasHeader ? 0 : 0, 1, 1, 1)
  },

  /**
   * 自动筛选
   */
  filterData(rangeAddress, field, criteria1) {
    const range = this._getRange(rangeAddress)
    range.AutoFilter(field, criteria1)
  },

  mergeCells(address) {
    this._getRange(address).Merge()
  },

  unmergeCells(address) {
    this._getRange(address).UnMerge()
  },

  copyRange(sourceAddress, destAddress) {
    const src = this._getRange(sourceAddress)
    const dest = this._getRange(destAddress)
    src.Copy(dest)
  },

  clearRange(address) {
    this._getRange(address).Clear()
  },

  // ============= 行列操作 =============

  insertRows(address, count) {
    const range = this._getRange(address)
    for (let i = 0; i < count; i++) {
      range.Insert(-4121) // xlDown
    }
  },

  deleteRows(address, count) {
    const range = this._getRange(address)
    for (let i = 0; i < count; i++) {
      range.EntireRow.Delete()
    }
  },

  insertColumns(address, count) {
    const range = this._getRange(address)
    for (let i = 0; i < count; i++) {
      range.Insert(-4159) // xlToRight
    }
  },

  deleteColumns(address, count) {
    const range = this._getRange(address)
    for (let i = 0; i < count; i++) {
      range.EntireColumn.Delete()
    }
  },

  freezePanes(address) {
    const app = wps.Application
    const range = this._getRange(address)
    range.Select()
    app.ActiveWindow.FreezePanes = true
  },

  // ============= 公式填充 =============

  fillFormulaDown(fromAddress, toAddress) {
    const app = wps.Application
    const fromRange = app.Range(fromAddress)
    const toRange = app.Range(toAddress)
    fromRange.AutoFill(toRange, 0) // xlFillDefault=0
  },

  // ============= 工具方法 =============

  /**
   * 获取 Range 对象
   * 支持 "A1:B10" 和 "Sheet1!A1:B10" 两种格式
   */
  _getRange(address) {
    const app = wps.Application
    if (address.includes('!')) {
      const [sheetName, rangeAddr] = address.split('!')
      const ws = app.ActiveWorkbook.Worksheets(sheetName)
      return ws.Range(rangeAddr)
    }
    return app.Range(address)
  },

  /**
   * 获取范围的所有 NumberFormat（2D 数组）
   */
  _getNumberFormats(range) {
    try {
      const rowCount = range.Rows.Count
      const colCount = range.Columns.Count
      const formats = []
      for (let r = 1; r <= rowCount; r++) {
        const row = []
        for (let c = 1; c <= colCount; c++) {
          row.push(range.Cells(r, c).NumberFormat)
        }
        formats.push(row)
      }
      return formats
    } catch (e) {
      return null
    }
  },

  /**
   * 列字母转索引：A→1, B→2, ..., AA→27
   */
  _columnLetterToIndex(letter) {
    let result = 0
    for (let i = 0; i < letter.length; i++) {
      result = result * 26 + (letter.charCodeAt(i) - 64)
    }
    return result
  },

  /**
   * 解析颜色字符串为 RGB 数字（WPS 使用 BGR 格式：0xBBGGRR）
   */
  _parseColor(color) {
    if (typeof color === 'number') return color
    if (typeof color === 'string' && color.startsWith('#')) {
      const hex = color.slice(1)
      const r = parseInt(hex.slice(0, 2), 16)
      const g = parseInt(hex.slice(2, 4), 16)
      const b = parseInt(hex.slice(4, 6), 16)
      return r + g * 256 + b * 65536
    }
    return 0
  },

  /**
   * 水平对齐字符串转枚举值
   */
  _parseHAlign(align) {
    const map = {
      'left': -4131,    // xlLeft
      'center': -4108,  // xlCenter
      'right': -4152,   // xlRight
      'general': 1,     // xlGeneral
    }
    return map[align] || 1
  },
}

module.exports = WpsActions
