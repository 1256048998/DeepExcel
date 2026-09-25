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

const Paging = require('./range-paging')

// Find 的枚举值（WPS JS API 用数字常量）
const XL_VALUES = -4163
const XL_FORMULAS = -4123
const XL_WHOLE = 1
const XL_PART = 2
const XL_BY_ROWS = 1
const XL_NEXT = 1
// list 每类最多列这么多个，和 C# 端 MaxListEntries 一致
const MAX_LIST_ENTRIES = 200

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
   * 分页读取（对应 C# ReadRangePage）：先裁到已用区域，再只读一页。
   * 以前 read_range("A:A") 会逐格读 104 万行的数字格式，WPS 直接卡死。
   */
  readRangePage(address, offset, limit) {
    let range = this._getRange(address)
    try { if (range.Areas.Count > 1) range = range.Areas(1) } catch (e) { /* 单块区域 */ }
    const ws = range.Worksheet
    const requestedBox = {
      row1: range.Row, col1: range.Column,
      row2: range.Row + range.Rows.Count - 1, col2: range.Column + range.Columns.Count - 1,
    }
    const used = ws.UsedRange
    const usedBox = {
      row1: used.Row, col1: used.Column,
      row2: used.Row + used.Rows.Count - 1, col2: used.Column + used.Columns.Count - 1,
    }
    const requested = Paging.qualifiedAddress(ws.Name, this._address(range, true))

    const box = Paging.clip(requestedBox, usedBox)
    if (!box) {
      return {
        address: this._address(range, false), values: [], formulas: [], rowCount: 0, columnCount: 0,
        paging: { requested, total_rows: 0, total_columns: 0, offset: 0, returned_rows: 0, next_offset: null, columns_truncated: false, clipped_to_used_range: true },
        hint: '这个区域在工作表的已用范围之外，全是空白。',
      }
    }

    const totalRows = Paging.boxRows(box)
    const totalColumns = Paging.boxColumns(box)
    const plan = Paging.plan(totalRows, totalColumns, offset, limit)
    if (plan.error) return { error: plan.error, suggestion: '从 offset=0 开始读，或换一个地址' }

    const first = ws.Cells(box.row1 + plan.offset, box.col1)
    const last = ws.Cells(box.row1 + plan.offset + plan.rows - 1, box.col1 + plan.columns - 1)
    const page = ws.Range(first, last)
    return {
      address: this._address(page, false),
      values: this._as2d(page.Value2),
      formulas: this._as2d(page.Formula),
      rowCount: plan.rows,
      columnCount: plan.columns,
      numberFormats: this._getNumberFormats(page),
      paging: {
        requested,
        total_rows: totalRows,
        total_columns: totalColumns,
        offset: plan.offset,
        returned_rows: plan.rows,
        next_offset: plan.nextOffset,
        columns_truncated: plan.columnsTruncated,
        clipped_to_used_range: box.row1 !== requestedBox.row1 || box.col1 !== requestedBox.col1 ||
          box.row2 !== requestedBox.row2 || box.col2 !== requestedBox.col2,
      },
      hint: Paging.hint(requested, totalRows, totalColumns, plan),
    }
  },

  /**
   * 全工作簿查找（对应 C# FindCells）：UsedRange.Find + FindNext 转一圈。
   */
  findCells(query, inFormulas, sheets, wholeCell, maxResults) {
    const wb = wps.Application.ActiveWorkbook
    if (!wb) return { error: '没有打开的工作簿' }
    const wanted = (sheets || []).map(s => String(s).toLowerCase())
    const state = { matches: [], total: 0, truncated: false, maxResults }
    const perSheet = []
    let searched = 0
    for (let i = 1; i <= wb.Worksheets.Count; i++) {
      const ws = wb.Worksheets(i)
      if (wanted.length && wanted.indexOf(String(ws.Name).toLowerCase()) < 0) continue
      searched++
      let used
      try { used = ws.UsedRange } catch (e) { continue }
      let first = null
      try {
        // Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase)
        first = used.Find(query, undefined,
          inFormulas ? XL_FORMULAS : XL_VALUES, wholeCell ? XL_WHOLE : XL_PART,
          XL_BY_ROWS, XL_NEXT, false)
      } catch (e) { first = null }
      if (!first) continue
      const counted = Paging.collectMatches(
        first,
        cell => { try { return used.FindNext(cell) } catch (e) { return null } },
        cell => this._address(cell, false),
        cell => ({
          sheet: ws.Name,
          address: this._address(cell, true),
          value: Paging.clipText(cell.Text),
          formula: cell.HasFormula ? Paging.clipText(cell.Formula, 200) : null,
        }),
        state,
      )
      perSheet.push({ sheet: ws.Name, count: counted.count, complete: counted.complete })
    }
    if (wanted.length && searched === 0) {
      return { error: '找不到指定的工作表：' + sheets.join('、'), suggestion: '先用 list(kind=sheets) 看有哪些表' }
    }
    return {
      query,
      scope: inFormulas ? 'formulas' : 'values',
      total: state.total,
      returned: state.matches.length,
      truncated: state.truncated,
      matches: state.matches,
      per_sheet: perSheet,
      hint: Paging.findHint(state.total, state.matches.length, state.truncated, inFormulas),
    }
  },

  /**
   * 列出工作簿对象（对应 C# ListObjects）：sheets / names / tables / pivots / charts
   */
  listObjects(kind) {
    const wb = wps.Application.ActiveWorkbook
    if (!wb) return { error: '没有打开的工作簿' }
    const k = String(kind || 'sheets').trim().toLowerCase()
    const items = []
    const full = () => items.length >= MAX_LIST_ENTRIES
    const each = (collection, fn) => {
      let count = 0
      try { count = collection.Count } catch (e) { return }
      for (let i = 1; i <= count && !full(); i++) {
        try { fn(collection.Item(i)) } catch (e) { /* 单个对象读不到就跳过 */ }
      }
    }
    const sheetsOf = () => {
      const list = []
      for (let i = 1; i <= wb.Worksheets.Count; i++) list.push(wb.Worksheets(i))
      return list
    }

    switch (k) {
      case 'sheets':
        for (const ws of sheetsOf()) {
          let usedRange = null; let rows = 0; let columns = 0
          try { const u = ws.UsedRange; usedRange = this._address(u, true); rows = u.Rows.Count; columns = u.Columns.Count } catch (e) { /* 空表 */ }
          let isProtected = false
          try { isProtected = !!ws.ProtectContents } catch (e) { /* 读不到按未保护 */ }
          items.push({ name: ws.Name, index: ws.Index, visibility: this._visibility(ws.Visible), used_range: usedRange, rows, columns, protected: isProtected })
        }
        break
      case 'names':
        each(wb.Names, n => {
          let refersTo = null
          try { refersTo = String(n.RefersTo) } catch (e) { /* 读不到 */ }
          let scope = 'workbook'
          try { if (n.Parent && n.Parent.Name !== wb.Name) scope = n.Parent.Name } catch (e) { /* 工作簿级 */ }
          let visible = true
          try { visible = n.Visible !== false } catch (e) { /* 默认可见 */ }
          items.push({ name: n.Name, refers_to: refersTo == null ? null : Paging.clipText(refersTo, 200), scope, visible, broken: !!refersTo && refersTo.indexOf('#REF!') >= 0 })
        })
        break
      case 'tables':
        for (const ws of sheetsOf()) {
          each(ws.ListObjects, t => {
            const headers = []
            each(t.ListColumns, c => { if (headers.length < 30) headers.push(c.Name) })
            let rows = 0
            try { rows = t.ListRows.Count } catch (e) { /* 空表格 */ }
            items.push({ name: t.Name, sheet: ws.Name, range: this._address(t.Range, true), rows, columns: headers })
          })
        }
        break
      case 'pivots':
        for (const ws of sheetsOf()) {
          let pivots
          try { pivots = ws.PivotTables() } catch (e) { continue }
          each(pivots, p => {
            let location = null; let source = null
            try { location = this._address(p.TableRange1, true) } catch (e) { /* 读不到 */ }
            try { source = Paging.clipText(p.SourceData, 200) } catch (e) { /* 读不到 */ }
            items.push({ name: p.Name, sheet: ws.Name, location, source })
          })
        }
        break
      case 'charts':
        for (const ws of sheetsOf()) {
          let charts
          try { charts = ws.ChartObjects() } catch (e) { continue }
          each(charts, co => {
            let title = null; let anchor = null
            try { if (co.Chart.HasTitle) title = co.Chart.ChartTitle.Text } catch (e) { /* 无标题 */ }
            try { anchor = this._address(co.TopLeftCell, true) } catch (e) { /* 读不到 */ }
            let type = null
            try { type = String(co.Chart.ChartType) } catch (e) { /* 读不到 */ }
            items.push({ name: co.Name, sheet: ws.Name, type, title, top_left: anchor, chart_sheet: false })
          })
        }
        break
      default:
        return { error: 'kind 只能是 sheets / names / tables / pivots / charts', suggestion: '例如 list(kind=tables)' }
    }
    return { kind: k, count: items.length, truncated: full(), items }
  },

  /**
   * 一张表的有界快照（对应 C# SheetSnapshot），交给侧车 perception 包做结构分析。
   * 格式见 src/DeepExcel.Sidecar/perception/grid.py。WPS 的 Value2 分不出日期，
   * 所以按列读一次 NumberFormat，日期格式的列把序列号转成日期。溢出区域不读（spills 为空）。
   */
  sheetSnapshot(sheetName, maxCells) {
    const wb = wps.Application.ActiveWorkbook
    if (!wb) return { error: '没有打开的工作簿' }
    let ws = null
    if (sheetName && String(sheetName).trim()) {
      const wanted = String(sheetName).trim().toLowerCase()
      for (let i = 1; i <= wb.Worksheets.Count; i++) {
        const candidate = wb.Worksheets(i)
        if (String(candidate.Name).toLowerCase() === wanted) { ws = candidate; break }
      }
      if (!ws) return { error: '找不到工作表：' + sheetName, suggestion: '先用 list(kind=sheets) 看有哪些表' }
    } else {
      ws = wps.Application.ActiveSheet
      if (!ws) return { error: '当前没有活动的工作表', suggestion: '传入 sheet 参数' }
    }

    const used = ws.UsedRange
    const row1 = used.Row
    const col1 = used.Column
    const totalRows = used.Rows.Count
    const totalColumns = used.Columns.Count
    const usedValue = totalRows === 1 && totalColumns === 1 ? used.Value2 : 1
    const empty = totalRows === 1 && totalColumns === 1 && (usedValue === null || usedValue === undefined || usedValue === '') && !used.HasFormula
    const window = Paging.planSnapshotWindow(totalRows, totalColumns, maxCells)
    const cells = []
    const formulas = []
    const merges = []
    let formulasTruncated = false
    let mergesTruncated = false

    if (!empty && window.rows > 0) {
      const range = ws.Range(ws.Cells(row1, col1), ws.Cells(row1 + window.rows - 1, col1 + window.columns - 1))
      const dateColumns = []
      for (let c = 1; c <= window.columns; c++) {
        let format = null
        try { format = range.Columns(c).NumberFormat } catch (e) { format = null }
        dateColumns.push(Paging.isDateFormat(format))
      }
      const values = this._as2d(range.Value2)
      for (let r = 0; r < window.rows; r++) {
        const row = []
        for (let c = 0; c < window.columns; c++) {
          row.push(Paging.encodeValue(values[r] ? values[r][c] : null, dateColumns[c]))
        }
        cells.push(row)
      }

      if (range.HasFormula !== false) {
        const r1c1 = this._as2d(range.FormulaR1C1)
        for (let r = 0; r < window.rows && !formulasTruncated; r++) {
          for (let c = 0; c < window.columns; c++) {
            const f = r1c1[r] ? r1c1[r][c] : null
            if (typeof f !== 'string' || f.length < 2 || f[0] !== '=') continue
            if (formulas.length >= Paging.SNAPSHOT_MAX_FORMULAS) { formulasTruncated = true; break }
            formulas.push([r, c, Paging.clipFormula(f)])
          }
        }
      }

      if (range.MergeCells !== false) {
        const seen = new Set()
        for (let r = 1; r <= window.rows && !mergesTruncated; r++) {
          let rowMerge = null
          try { rowMerge = range.Rows(r).MergeCells } catch (e) { rowMerge = null }
          if (rowMerge === false) continue
          for (let c = 1; c <= window.columns; c++) {
            const cell = range.Cells(r, c)
            if (cell.MergeCells !== true) continue
            const area = cell.MergeArea
            const key = this._address(area, true)
            const areaColumns = area.Columns.Count
            if (!seen.has(key)) {
              seen.add(key)
              if (merges.length >= Paging.SNAPSHOT_MAX_MERGES) { mergesTruncated = true; break }
              merges.push([area.Row - row1, area.Column - col1,
                area.Row + area.Rows.Count - 1 - row1, area.Column + areaColumns - 1 - col1])
            }
            c = area.Column + areaColumns - 1 - col1 + 1  // 跳过这块合并区域剩下的列
          }
        }
      }
    }

    const onThisSheet = kind => {
      const listed = this.listObjects(kind)
      return (listed.items || []).filter(item => item.sheet === ws.Name)
    }
    return {
      sheet: ws.Name,
      used: empty ? null : this._address(used, true),
      origin: [row1, col1],
      total_rows: empty ? 0 : totalRows,
      total_columns: empty ? 0 : totalColumns,
      truncated: !empty && (window.rows < totalRows || window.columns < totalColumns),
      cells,
      formulas,
      formulas_truncated: formulasTruncated,
      merges,
      merges_truncated: mergesTruncated,
      spills: [],
      objects: { tables: onThisSheet('tables'), charts: onThisSheet('charts'), pivots: onThisSheet('pivots') },
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
    const bang = address.lastIndexOf('!')
    if (bang > 0) {
      // 'Q1 销售'!A1 这种带引号的表名：去掉外层引号，'' 还原成 '
      let sheetName = address.slice(0, bang).trim()
      if (sheetName.length >= 2 && sheetName[0] === "'" && sheetName[sheetName.length - 1] === "'") {
        sheetName = sheetName.slice(1, -1).replace(/''/g, "'")
      }
      const ws = app.ActiveWorkbook.Worksheets(sheetName)
      return ws.Range(address.slice(bang + 1))
    }
    return app.Range(address)
  },

  /**
   * 获取范围的所有 NumberFormat（2D 数组）。
   * 整块格式一致时 range.NumberFormat 直接是字符串；不一致时是 null，再按列取，
   * 只有列内也混用格式时才逐格读——一页 1 万格逐格跨 COM 读要好几秒。
   */
  _getNumberFormats(range) {
    try {
      const rowCount = range.Rows.Count
      const colCount = range.Columns.Count
      const uniform = range.NumberFormat
      const columns = []
      for (let c = 1; c <= colCount; c++) {
        let column = typeof uniform === 'string' ? uniform : null
        if (column == null) {
          try { column = range.Columns(c).NumberFormat } catch (e) { column = null }
        }
        if (typeof column === 'string') {
          columns.push(new Array(rowCount).fill(column))
        } else {
          const cells = []
          for (let r = 1; r <= rowCount; r++) cells.push(range.Cells(r, c).NumberFormat)
          columns.push(cells)
        }
      }
      const formats = []
      for (let r = 0; r < rowCount; r++) formats.push(columns.map(col => col[r]))
      return formats
    } catch (e) {
      return null
    }
  },

  /** 地址文本：relative=true 去掉 $（WPS 的 Address 有时是属性、有时可带参调用，两种都兼容） */
  _address(range, relative) {
    let text
    try {
      text = typeof range.Address === 'function' ? range.Address(!relative, !relative) : range.Address
    } catch (e) {
      text = range.Address
    }
    text = String(text == null ? '' : text)
    return relative ? text.replace(/\$/g, '') : text
  },

  /** Value2 / Formula 单格时是标量，多格时是二维数组；统一成二维数组 */
  _as2d(value) {
    if (Array.isArray(value)) return value.map(row => (Array.isArray(row) ? row : [row]))
    return [[value]]
  },

  /** Worksheet.Visible：-1/true 可见，0/false 隐藏，2 深度隐藏 */
  _visibility(visible) {
    if (visible === -1 || visible === true) return 'visible'
    if (visible === 2) return 'very_hidden'
    return 'hidden'
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
