// 工具步骤的中文叙事文案：「读取 Sheet1!A1:D20」而不是「read_range」。
//
// Claude Code 的每个工具调用都是一行人能读懂的话（⏺ Read(src/App.tsx)），用户扫一眼
// 就知道它在干什么、动了哪里。这里给每个已注册的工具一句同样的话。
// toolCatalog.test.ts 会拿侧车 excel_tools.py 的注册表逐条比对：新增工具没写文案，测试失败。

type Args = Record<string, unknown>
type Label = (a: Args) => string

function s(v: unknown): string {
  if (v === undefined || v === null) return ''
  if (typeof v === 'string') return v
  if (typeof v === 'number' || typeof v === 'boolean') return String(v)
  try { return JSON.stringify(v) } catch { return '' }
}

function at(prefix: string, v: unknown): string {
  const text = s(v)
  return text ? `${prefix} ${text}` : prefix
}

function clip(text: string, max = 40): string {
  const one = text.replace(/\s+/g, ' ').trim()
  return one.length > max ? one.slice(0, max) + '…' : one
}

function shape(v: unknown): string {
  if (v && typeof v === 'object' && '__shape' in (v as Args)) {
    const [r, c] = (v as { __shape: [number, number] }).__shape
    return `${r} 行 × ${c} 列`
  }
  if (Array.isArray(v)) return `${v.length} 行`
  return ''
}

function codeLines(v: unknown): string {
  const text = s(v)
  if (!text) return ''
  const n = text.split('\n').length
  return `（${n} 行）`
}

const LIST_KINDS: Record<string, string> = {
  sheets: '工作表',
  names: '定义的名称',
  tables: '表格',
  pivots: '数据透视表',
  charts: '图表',
}

// 内置知识技能的中文名（服务端下发的新技能没有条目时显示技能名）
const SKILL_TITLES: Record<string, string> = {
  'delivery-quality': '交付质量与验证',
  'cn-data-cleaning': '中文数据清洗',
  'cn-formula-writing': '中文用户的公式写入',
  'cn-financial-reconciliation': '财务报表勾稽核对',
  'cn-payroll-tax': '工资表与个税',
  'cn-attendance-roster': '考勤与花名册',
  'sales-reporting': '销售报表',
  'ar-aging-reconciliation': '应收账龄与对账',
  'vba-writing-debugging': 'VBA 编写与调试',
  'wps-jsa': 'WPS JSA 宏',
}

export const TOOL_LABELS: Record<string, Label> = {
  // 读
  read_workbook: () => '读取工作簿结构',
  read_selection: () => '读取当前选区',
  read_range: a => {
    const offset = Number(a.offset)
    return offset > 0 ? `读取 ${s(a.address)}（从第 ${offset + 1} 行起）` : at('读取', a.address)
  },
  find: a => {
    const where = Array.isArray(a.sheets) && a.sheets.length ? `（${a.sheets.map(s).join('、')}）` : ''
    return `查找${a.scope === 'formulas' ? '公式里的' : ''}「${clip(s(a.query), 20)}」${where}`
  },
  list: a => `列出${LIST_KINDS[s(a.kind)] || '工作表'}`,
  inspect_sheet: a => at('分析表结构', a.sheet),
  explore_workbook: a => `分头摸底（${Array.isArray(a.tasks) ? a.tasks.length : 1} 个子任务）`,
  read_attachment: a => at('读取附件', a.file_name),

  // 写
  write_formula: a => `写入公式 ${s(a.address)} ${clip(s(a.formula), 30)}`.trim(),
  write_value: a => `写入 ${s(a.address)} = ${clip(s(a.value), 30)}`.trim(),
  write_range: a => `批量写入 ${s(a.address)} ${shape(a.values)}`.trim(),
  fill_formula_down: a => `向下填充公式 ${s(a.from_address)}${a.row_count ? ` × ${s(a.row_count)} 行` : ''}`,
  replace_formula: a => `替换公式 ${s(a.range_address)}：${clip(s(a.find), 16)} → ${clip(s(a.replace), 16)}`,
  copy_range: a => `复制 ${s(a.source_address)} → ${s(a.dest_address)}`,
  clear_range: a => `清空 ${s(a.address)}${a.clear_type && a.clear_type !== 'all' ? `（${a.clear_type === 'formats' ? '格式' : '内容'}）` : ''}`,

  // 清洗
  clean_data: a => `清洗数据 ${s(a.range_address)}`,
  delete_blank_rows: a => at('删除空行', a.range_address),
  split_text_to_columns: a => `分列 ${s(a.range_address)}${a.delimiter ? `（按「${s(a.delimiter)}」）` : ''}`,
  fill_blank_cells: a => at('向下填充空白', a.range_address),
  highlight_duplicates: a => at('标记重复值', a.range_address),
  remove_special_chars: a => at('去除特殊字符', a.range_address),
  clean_amount: a => at('清洗金额', a.range_address),
  merge_columns: a => `合并列 ${s(a.range_address)}${a.target_column ? ` → ${s(a.target_column)}` : ''}`,
  rename_columns: a => at('重命名列标题', a.range_address),
  collapse_spaces: a => at('压缩多余空格', a.range_address),

  // 图表
  create_chart: a => `创建图表 ${s(a.data_range)}${a.title ? `「${clip(s(a.title), 20)}」` : ''}`,
  create_combo_chart: a => `创建组合图 ${s(a.data_range)}${a.title ? `「${clip(s(a.title), 20)}」` : ''}`,
  add_data_labels: a => at('添加数据标签', a.chart_name),
  set_chart_title: a => `设置图表标题「${clip(s(a.title), 20)}」`,
  set_chart_colors: a => at('设置图表颜色', a.chart_name),
  export_chart: a => at('导出图表', a.output_path || a.chart_name),

  // 透视表
  create_pivot_table: a => `创建透视表 ${s(a.source_range)}${a.destination_sheet ? ` → ${s(a.destination_sheet)}` : ''}`,
  refresh_pivot: a => at('刷新透视表', a.pivot_table_name),
  group_pivot_date: a => `透视表按日期分组 ${s(a.field_name)}（${s(a.group_by)}）`,
  set_pivot_value_display: a => `设置透视表值显示 ${s(a.value_field)}`,
  set_pivot_totals: a => at('设置透视表总计', a.pivot_table_name),
  add_pivot_slicer: a => `添加切片器 ${s(a.field_name)}`,

  // 代码
  execute_vba: a => `运行 VBA${codeLines(a.code)}`,
  execute_jsa: a => `运行 JS 宏${codeLines(a.code)}`,
  execute_python: a => `运行 Python 计算${codeLines(a.code)}`,

  // 快照
  create_snapshot: () => '创建快照',
  rollback: () => '恢复到快照',

  // 工作表与结构
  add_sheet: a => at('新建工作表', a.name),
  delete_sheet: a => at('删除工作表', a.name),
  rename_sheet: a => `重命名工作表 ${s(a.old_name)} → ${s(a.new_name)}`,
  insert_rows: a => `在第 ${s(a.row)} 行前插入 ${s(a.count) || 1} 行`,
  delete_rows: a => `从第 ${s(a.row)} 行起删除 ${s(a.count) || 1} 行`,
  insert_columns: a => `在第 ${s(a.column)} 列前插入 ${s(a.count) || 1} 列`,
  delete_columns: a => `从第 ${s(a.column)} 列起删除 ${s(a.count) || 1} 列`,
  freeze_panes: a => at('冻结窗格', a.address),

  // 格式
  set_number_format: a => `设置数字格式 ${s(a.address)} ${s(a.format)}`.trim(),
  set_column_width: a => `${a.auto_fit ? '自动调整列宽' : '设置列宽'} ${s(a.address)}`.trim(),
  set_cell_style: a => at('设置样式', a.address),
  merge_cells: a => at('合并单元格', a.address),
  unmerge_cells: a => at('取消合并', a.address),
  apply_conditional_format: a => at('条件格式', a.address),
  write_table: a => `转为表格 ${s(a.address)}${a.table_name ? `「${s(a.table_name)}」` : ''}`,
  sort_data: a => `排序 ${s(a.range_address)}（按 ${s(a.sort_column)} ${a.descending ? '降序' : '升序'}）`,
  filter_data: a => `筛选 ${s(a.range_address)}${a.criteria ? `（${clip(s(a.criteria), 16)}）` : ''}`,

  // 交互
  clarify_intent: () => '向你确认需求',
  todo_write: () => '更新计划',
  update_workbook_notes: () => '更新工作簿记忆',
  load_skill: a => `查阅知识「${SKILL_TITLES[s(a.name)] || s(a.name)}」`,
  screenshot_excel: () => '截取 Excel 窗口',
  send_keys: a => at('发送按键', a.keys),
}

export function toolLabel(name: string, args?: Args): string {
  const bare = (name || '').replace(/^mcp__excel__/, '')
  const label = TOOL_LABELS[bare]
  if (!label) return bare
  try {
    return label(args ?? {})
  } catch {
    return bare
  }
}
