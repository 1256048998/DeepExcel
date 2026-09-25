/**
 * 首次使用：空工作簿给示例数据 + 三条示范问题；有数据的工作簿按结构给「为你的文件推荐」。
 *
 * 宿主（Excel / WPS）只把每张表的前几十行原样交过来（get_starter → starter），
 * 列类型判断和推荐全在这里做：两个宿主共用一份逻辑，推荐不花模型调用、打开面板就有。
 * 示例数据也在这里生成，宿主只负责把它写进一张新表（insert_sample），不动用户已有的表。
 */

export type Cell = string | number | boolean | null

export interface OutlineSheet {
  name: string
  /** 已用区域的总行数（含表头） */
  rows: number
  columns: number
  /** 已用区域左上角开始的前若干行，第一行当表头 */
  grid: Cell[][]
  /** 每列第二行的数字格式像日期（Value2 里日期只是序列号） */
  date_columns?: boolean[]
}

export interface WorkbookOutline {
  workbook_name: string
  sheets: OutlineSheet[]
}

export type ColumnKind = 'number' | 'date' | 'text' | 'mixed' | 'empty'

export interface ColumnProfile {
  header: string
  kind: ColumnKind
  /** 看起来是数字、实际存成文本的格 */
  textNumbers: number
  /** 首尾有空格的文本 */
  padded: number
  blanks: number
  distinct: number
  filled: number
}

export interface SheetProfile {
  name: string
  dataRows: number
  columns: ColumnProfile[]
  errors: number
  duplicateRows: number
}

export interface StarterSuggestion {
  title: string
  prompt: string
}

export interface StarterPlan {
  mode: 'empty' | 'recommend'
  suggestions: StarterSuggestion[]
}

const AMOUNT_HEADER = /金额|销售额|收入|营收|成本|费用|工资|应发|实发|合计|总额|单价|价格|余额|应收|应付|利润|数量|销量|回款/
const CATEGORY_HEADER = /区域|地区|部门|门店|城市|省|客户|供应商|产品|品类|类别|类型|渠道|项目|销售员|业务员|负责人|科目/
const DATE_HEADER = /日期|时间|月份|年月|date/i
const TEXT_NUMBER = /^\s*[-+]?[¥￥$]?\s*\d{1,3}(,\d{3})+(\.\d+)?\s*(元)?\s*$|^\s*[-+]?[¥￥$]?\s*\d+(\.\d+)?\s*(元)?\s*$/
const ERROR_VALUE = /^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|UNKNOWN!|FIELD!|ERROR)$/

function isBlank(v: Cell): boolean {
  return v === null || v === undefined || (typeof v === 'string' && v.trim() === '')
}

export function profileSheet(sheet: OutlineSheet): SheetProfile {
  const grid = sheet.grid || []
  const header = grid[0] || []
  const body = grid.slice(1)
  const width = Math.max(header.length, ...body.map(r => r.length), 0)
  let errors = 0
  const columns: ColumnProfile[] = []

  for (let c = 0; c < width; c++) {
    const values = body.map(r => (r[c] === undefined ? null : r[c]))
    let numbers = 0, texts = 0, textNumbers = 0, padded = 0, blanks = 0
    const seen = new Set<string>()
    for (const v of values) {
      if (isBlank(v)) { blanks++; continue }
      seen.add(String(v).trim())
      if (typeof v === 'number') { numbers++; continue }
      if (typeof v === 'string') {
        if (ERROR_VALUE.test(v)) { errors++; continue }
        texts++
        if (TEXT_NUMBER.test(v)) textNumbers++
        if (v !== v.trim()) padded++
      }
    }
    const filled = values.length - blanks
    const headerText = isBlank(header[c]) ? '' : String(header[c]).trim()
    let kind: ColumnKind
    if (filled === 0) kind = 'empty'
    else if (numbers > 0 && numbers >= filled * 0.8 && (sheet.date_columns?.[c] || DATE_HEADER.test(headerText))) kind = 'date'
    // 文本型数字也算数字列：它们正是要被清洗的那几个格
    else if (numbers + textNumbers >= filled * 0.8) kind = 'number'
    else if (texts === filled) kind = 'text'
    else kind = 'mixed'
    columns.push({ header: headerText || `第 ${c + 1} 列`, kind, textNumbers, padded, blanks, distinct: seen.size, filled })
  }

  const keys = body.filter(r => r.some(v => !isBlank(v))).map(r => JSON.stringify(r.map(v => (typeof v === 'string' ? v.trim() : v))))
  const duplicateRows = keys.length - new Set(keys).size

  return { name: sheet.name, dataRows: Math.max(0, sheet.rows - 1), columns, errors, duplicateRows }
}

function quote(name: string): string {
  return `「${name}」`
}

function pickAmount(p: SheetProfile): ColumnProfile | undefined {
  const numeric = p.columns.filter(c => c.kind === 'number')
  return numeric.find(c => AMOUNT_HEADER.test(c.header) && !/单价|价格/.test(c.header)) ??
    numeric.find(c => AMOUNT_HEADER.test(c.header)) ?? numeric[numeric.length - 1]
}

function pickCategory(p: SheetProfile): ColumnProfile | undefined {
  const texts = p.columns.filter(c => c.kind === 'text' && c.filled > 0)
  // 表头像分类的，只要取值不太多就行；认不出表头的，要有明显的重复才像分类
  const named = (c: ColumnProfile) => c.distinct >= 2 && c.distinct <= 30
  const repeated = (c: ColumnProfile) => c.distinct >= 2 && c.distinct <= Math.min(30, c.filled * 0.6)
  return texts.find(c => CATEGORY_HEADER.test(c.header) && named(c)) ?? texts.find(repeated)
}

function hasData(p: SheetProfile): boolean {
  return p.dataRows > 0 && p.columns.some(c => c.filled > 0)
}

/** 按工作簿结构给出最多 3 条推荐。纯函数。 */
export function recommendStarters(outline: WorkbookOutline | null | undefined): StarterPlan {
  const profiles = (outline?.sheets || []).map(profileSheet).filter(hasData)
  if (profiles.length === 0) return { mode: 'empty', suggestions: [] }

  const suggestions: StarterSuggestion[] = []
  const add = (s: StarterSuggestion) => {
    if (suggestions.length < 3 && !suggestions.some(x => x.prompt === s.prompt)) suggestions.push(s)
  }
  const main = [...profiles].sort((a, b) => b.dataRows - a.dataRows)[0]

  // 1. 数据问题：最常见、最该先做，也最能让用户看到「它读懂了我的表」
  for (const p of profiles) {
    const dirty = p.columns.filter(c => c.textNumbers > 0 || c.padded > 0)
    if (dirty.length === 0 && p.duplicateRows === 0) continue
    const issues: string[] = []
    const textNum = dirty.filter(c => c.textNumbers > 0)
    if (textNum.length) issues.push(`${textNum.slice(0, 2).map(c => quote(c.header)).join('、')}有文本型数字`)
    const pad = dirty.filter(c => c.padded > 0)
    if (pad.length) issues.push(`${pad.slice(0, 2).map(c => quote(c.header)).join('、')}有多余空格`)
    if (p.duplicateRows > 0) issues.push('有重复行')
    add({
      title: `清洗${quote(p.name)}：${issues.join('，')}`,
      prompt: `检查${quote(p.name)}的数据问题（${issues.join('；')}），先把问题逐条列出来，确认后再清洗，不要改动其他列`,
    })
    break
  }

  // 2. 公式错误
  const withErrors = profiles.find(p => p.errors > 0)
  if (withErrors) {
    add({
      title: `修复${quote(withErrors.name)}里的公式错误`,
      prompt: `找出${quote(withErrors.name)}里所有公式错误（#N/A、#REF! 等），说明每个错误的原因，再给出修改方案`,
    })
  }

  // 3. 分类汇总
  const amount = pickAmount(main)
  const category = pickCategory(main)
  if (amount && category) {
    add({
      title: `按${quote(category.header)}汇总${quote(amount.header)}`,
      prompt: `按${quote(category.header)}汇总${quote(main.name)}的${quote(amount.header)}，结果放到一张新工作表，并画一张柱状图`,
    })
  }

  // 4. 按月趋势
  const date = main.columns.find(c => c.kind === 'date')
  if (date && amount) {
    add({
      title: `按月看${quote(amount.header)}的趋势`,
      prompt: `按${quote(date.header)}的月份汇总${quote(main.name)}的${quote(amount.header)}，放到新工作表并画折线图`,
    })
  }

  // 5. 结构相同的多张表
  const signature = (p: SheetProfile) => p.columns.map(c => c.header).join('|')
  const groups = new Map<string, SheetProfile[]>()
  for (const p of profiles) groups.set(signature(p), [...(groups.get(signature(p)) || []), p])
  const same = [...groups.values()].find(g => g.length >= 2)
  if (same) {
    add({
      title: `把 ${same.length} 张结构相同的表合并成一张`,
      prompt: `${same.map(p => quote(p.name)).join('、')}表头相同，把它们合并到一张新工作表，加一列标明来源表`,
    })
  }

  // 兜底：先让它讲讲这个工作簿
  add({
    title: '先看看这个工作簿里都有什么',
    prompt: '先帮我看看这个工作簿有哪些表、各自是什么数据，再给我三条可以做的分析建议',
  })
  if (amount && !category) {
    add({
      title: `给${quote(amount.header)}做个统计`,
      prompt: `统计${quote(main.name)}的${quote(amount.header)}：合计、平均、最大最小值，并指出明显异常的行`,
    })
  }

  return { mode: 'recommend', suggestions }
}

// ============ 示例数据 ============

export const SAMPLE_SHEET_NAME = '示例-销售明细'
export const SAMPLE_HEADERS = ['订单日期', '区域', '销售员', '产品', '数量', '单价', '金额']
/** 各列的数字格式（列号从 0 开始） */
export const SAMPLE_NUMBER_FORMATS: Record<number, string> = { 0: 'yyyy-mm-dd', 5: '0.00', 6: '0.00' }

/** 固定种子：每次插入的示例一模一样，示范问题的结果才可预期 */
function seeded(seed: number): () => number {
  let s = seed >>> 0
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0
    return s / 4294967296
  }
}

/** 2024-07-01 的 Excel 日期序列号 */
const JULY_1_2024 = 45474

/**
 * 合成的销售明细：40 行，7 月到 9 月，4 个区域、6 个销售员、5 种产品。
 * 刻意埋了几处真实表里常见的问题，让「清洗」这条示范问题有东西可做：
 * 两个金额存成了文本、两个销售员名字带空格、一行重复、一个区域空着。
 */
export function sampleRows(): Cell[][] {
  const rand = seeded(20240701)
  const regions = ['华东', '华南', '华北', '西南']
  const people = ['张伟', '李娜', '王芳', '刘洋', '陈静', '赵磊']
  const products: [string, number][] = [['办公椅', 420], ['升降桌', 1680], ['显示器支架', 199], ['文件柜', 860], ['会议桌', 3200]]
  const body: Cell[][] = []
  for (let i = 0; i < 40; i++) {
    const day = JULY_1_2024 + Math.floor(rand() * 92)
    const region = regions[Math.floor(rand() * regions.length)]
    const person = people[Math.floor(rand() * people.length)]
    const [product, price] = products[Math.floor(rand() * products.length)]
    const qty = 1 + Math.floor(rand() * 12)
    body.push([day, region, person, product, qty, price, qty * price])
  }
  body.sort((a, b) => (a[0] as number) - (b[0] as number))
  const rows: Cell[][] = [SAMPLE_HEADERS, ...body]
  // 埋问题：文本金额（前导撇号让宿主按文本写入）、名字带空格、重复行、空区域
  rows[5][6] = `'${(rows[5][6] as number).toLocaleString('en-US')}`
  rows[17][6] = `'${rows[17][6]}`
  rows[9][2] = `${rows[9][2]} `
  rows[26][2] = ` ${rows[26][2]}`
  rows[33][1] = ''
  rows.splice(21, 0, [...rows[20]])
  return rows
}

export function sampleQuestions(sheet: string): StarterSuggestion[] {
  const s = quote(sheet)
  return [
    { title: '按区域汇总金额并画图', prompt: `按「区域」汇总${s}的「金额」，结果放到一张新工作表，并画一张柱状图` },
    { title: '找出并清洗数据问题', prompt: `检查${s}的数据问题（文本型数字、多余空格、重复行、空值），先逐条列出来，确认后再清洗` },
    { title: '加一列提成', prompt: `给${s}加一列「提成」：金额超过 5000 的按 3%，其余按 1%，用公式计算` },
  ]
}
