// 选区条（输入框上方「▦ Sheet1!A1:D20 · 80 格  ×」）：宿主推 selection_brief，这里决定显示什么。
// 纯函数，有测试：selection.test.ts

export type SelectionBrief = { sheet?: string; address: string; rows?: number; cols?: number; cells?: number }

/** 只选了一个格通常只是光标停在那里，不值得占一行；多于一格才显示。 */
export function shouldShowSelection(sel: SelectionBrief | null | undefined): sel is SelectionBrief {
  return !!sel && !!sel.address && (sel.cells ?? 0) > 1
}

/** 「Sheet1!A1:D20」；工作表名带空格或符号时加引号，和 Excel 的写法一致。 */
export function selectionLabel(sel: SelectionBrief): string {
  if (!sel.sheet) return sel.address
  const sheet = /^[\p{L}\p{N}_]+$/u.test(sel.sheet) ? sel.sheet : `'${sel.sheet.replace(/'/g, "''")}'`
  return `${sheet}!${sel.address}`
}

export function cellCountText(cells: number): string {
  if (cells >= 1e8) return `${(cells / 1e8).toFixed(1).replace(/\.0$/, '')} 亿格`
  if (cells >= 1e4) return `${(cells / 1e4).toFixed(1).replace(/\.0$/, '')} 万格`
  return `${cells} 格`
}

/** 同一块选区再推一次（例如切回面板）不算变化，用户点过的 × 仍然有效。 */
export function sameSelection(a: SelectionBrief | null, b: SelectionBrief | null): boolean {
  return !!a && !!b && a.sheet === b.sheet && a.address === b.address
}
