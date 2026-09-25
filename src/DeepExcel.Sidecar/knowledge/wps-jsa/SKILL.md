---
name: wps-jsa
title: WPS JSA 宏
description: 在 WPS 里用 execute_jsa 写宏之前读：JS 语法 + VBA 对象模型的写法差异、常量、日期、批量读写和调试
version: 1
hosts: wps
tools: execute_jsa
---

# WPS JSA 宏（execute_jsa）

WPS 表格里没有 VBA 时用 JSA：语法是 JavaScript（ES6），对象模型和 VBA 基本一致（Application / Workbook / Worksheet / Range）。
能用专门工具完成的，不要写 JSA：专门工具有写前检查和写后体检。

## 一、执行方式

- 代码直接执行，**最后一个表达式的值会作为结果返回**。最后放一个简短的核对结果，方便确认：
  ```js
  const ws = Application.ActiveWorkbook.Worksheets("明细");
  // ……处理……
  `处理了 ${n} 行，第一行结果：${ws.Range("D2").Text}`
  ```
- 抛出的异常会作为错误返回。不要用 `alert` / `MsgBox`，结果靠返回值和回读区域确认。

## 二、和 VBA 写法的差异

| VBA | JSA |
|---|---|
| `Worksheets("明细")` | `Worksheets("明细")` 或 `Worksheets.Item("明细")` |
| `Set ws = ...` | `const ws = ...`（没有 Set） |
| `ws.Range("A1").Value = 1` | `ws.Range("A1").Value2 = 1` |
| `Dim i As Long` | `let i` |
| `For i = 2 To n` | `for (let i = 2; i <= n; i++)` |
| `If x = "" Then` | `if (x === "")`（空单元格还要判断 null，见表下） |
| `"a" & "b"` | `"a" + "b"` 或模板字符串 |
| `Nothing` | `null` / `undefined` |
| 命名参数 `Find(What:="x")` | 按位置传参，不支持命名参数 |

- 行列序号从 1 开始（`Cells(行, 列)`），和 VBA 一样；**JS 数组下标从 0 开始**，两者混用时最容易错一位。
- 枚举常量（xlUp、xlToLeft……）在 JSA 里未必可用，直接写数值：xlUp = -4162、xlDown = -4121、xlToLeft = -4159、xlToRight = -4161。
  最后一行：`ws.Cells(ws.Rows.Count, 1).End(-4162).Row`。
- 单元格为空时读出来是 `undefined`、`null` 或空字符串，都要判断：`if (v === undefined || v === null || v === "")`。
- 数字和字符串：`"10" + 1` 得到 `"101"`；参与计算前用 `Number(x)` 转换，并检查 `isNaN`。

## 三、日期

- `Value2` 读出来是序列号（天数），不是 JS 的 Date。转换：
  `new Date(Math.round((serial - 25569) * 86400000))`（得到的是 UTC 时间，取年月日用 `getUTCFullYear` 等）。
- 写日期：直接写序列号，再设数字格式（`NumberFormat = "yyyy-mm-dd"`）；或写 `"2026/9/25"` 这样的字符串让表格识别——
  前者更可靠。

## 四、批量读写

逐格读写很慢。多格区域可以整块读写，但不同 WPS 版本返回的形态可能不同：**先在一个小区域上读一次，看返回的是不是二维数组**，
确认后再对整块操作；拿不准时用逐行循环，数据量大就分批。

```js
const ws = Application.ActiveWorkbook.Worksheets("明细");
const last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row;
let changed = 0;
for (let r = 2; r <= last; r++) {
  const cell = ws.Cells(r, 3);
  const v = cell.Value2;
  if (v === undefined || v === null || String(v).trim() === "") { cell.Value2 = "未填"; changed++; }
}
`填充了 ${changed} 格（共 ${last - 1} 行）`
```

## 五、调试

1. 报错信息里有属性名的，多半是这个对象在当前 WPS 版本里没有这个属性：换一种写法，或者改用专门工具。
2. 先在 3–5 行上跑通再放开到全部。
3. 同一个错误出现两次就换思路，不要反复小改。
4. 做完用 read_range 回读确认，不要只看返回值。

## 常见报错

| 现象 | 常见原因 | 怎么办 |
|---|---|---|
| xlUp is not defined | 枚举常量不可用 | 写数值 -4162 |
| 结果差一行 / 一列 | JS 数组下标从 0、单元格从 1 | 统一换算，回读首末行确认 |
| 数字变成拼接的字符串 | 单元格值是文本，`+` 做了拼接 | Number() 转换 |
| 日期显示成 46290 | 写入了序列号但没设日期格式 | 设 NumberFormat |
