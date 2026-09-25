---
name: vba-writing-debugging
title: VBA 编写与调试
description: 用 execute_vba 写宏之前读：这里的执行方式、写法规范（不用 Select、数组批量读写）、禁用的东西、运行时错误号怎么查
version: 1
hosts: excel
tools: execute_vba
---

# VBA 编写与调试（execute_vba）

能用专门工具（write_formula、sort_data、clean_amount……）完成的，不要写 VBA：专门工具有写前检查、自动备份和写后体检。
VBA 用在专门工具做不到的批量、循环、条件逻辑上。

## 一、这里是怎么执行的

- 代码里没有 `Sub` 声明时，会被包进 `Sub DeepExcel_TempMacro()` 执行；写了多个过程时，入口是 `DeepExcel_TempMacro`（或唯一的那个 Sub）。
- 执行前先编译：编译错误会带着出错行返回，不会执行。
- 运行时错误会返回错误号、描述和行号（有行号标签时）。
- 执行器会保存并恢复 ScreenUpdating、EnableEvents、DisplayAlerts、计算模式和状态栏：**不用自己在结尾恢复**，
  但中途报错也不会留下「屏幕不刷新」的烂摊子。
- 执行期间弹出的对话框会被自动关掉并记录下来；**不要用 MsgBox / InputBox**（会被拦下），结果靠执行后读回区域确认。
- 文件、网络、Shell 相关的对象（FileSystemObject、WScript.Shell、XMLHTTP、ADODB.Stream 等）和写文件会被拦截。
- 中文字符串可以直接写，执行器会处理代码页。
- 执行前会自动备份；改错了可以 rollback。

## 二、写法规范

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("销售明细")     ' 显式指定工作表，不依赖当前活动表
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

Dim data As Variant, r As Long
data = ws.Range("A2:F" & lastRow).Value        ' 一次读进数组
For r = 1 To UBound(data, 1)
    If Trim$(CStr(data(r, 3))) = "" Then data(r, 3) = "未填"
Next r
ws.Range("A2:F" & lastRow).Value = data        ' 一次写回
```

- **不要 Select / Activate**，直接对对象操作；所有 Range、Cells 都带上 `ws.` 前缀。
- 大量数据用数组批量读写，逐格读写几万行会慢几十倍。
- 行号、计数用 `Long`，不用 `Integer`（超过 32767 溢出，错误 6）。
- 日期用 `DateSerial(年, 月, 日)` 构造，不要拼字符串再转。
- `On Error Resume Next` 只包住确实可能失败的那一行，之后立刻 `On Error GoTo 0`；不要让它盖住整段代码——
  出了错会悄悄继续，把错误结果写进表里。
- 数组从单列区域读出来也是二维的：`data(r, 1)`；单格区域的 `.Value` 不是数组，先判断行数。
- 删除行从下往上删：`For r = lastRow To 2 Step -1`。
- 写完回读几格确认（read_range），不要只凭「执行成功」就汇报完成。

## 三、运行时错误号

| 错误号 | 意思 | 常见原因 |
|---|---|---|
| 1004 | 应用程序定义或对象定义错误 | 区域地址不对、工作表受保护、公式字符串写错、对合并单元格做了不允许的操作 |
| 9 | 下标越界 | 工作表名不存在（注意全角半角、首尾空格）、数组下标超出 |
| 13 | 类型不匹配 | 单元格里是错误值（#N/A）或文本却参与运算；用 `IsError` / `IsNumeric` 先判断 |
| 91 | 对象变量或 With 块变量未设置 | `Find` 没找到返回 Nothing 后直接用了它的属性 |
| 424 | 要求对象 | 给对象赋值忘了 `Set` |
| 438 | 对象不支持该属性或方法 | 属性名拼错，或者这个对象没有这个方法 |
| 6 | 溢出 | 用了 Integer；或者除以 0 的结果 |

`Find` 的正确用法：

```vba
Dim hit As Range
Set hit = ws.Columns("A").Find(What:="应收", LookIn:=xlValues, LookAt:=xlWhole)
If Not hit Is Nothing Then hit.Offset(0, 1).Value = "已找到"
```

## 四、调试步骤

1. 看返回的错误号和行号，对照上表。
2. 缩小范围：先在 3–5 行数据上跑，确认逻辑对了再放开到全部。
3. 同一个错误连续出现两次，换思路（改用专门工具，或者换一种写法），不要反复小改重试。
4. 不确定表结构时先读（inspect_sheet / read_range），不要在 VBA 里猜列位置。

## 常见报错

| 现象 | 常见原因 | 怎么办 |
|---|---|---|
| 编译错误：变量未定义 | 模块开了 Option Explicit，变量没 Dim | 补上 Dim |
| 编译错误：缺少 End Sub / Next / End If | 块没闭合 | 按返回的行号数一下块的配对 |
| 执行成功但表没变 | 操作的是别的工作表（没带 ws. 前缀，落到了活动表上） | 所有 Range 带上工作表对象 |
| 结果只处理了一部分行 | 最后一行取错（用了有空行的列算 lastRow） | 换一列算，或用 UsedRange 核对 |
