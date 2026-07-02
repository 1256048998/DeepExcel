# DeepExcel VBA功能测试用例

> 来源：ExcelHome论坛、MrExcel、CSDN等用户高频VBA问题整理
> 日期：2026-06-29

---

## 一、基础VBA执行

### 测试用例 V-001：单元格格式设置
- **用户需求**：`用VBA把A1到A10的字体改成红色，加粗`
- **预期VBA**：
```vba
Sub SetFontRed()
    Range("A1:A10").Font.Color = vbRed
    Range("A1:A10").Font.Bold = True
End Sub
```
- **验证点**：格式正确应用、代码可执行

### 测试用例 V-002：背景色设置
- **用户需求**：`给选中的单元格添加黄色背景`
- **预期VBA**：
```vba
Sub SetYellowBackground()
    Selection.Interior.Color = vbYellow
End Sub
```
- **验证点**：背景色正确

### 测试用例 V-003：条件格式（VBA方式）
- **用户需求**：`把大于100的数字标红`
- **预期VBA**：
```vba
Sub HighlightAbove100()
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) And cell.Value > 100 Then
            cell.Font.Color = vbRed
        End If
    Next cell
End Sub
```
- **验证点**：只修改符合条件的单元格

---

## 二、数据处理VBA

### 测试用例 V-010：删除空行
- **用户需求**：`删除A列为空的所有行`
- **预期VBA**：
```vba
Sub DeleteEmptyRows()
    Dim i As Long
    For i = Cells(Rows.Count, "A").End(xlUp).Row To 1 Step -1
        If IsEmpty(Cells(i, "A")) Or Cells(i, "A").Value = "" Then
            Rows(i).Delete
        End If
    Next i
End Sub
```
- **验证点**：从下往上删除避免行号错位

### 测试用例 V-011：数据清洗（去空格）
- **用户需求**：`清除A列所有单元格的前后空格`
- **预期VBA**：
```vba
Sub TrimSpaces()
    Dim cell As Range
    For Each cell In Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row)
        If Not IsEmpty(cell) Then
            cell.Value = Trim(cell.Value)
        End If
    Next cell
End Sub
```
- **验证点**：正确去除首尾空格

### 测试用例 V-012：批量替换
- **用户需求**：`把A列所有的"北京"替换成"BJ"`)
- **预期VBA**：
```vba
Sub BatchReplace()
    Range("A:A").Replace What:="北京", Replacement:="BJ", LookAt:=xlPart
End Sub
```
- **验证点**：正确替换、支持部分匹配

### 测试用例 V-013：文本转数字
- **用户需求**：`A列的数字是文本格式，帮我转成真正的数字`
- **预期VBA**：
```vba
Sub TextToNumber()
    Dim cell As Range
    For Each cell In Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row)
        If IsNumeric(cell.Value) And cell.Value <> "" Then
            cell.Value = CDbl(cell.Value)
        End If
    Next cell
End Sub
```
- **验证点**：正确转换、非数字跳过

---

## 三、循环与遍历

### 测试用例 V-020：遍历所有工作表
- **用户需求**：`在每个工作表的A1写入当前日期`
- **预期VBA**：
```vba
Sub WriteDateToAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Range("A1").Value = Date
    Next ws
End Sub
```
- **验证点**：所有工作表都被处理

### 测试用例 V-021：遍历区域
- **用户需求**：`把A列大于100的值复制到B列`
- **预期VBA**：
```vba
Sub CopyLargeValues()
    Dim i As Long, lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 1 To lastRow
        If IsNumeric(Cells(i, "A").Value) And Cells(i, "A").Value > 100 Then
            Cells(i, "B").Value = Cells(i, "A").Value
        End If
    Next i
End Sub
```
- **验证点**：正确判断条件、复制到正确位置

### 测试用例 V-022：Do While循环
- **用户需求**：`从A1开始向下读取，直到遇到空单元格`
- **预期VBA**：
```vba
Sub ReadUntilEmpty()
    Dim cell As Range
    Set cell = Range("A1")
    Do While cell.Value <> ""
        Debug.Print cell.Value
        Set cell = cell.Offset(1, 0)
    Loop
End Sub
```
- **验证点**：正确终止循环

---

## 四、文件操作

### 测试用例 V-030：打开工作簿
- **用户需求**：`打开指定路径的Excel文件`
- **预期VBA**：
```vba
Sub OpenWorkbook()
    Dim wb As Workbook
    Set wb = Workbooks.Open("C:\data\report.xlsx")
    ' 处理完成后关闭
    wb.Close SaveChanges:=False
End Sub
```
- **验证点**：文件正确打开和关闭

### 测试用例 V-031：另存为
- **用户需求**：`把当前工作簿另存为新文件`
- **预期VBA**：
```vba
Sub SaveAsNew()
    ThisWorkbook.SaveAs Filename:="C:\data\backup_" & Format(Date, "yyyymmdd") & ".xlsx"
End Sub
```
- **验证点**：文件名包含日期、格式正确

### 测试用例 V-032：合并多个工作簿
- **用户需求**：`把文件夹里所有Excel文件的第一个工作表合并到当前工作簿`
- **预期VBA**：
```vba
Sub MergeWorkbooks()
    Dim folder As String, file As String
    Dim wb As Workbook, ws As Worksheet
    folder = "C:\data\"
    file = Dir(folder & "*.xlsx")
    Do While file <> ""
        Set wb = Workbooks.Open(folder & file)
        wb.Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wb.Close SaveChanges:=False
        file = Dir()
    Loop
End Sub
```
- **验证点**：所有文件被处理、数据正确合并

---

## 五、性能优化

### 测试用例 V-040：关闭屏幕刷新
- **用户需求**：`处理10000行数据时Excel很卡，帮我优化`
- **预期VBA**：
```vba
Sub ProcessWithOptimization()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ... 处理代码 ...
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```
- **验证点**：执行速度明显提升

### 测试用例 V-041：使用数组批量写入
- **用户需求**：`把A列数据复制到B列，但要快一点`
- **预期VBA**：
```vba
Sub FastCopyWithArray()
    Dim arr As Variant
    arr = Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row).Value
    Range("B1:B" & UBound(arr, 1)).Value = arr
End Sub
```
- **验证点**：比逐单元格复制快10倍以上

---

## 六、错误处理

### 测试用例 V-050：基础错误处理
- **用户需求**：`如果工作表不存在就提示用户`
- **预期VBA**：
```vba
Sub SafeAccessSheet()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("目标表")
    ' ... 操作 ...
    Exit Sub
ErrHandler:
    MsgBox "工作表不存在，请检查！", vbExclamation
End Sub
```
- **验证点**：错误被捕获、用户看到友好提示

### 测试用例 V-051：类型错误处理
- **用户需求**：`用户输入的可能是文本数字，要能正确处理`
- **预期VBA**：
```vba
Sub HandleTypeConversion()
    Dim val As Variant
    val = InputBox("请输入数字：")
    If IsNumeric(val) Then
        MsgBox "你输入的是：" & CDbl(val)
    Else
        MsgBox "输入的不是有效数字！", vbExclamation
    End If
End Sub
```
- **验证点**：非数字输入不导致程序崩溃

---

## 七、用户界面

### 测试用例 V-060：InputBox输入
- **用户需求**：`让用户输入要处理的行数`
- **预期VBA**：
```vba
Sub GetUserInput()
    Dim numRows As String
    numRows = InputBox("请输入要处理的行数：", "数据处理", "100")
    If numRows <> "" Then
        ' 处理指定行数
    End If
End Sub
```
- **验证点**：输入验证、取消处理

### 测试用例 V-061：MsgBox输出
- **用户需求**：`处理完成后告诉用户结果`
- **预期VBA**：
```vba
Sub ShowResult()
    Dim count As Long
    count = Application.CountA(Range("A:A")) - 1
    MsgBox "处理完成！" & vbCrLf & "共处理 " & count & " 行数据。", vbInformation
End Sub
```
- **验证点**：信息清晰、格式正确

---

## 八、安全相关

### 测试用例 V-070：操作前备份
- **用户需求**：`批量修改前先备份当前工作表`
- **预期VBA**：
```vba
Sub BackupBeforeModify()
    ActiveSheet.Copy After:=ActiveSheet
    ActiveSheet.Name = Format(Now, "yyyymmdd_hhmmss") & "_备份"
    ' ... 执行修改操作 ...
End Sub
```
- **验证点**：备份正确创建

### 测试用例 V-071：撤销保护
- **用户需求**：`保护工作表但允许用户编辑特定区域`
- **预期VBA**：
```vba
Sub ProtectWithExceptions()
    ActiveSheet.Unprotect Password:="123"
    ActiveSheet.Cells.Locked = True
    ActiveSheet.Range("A1:D10").Locked = False
    ActiveSheet.Protect Password:="123"
End Sub
```
- **验证点**：指定区域可编辑、其他区域被保护

---

## 注意事项

1. **VBA执行需要用户确认**：DeepExcel应在执行VBA前弹出确认对话框
2. **代码安全性**：检查生成的VBA是否包含危险操作（如删除文件、格式化硬盘）
3. **错误恢复**：执行失败时应能回滚到操作前状态
4. **COM线程**：VBA必须在Excel主线程执行
