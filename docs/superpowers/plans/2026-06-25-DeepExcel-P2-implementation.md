# DeepExcel P2 能力扩展实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 覆盖主流场景，支持VBA+Python+图表/透视表+多步工作流

**Architecture:** 
- 在P1基础上扩展工具集
- 实现复杂多步骤工作流编排
- 添加模板和素材库支持
- 完善提示词工程

**Tech Stack:** 
- C# / .NET Framework 4.7.2+ / VSTO
- Microsoft.Office.Interop.Excel
- React + TypeScript
- Claude/DeepSeek API

---

## Global Constraints

- Excel COM对象操作必须在UI线程(STA)执行
- COM引用释放必须使用Marshal.ReleaseComObject
- 原始.xlsx文件永不上传，只传任务相关的结构数据
- 所有批量操作前必须快照
- 工作流执行中间状态需要持久化

---

## Task 1: 多步骤工作流编排器

**Files:**
- Create: `src/DeepExcel.AddIn/Agent/WorkflowOrchestrator.cs`
- Create: `src/DeepExcel.AddIn/Agent/WorkflowStep.cs`
- Create: `src/DeepExcel.AddIn/Agent/WorkflowState.cs`

**Interfaces:**
- Consumes: 复合任务描述
- Produces: 多步骤执行结果

- [ ] **Step 1: 创建工作流模型**

```csharp
public class WorkflowStep
{
    public int Order { get; set; }
    public string Description { get; set; }
    public string ToolName { get; set; }
    public Dictionary<string, object> Arguments { get; set; }
    public StepStatus Status { get; set; }
    public string Result { get; set; }
    public string Error { get; set; }
}

public enum StepStatus
{
    Pending,
    Running,
    Completed,
    Failed,
    Skipped
}

public class WorkflowState
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public List<WorkflowStep> Steps { get; set; } = new();
    public int CurrentStepIndex { get; set; }
    public WorkflowStatus Status { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

public enum WorkflowStatus
{
    Running,
    Completed,
    Failed,
    Cancelled
}
```

- [ ] **Step 2: 创建WorkflowOrchestrator**

```csharp
public class WorkflowOrchestrator
{
    private readonly Orchestrator _baseOrchestrator;
    
    public async Task<WorkflowResult> ExecuteWorkflowAsync(
        string goal,
        ExcelContext context)
    {
        var state = new WorkflowState();
        
        // 1. 分解任务为步骤
        var steps = await DecomposeGoalAsync(goal, context);
        state.Steps = steps;
        
        // 2. 依次执行每步
        foreach (var step in state.Steps)
        {
            step.Status = StepStatus.Running;
            var result = await ExecuteStepAsync(step, context);
            
            if (result.Success)
            {
                step.Status = StepStatus.Completed;
                step.Result = result.Output;
            }
            else
            {
                step.Status = StepStatus.Failed;
                step.Error = result.Error;
                state.Status = WorkflowStatus.Failed;
                break;
            }
        }
        
        return new WorkflowResult { State = state };
    }
    
    private async Task<List<WorkflowStep>> DecomposeGoalAsync(
        string goal, 
        ExcelContext context)
    {
        var prompt = $@"将以下任务分解为具体步骤:
任务: {goal}

当前Excel状态:
工作表: {context.Workbook?.Name}
选区: {context.CurrentSelection?.Address}

返回JSON格式的步骤数组，每步包含description和toolName。";
        
        var response = await _modelAdapter.SendMessageAsync(prompt);
        return ParseStepsFromResponse(response.Content);
    }
}
```

- [ ] **Step 3: 提交**

---

## Task 2: 条件格式与数据验证工具

**Files:**
- Create: `src/DeepExcel.AddIn/Tools/ConditionalFormatTool.cs`
- Create: `src/DeepExcel.AddIn/Tools/DataValidationTool.cs`

**Interfaces:**
- Consumes: 范围 + 格式/验证规则
- Produ产: 应用后的格式/验证

- [ ] **Step 1: 创建ConditionalFormatTool**

```csharp
public class ConditionalFormatTool
{
    public FormatResult ApplyConditionalFormat(
        Range range,
        ConditionalFormatRule rule)
    {
        try
        {
            var format = range.FormatConditions.Add(
                rule.Type,
                rule.Operator,
                rule.Formula1,
                rule.Formula2);
            
            format.Interior.Color = rule.FillColor;
            format.Font.Color = rule.FontColor;
            format.Font.Bold = rule.IsBold;
            
            return new FormatResult { Success = true };
        }
        catch (Exception ex)
        {
            return new FormatResult { Success = false, Error = ex.Message };
        }
    }
}

public class ConditionalFormatRule
{
    public XlFormatConditionType Type { get; set; }
    public XlFormatConditionOperator Operator { get; set; }
    public string Formula1 { get; set; }
    public string Formula2 { get; set; }
    public Color FillColor { get; set; }
    public Color FontColor { get; set; }
    public bool IsBold { get; set; }
}
```

- [ ] **Step 2: 创建DataValidationTool**

```csharp
public class DataValidationTool
{
    public ValidationResult ApplyValidation(
        Range range,
        DataValidationRule rule)
    {
        try
        {
            var validation = range.Validation.Add(
                rule.Type,
                XlAlertStyle.xlAlertStyleStop,
                rule.Operator,
                rule.Formula1,
                rule.Formula2);
            
            validation.InputTitle = rule.InputTitle;
            validation.InputMessage = rule.InputMessage;
            validation.ErrorTitle = rule.ErrorTitle;
            validation.ErrorMessage = rule.ErrorMessage;
            
            return new ValidationResult { Success = true };
        }
        catch (Exception ex)
        {
            return new ValidationResult { Success = false, Error = ex.Message };
        }
    }
}

public class DataValidationRule
{
    public XlDVType Type { get; set; }
    public XlDataValidationOperator Operator { get; set; }
    public string Formula1 { get; set; }
    public string Formula2 { get; set; }
    public string InputTitle { get; set; }
    public string InputMessage { get; set; }
    public string ErrorTitle { get; set; }
    public string ErrorMessage { get; set; }
}
```

- [ ] **Step 3: 提交**

---

## Task 3: 命名区域与表管理

**Files:**
- Create: `src/DeepExcel.AddIn/Tools/NamedRangeTool.cs`
- Create: `src/DeepExcel.AddIn/Tools/TableTool.cs`

**Interfaces:**
- Consumes: 范围 + 名称/表配置
- 产生活名区域或Excel表

- [ ] **Step 1: 创建NamedRangeTool**

```csharp
public class NamedRangeTool
{
    public NamedRangeResult CreateNamedRange(
        Range range,
        string name)
    {
        try
        {
            var wb = range.Worksheet.Parent as Workbook;
            wb.Names.Add(name, range);
            
            return new NamedRangeResult { Success = true, Name = name };
        }
        catch (Exception ex)
        {
            return new NamedRangeResult { Success = false, Error = ex.Message };
        }
    }
    
    public NamedRangeResult DeleteNamedRange(string name)
    {
        try
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            wb.Names.Item[name].Delete();
            return new NamedRangeResult { Success = true };
        }
        catch (Exception ex)
        {
            return new NamedRangeResult { Success = false, Error = ex.Message };
        }
    }
}
```

- [ ] **Step 2: 创建TableTool**

```csharp
public class TableTool
{
    public TableResult CreateTable(
        Range range,
        TableConfig config)
    {
        try
        {
            var listObject = range.Worksheet.ListObjects.Add(
                XlListObjectSourceType.xlSrcRange,
                range,
                config.HasHeaders ? XlYesNoGuess.xlYes : XlYesNoGuess.xlNo);
            
            listObject.Name = config.TableName;
            listObject.TableStyle = config.StyleName;
            
            return new TableResult { Success = true, Table = listObject };
        }
        catch (Exception ex)
        {
            return new TableResult { Success = false, Error = ex.Message };
        }
    }
}

public class TableConfig
{
    public string TableName { get; set; }
    public bool HasHeaders { get; set; } = true;
    public string StyleName { get; set; } = "TableStyleMedium2";
}
```

- [ ] **Step 3: 提交**

---

## Task 4: 工作表与工作簿操作

**Files:**
- Create: `src/DeepExcel.AddIn/Tools/WorksheetTool.cs`
- Create: `src/DeepExcel.AddIn/Tools/WorkbookTool.cs`

**Interfaces:**
- 产生活表/工作簿操作

- [ ] **Step 1: 创建WorksheetTool**

```csharp
public class WorksheetTool
{
    public WorksheetResult CreateWorksheet(string name, string afterSheet = null)
    {
        try
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var newSheet = wb.Worksheets.Add(
                After: afterSheet != null 
                    ? wb.Worksheets[afterSheet] 
                    : wb.Worksheets[wb.Worksheets.Count]);
            
            if (name != null) newSheet.Name = name;
            
            return new WorksheetResult { Success = true, Worksheet = newSheet };
        }
        catch (Exception ex)
        {
            return new WorksheetResult { Success = false, Error = ex.Message };
        }
    }
    
    public WorksheetResult DuplicateWorksheet(string sourceName, string newName)
    {
        try
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var source = wb.Worksheets[sourceName];
            var duplicate = source.Copy(After: source);
            
            if (newName != null) (duplicate as Worksheet).Name = newName;
            
            return new WorksheetResult { Success = true, Worksheet = duplicate as Worksheet };
        }
        catch (Exception ex)
        {
            return new WorksheetResult { Success = false, Error = ex.Message };
        }
    }
}
```

- [ ] **Step 2: 提交**

---

## Task 5: 模板系统

**Files:**
- Create: `src/DeepExcel.AddIn/Templates/TemplateManager.cs`
- Create: `src/DeepExcel.AddIn/Templates/TemplateRegistry.cs`
- Create: `src/DeepExcel.AddIn/Templates/FinancialTemplate.cs`

**Interfaces:**
- 提供预定义模板（财务模型、报表等）
- 支持用户自定义模板

- [ ] **Step 1: 创建模板管理器**

```csharp
public class TemplateManager
{
    private readonly string _templateFolder;
    private readonly Dictionary<string, TemplateDefinition> _templates;
    
    public TemplateManager()
    {
        _templateFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DeepExcel", "Templates");
        Directory.CreateDirectory(_templateFolder);
        _templates = LoadTemplates();
    }
    
    public List<string> GetAvailableTemplates()
    {
        return _templates.Keys.ToList();
    }
    
    public TemplateResult ApplyTemplate(string templateName, ExcelContext context)
    {
        if (!_templates.ContainsKey(templateName))
            return new TemplateResult { Success = false, Error = "模板不存在" };
        
        var template = _templates[templateName];
        return template.Apply(context);
    }
    
    private Dictionary<string, TemplateDefinition> LoadTemplates()
    {
        return new Dictionary<string, TemplateDefinition>
        {
            ["财务损益表"] = new FinancialIncomeStatementTemplate(),
            ["月度预算"] = new MonthlyBudgetTemplate(),
            ["数据看板"] = new DataDashboardTemplate()
        };
    }
}

public abstract class TemplateDefinition
{
    public string Name { get; set; }
    public string Description { get; set; }
    
    public abstract TemplateResult Apply(ExcelContext context);
}
```

- [ ] **Step 2: 创建财务模板示例**

```csharp
public class FinancialIncomeStatementTemplate : TemplateDefinition
{
    public FinancialIncomeStatementTemplate()
    {
        Name = "财务损益表";
        Description = "创建标准月度/年度损益表";
    }
    
    public override TemplateResult Apply(ExcelContext context)
    {
        // 创建损益表结构
        var ws = context.Workbook.Worksheets.FirstOrDefault();
        ws.Name = "损益表";
        
        // 写入表头
        ws.Range["A1"].Value = "损益表";
        ws.Range["A3"].Value = "项目";
        ws.Range["B3"].Value = "本月";
        ws.Range["C3"].Value = "本年累计";
        
        // 写入明细行
        var rows = new[] { "营业收入", "营业成本", "营业利润", "净利润" };
        for (int i = 0; i < rows.Length; i++)
        {
            ws.Range[$"A{4+i}"].Value = rows[i];
            ws.Range[$"B{4+i}"].Formula = $"=SUM(...)";
        }
        
        return new TemplateResult { Success = true };
    }
}
```

- [ ] **Step 3: 提交**

---

## Task 6: 提示词工程与工具定义

**Files:**
- Create: `src/DeepExcel.AddIn/Prompts/PromptTemplates.cs`
- Create: `src/DeepExcel.AddIn/Prompts/ToolDefinitions.cs`
- Create: `src/DeepExcel.AddIn/Prompts/ContextBuilders.cs`

**Interfaces:**
- 产生活性化的提示词和工具定义
- 支持提示词热更新

- [ ] **Step 1: 创建工具定义**

```csharp
public static class ToolDefinitions
{
    public static List<ToolDefinition> GetAllTools()
    {
        return new List<ToolDefinition>
        {
            new ToolDefinition
            {
                Name = "read_workbook",
                Description = "读取当前工作簿的结构概览",
                Parameters = new { }
            },
            new ToolDefinition
            {
                Name = "read_range",
                Description = "读取指定范围的单元格数据",
                Parameters = new { address = "范围地址，如A1:B10" }
            },
            new ToolDefinition
            {
                Name = "write_formula",
                Description = "向指定单元格写入公式",
                Parameters = new { 
                    address = "目标单元格地址",
                    formula = "Excel公式，如=SUM(A1:A10)"
                }
            },
            new ToolDefinition
            {
                Name = "execute_vba",
                Description = "执行VBA代码完成复杂任务",
                Parameters = new { code = "完整的VBA代码" }
            },
            new ToolDefinition
            {
                Name = "create_chart",
                Description = "根据数据创建图表",
                Parameters = new { 
                    dataRange = "数据范围",
                    chartType = "图表类型",
                    title = "图表标题"
                }
            },
            new ToolDefinition
            {
                Name = "clean_data",
                Description = "清洗数据（统一格式、删除重复、标红缺失）",
                Parameters = new { 
                    range = "数据范围",
                    operations = new[] { "unify_date", "remove_duplicates", "highlight_missing" }
                }
            },
            new ToolDefinition
            {
                Name = "snapshot",
                Description = "创建当前状态的快照（操作前必须）",
                Parameters = new { }
            },
            new ToolDefinition
            {
                Name = "rollback",
                Description = "回滚到指定快照",
                Parameters = new { snapshotId = "快照ID" }
            }
        };
    }
}
```

- [ ] **Step 2: 创建上下文构建器**

```csharp
public static class ContextBuilders
{
    public static string BuildReadableContext(ExcelContext context)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("【工作簿信息】");
        sb.AppendLine($"名称: {context.Workbook?.Name}");
        sb.AppendLine($"工作表数量: {context.Workbook?.Worksheets?.Length}");
        
        if (context.Workbook?.Worksheets != null)
        {
            sb.AppendLine("工作表列表:");
            foreach (var sheet in context.Workbook.Worksheets)
            {
                sb.AppendLine($"  - {sheet.Name} (使用范围: {sheet.UsedRangeAddress})");
            }
        }
        
        if (context.Workbook?.NamedRanges != null)
        {
            sb.AppendLine("命名区域:");
            foreach (var name in context.Workbook.NamedRanges)
            {
                sb.AppendLine($"  - {name.Name} -> {name.RefersTo}");
            }
        }
        
        if (context.CurrentSelection != null)
        {
            sb.AppendLine("【当前选区】");
            sb.AppendLine($"地址: {context.CurrentSelection.Address}");
            sb.AppendLine($"行数: {context.CurrentSelection.RowCount}, 列数: {context.CurrentSelection.ColumnCount}");
            
            if (context.CurrentSelection.Values != null)
            {
                sb.AppendLine("数据预览:");
                sb.AppendLine(FormatValuesPreview(context.CurrentSelection.Values));
            }
        }
        
        return sb.ToString();
    }
}
```

- [ ] **Step 3: 提交**

---

## Task 7: 配置管理系统

**Files:**
- Create: `src/DeepExcel.AddIn/Config/AppSettings.cs`
- Create: `src/DeepExcel.AddIn/Config/SettingsManager.cs`
- Create: `src/DeepExcel.AddIn/Config/ModelConfig.cs`

**Interfaces:**
- 管理用户配置
- 支持模型切换

- [ ] **Step 1: 创建设置模型**

```csharp
public class AppSettings
{
    public ModelSettings Model { get; set; } = new();
    public UIsettings UI { get; set; } = new();
    public SafetySettings Safety { get; set; } = new();
}

public class ModelSettings
{
    public string ActiveModel { get; set; } = "claude";
    public string ApiKey { get; set; }
    public string ApiEndpoint { get; set; } = "https://api.anthropic.com";
    public int MaxTokens { get; set; } = 4096;
    public double Temperature { get; set; } = 0.7;
}

public class SafetySettings
{
    public bool AutoSnapshot { get; set; } = true;
    public int MaxRetries { get; set; } = 3;
    public bool ConfirmBeforeExecute { get; set; } = false;
}
```

- [ ] **Step 2: 创建设置管理器**

```csharp
public class SettingsManager
{
    private readonly string _settingsPath;
    private AppSettings _settings;
    
    public SettingsManager()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var folder = Path.Combine(appData, "DeepExcel");
        Directory.CreateDirectory(folder);
        _settingsPath = Path.Combine(folder, "settings.json");
        Load();
    }
    
    public AppSettings Settings => _settings;
    
    public void Load()
    {
        if (File.Exists(_settingsPath))
        {
            var json = File.ReadAllText(_settingsPath);
            _settings = JsonSerializer.Deserialize<AppSettings>(json);
        }
        else
        {
            _settings = new AppSettings();
        }
    }
    
    public void Save()
    {
        var json = JsonSerializer.Serialize(_settings, new JsonSerializerOptions 
        { 
            WriteIndented = true 
        });
        File.WriteAllText(_settingsPath, json);
    }
    
    public void SetActiveModel(string modelName)
    {
        _settings.Model.ActiveModel = modelName;
        Save();
    }
}
```

- [ ] **Step 3: 提交**

---

## P2验证清单

- [ ] 多步骤工作流：输入"创建损益表并生成图表" → 自动分解并执行
- [ ] 条件格式：高亮显示大于100的单元格
- [ ] 数据验证：设置下拉列表验证
- [ ] 命名区域：创建"销售额"命名区域
- [ ] 表管理：将数据转换为Excel表
- [ ] 工作表操作：复制、重命名、移动工作表
- [ ] 模板应用：应用财务损益表模板
- [ ] 配置切换：Claude ↔ DeepSeek 切换
- [ ] 提示词热更新：更新工具定义无需重发安装包
