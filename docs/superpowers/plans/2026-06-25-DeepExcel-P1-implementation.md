# DeepExcel P1 MVP闭环实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 实现单任务Agent闭环，覆盖公式生成、VBA执行、数据清洗三大核心场景

**Architecture:** 
- 在P0架构基础上扩展Agent能力
- 增加流式输出、多轮对话、错误自动重试
- 支持Python脚本执行

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
- API Key存储在本地配置，不上传
- 所有批量操作前必须快照

---

## Task 1: 流式输出支持

**Files:**
- Modify: `src/DeepExcel.UI/src/App.tsx` - 添加流式渲染
- Create: `src/DeepExcel.UI/src/components/StreamingMessage.tsx`
- Modify: `src/DeepExcel.AddIn/Bridge/MessageBridge.cs` - 添加流式回调

**Interfaces:**
- Consumes: C#发送的流式token
- Produces: 实时更新的聊天消息

- [ ] **Step 1: 创建流式消息组件**

```tsx
// StreamingMessage.tsx
interface StreamingMessageProps {
  content: string
  isStreaming: boolean
}

export function StreamingMessage({ content, isStreaming }: StreamingMessageProps) {
  return (
    <div className={`message assistant ${isStreaming ? 'streaming' : ''}`}>
      {content}
      {isStreaming && <span className="cursor">▊</span>}
    </div>
  )
}
```

- [ ] **Step 2: 更新App.tsx支持流式**

```tsx
// App.tsx 修改
const [streamingContent, setStreamingContent] = useState('')

useEffect(() => {
  // 监听流式消息
  window.chrome.webview.addEventListener('message', (e) => {
    const msg = JSON.parse(e.data)
    if (msg.type === 'stream_delta') {
      setStreamingContent(prev => prev + msg.payload.delta)
    } else if (msg.type === 'stream_end') {
      setMessages(prev => [...prev, { role: 'assistant', content: streamingContent }])
      setStreamingContent('')
    }
  })
}, [streamingContent])
```

- [ ] **Step 3: 提交**

---

## Task 2: 多轮对话上下文

**Files:**
- Create: `src/DeepExcel.AddIn/Agent/ConversationContext.cs`
- Modify: `src/DeepExcel.AddIn/Agent/Orchestrator.cs` - 添加上下文管理

**Interfaces:**
- Consumes: 用户历史消息
- Produces: 带上下文的模型请求

- [ ] **Step 1: 创建对话上下文**

```csharp
public class ConversationContext
{
    public List<ChatMessage> Messages { get; } = new();
    public ExcelContext LastContext { get; set; }
    
    public void AddUserMessage(string content)
    {
        Messages.Add(new ChatMessage { Role = "user", Content = content });
    }
    
    public void AddAssistantMessage(string content)
    {
        Messages.Add(new ChatMessage { Role = "assistant", Content = content });
    }
}

public class ChatMessage
{
    public string Role { get; set; }
    public string Content { get; set; }
}
```

- [ ] **Step 2: 更新Orchestrator支持上下文**

```csharp
public async Task<OrchestrationResult> ProcessAsync(
    string userMessage, 
    ConversationContext context)
{
    context.AddUserMessage(userMessage);
    
    var request = new ModelRequest
    {
        UserMessage = userMessage,
        Context = context.LastContext,
        ConversationHistory = context.Messages.TakeLast(10).ToList(), // 最近10轮
        AvailableTools = GetToolDefinitions()
    };
    
    var response = await _modelAdapter.SendMessageAsync(request);
    context.AddAssistantMessage(response.Content);
    
    return new OrchestrationResult { Content = response.Content };
}
```

- [ ] **Step 3: 提交**

---

## Task 3: 公式生成与插入

**Files:**
- Create: `src/DeepExcel.AddIn/Tools/FormulaTool.cs`
- Create: `src/DeepExcel.AddIn/Tools/FormulaGenerator.cs`

**Interfaces:**
- Consumes: 自然语言公式描述 + 选区上下文
- Produces: 生成的Excel公式并插入

- [ ] **Step 1: 创建FormulaGenerator**

```csharp
public class FormulaTool
{
    private readonly RangeAnalyzer _rangeAnalyzer;
    
    public FormulaTool(RangeAnalyzer rangeAnalyzer)
    {
        _rangeAnalyzer = rangeAnalyzer;
    }
    
    public FormulaResult GenerateAndInsert(string description, Range targetRange)
    {
        var rangeInfo = _rangeAnalyzer.Analyze(targetRange);
        
        // 生成公式提示词
        var prompt = BuildFormulaPrompt(description, rangeInfo);
        
        // 调用模型生成公式
        var formula = CallModelForFormula(prompt);
        
        // 插入公式
        targetRange.Formula = formula;
        
        return new FormulaResult { Success = true, Formula = formula };
    }
    
    private string BuildFormulaPrompt(string description, RangeInfo context)
    {
        return $@"用户需求: {description}
选区地址: {context.Address}
列数: {context.ColumnCount}, 行数: {context.RowCount}
请生成一个Excel公式，直接返回公式内容，不要其他解释。";
    }
}
```

- [ ] **Step 2: 提交**

---

## Task 4: Python脚本执行

**Files:**
- Create: `src/DeepExcel.AddIn/Executor/PythonExecutor.cs`
- Create: `src/DeepExcel.AddIn/Executor/PythonEnvironment.cs`

**Interfaces:**
- Consumes: Python脚本代码
- Produces: 执行结果或错误

- [ ] **Step 1: 创建PythonExecutor**

```csharp
public class PythonExecutor
{
    private readonly string _pythonPath;
    
    public PythonExecutor()
    {
        _pythonPath = FindPythonPath();
    }
    
    public ExecutionResult Execute(string script, Dictionary<string, object> context)
    {
        if (string.IsNullOrEmpty(_pythonPath))
        {
            return new ExecutionResult 
            { 
                Success = false, 
                Error = "Python未安装或不在PATH中" 
            };
        }
        
        // 生成完整脚本（包含上下文注入）
        var fullScript = GenerateContextualScript(script, context);
        
        // 写入临时文件
        var tempFile = Path.GetTempFileName() + ".py";
        File.WriteAllText(tempFile, fullScript);
        
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = _pythonPath,
                Arguments = $"\"{tempFile}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            };
            
            using var process = Process.Start(psi);
            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();
            process.WaitForExit();
            
            return new ExecutionResult 
            { 
                Success = process.ExitCode == 0, 
                Output = output,
                Error = error
            };
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
    
    private string FindPythonPath()
    {
        // 查找Python可执行文件路径
        var paths = new[] { "python", "python3", @"C:\Python39\python.exe" };
        foreach (var p in paths)
        {
            try
            {
                var psi = new ProcessStartInfo { FileName = p, Arguments = "--version" };
                using var proc = Process.Start(psi);
                if (proc != null) return p;
            }
            catch { }
        }
        return null;
    }
}
```

- [ ] **Step 2: 提交**

---

## Task 5: 数据清洗工具集

**Files:**
- Create: `src/DeepExcel.AddIn/Tools/DataCleaner.cs`
- Create: `src/DeepExcel.AddIn/Tools/CleaningOperations.cs`

**Interfaces:**
- Consumes: 选区数据 + 清洗规则
- Produces: 清洗后的数据和操作日志

- [ ] **Step 1: 创建DataCleaner**

```csharp
public class DataCleaner
{
    public CleaningResult CleanData(
        Range range, 
        CleaningOptions options)
    {
        var logs = new List<string>();
        
        if (options.UnifyDateFormat)
        {
            var count = UnifyDateFormats(range);
            logs.Add($"统一日期格式: {count}个单元格");
        }
        
        if (options.RemoveDuplicates)
        {
            var count = RemoveDuplicates(range);
            logs.Add($"删除重复项: {count}行");
        }
        
        if (options.HighlightMissing)
        {
            var count = HighlightMissingValues(range);
            logs.Add($"标记缺失值: {count}个单元格");
        }
        
        return new CleaningResult 
        { 
            Success = true, 
            OperationLogs = logs 
        };
    }
    
    private int UnifyDateFormats(Range range)
    {
        int count = 0;
        foreach (Range cell in range)
        {
            if (cell.Value is DateTime dt)
            {
                cell.NumberFormat = "yyyy-mm-dd";
                count++;
            }
        }
        return count;
    }
    
    private int HighlightMissingValues(Range range)
    {
        int count = 0;
        foreach (Range cell in range)
        {
            if (string.IsNullOrEmpty(cell.Value?.ToString()))
            {
                cell.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                count++;
            }
        }
        return count;
    }
}

public class CleaningOptions
{
    public bool UnifyDateFormat { get; set; }
    public bool RemoveDuplicates { get; set; }
    public bool HighlightMissing { get; set; }
}
```

- [ ] **Step 2: 提交**

---

## Task 6: 图表与透视表创建

**Files:**
- Create: `src/DeepExcel.AddIn/Tools/ChartTool.cs`
- Create: `src/DeepExcel.AddIn/Tools/PivotTableTool.cs`

**Interfaces:**
- Consumes: 数据源范围 + 图表/透视表配置
- Produces: 创建的图表或透视表

- [ ] **Step 1: 创建ChartTool**

```csharp
public class ChartTool
{
    public ChartResult CreateChart(
        Range dataRange, 
        ChartConfig config)
    {
        try
        {
            var chart = dataRange.Parent.Shapes.AddChart(
                XlChartType.xlColumnClustered,
                dataRange.Left, dataRange.Top, 
                dataRange.Width, dataRange.Height
            ).Chart;
            
            chart.SetSourceData(dataRange);
            chart.HasTitle = true;
            chart.ChartTitle.Text = config.Title;
            
            return new ChartResult { Success = true, Chart = chart };
        }
        catch (Exception ex)
        {
            return new ChartResult { Success = false, Error = ex.Message };
        }
    }
}

public class ChartConfig
{
    public string Title { get; set; }
    public XlChartType ChartType { get; set; }
    public string SeriesName { get; set; }
}
```

- [ ] **Step 2: 创建PivotTableTool**

```csharp
public class PivotTableTool
{
    public PivotTableResult CreatePivotTable(
        Range dataRange,
        Range destination,
        PivotTableConfig config)
    {
        try
        {
            var pivotTable = destination.Worksheet.PivotTables().Add(
                PivotTable:=destination,
                TableData:=dataRange,
                TableName:=config.Name);
            
            // 设置行字段
            pivotTable.PivotFields(config.RowField).Orientation = 
                XlPivotFieldOrientation.xlRowField;
            
            // 设置值字段
            pivotTable.PivotFields(config.ValueField).Orientation = 
                XlPivotFieldOrientation.xlDataField;
            pivotTable.PivotFields(config.ValueField).Function = 
                XlConsolidationFunction.xlSum;
            
            return new PivotTableResult { Success = true, PivotTable = pivotTable };
        }
        catch (Exception ex)
        {
            return new PivotTableResult { Success = false, Error = ex.Message };
        }
    }
}
```

- [ ] **Step 3: 提交**

---

## Task 7: 错误自动重试机制

**Files:**
- Modify: `src/DeepExcel.AddIn/Agent/Orchestrator.cs` - 添加重试逻辑
- Create: `src/DeepExcel.AddIn/Agent/RetryPolicy.cs`

**Interfaces:**
- Consumes: 失败的工具调用结果
- Produces: 重试后的结果

- [ ] **Step 1: 创建RetryPolicy**

```csharp
public class RetryPolicy
{
    private readonly int _maxRetries = 3;
    
    public async Task<T> ExecuteWithRetryAsync<T>(
        Func<Task<T>> operation,
        Func<T, bool> isSuccess)
    {
        T result = default;
        for (int i = 0; i < _maxRetries; i++)
        {
            result = await operation();
            if (isSuccess(result)) return result;
            
            await Task.Delay(1000 * (i + 1)); // 指数退避
        }
        return result;
    }
}
```

- [ ] **Step 2: 更新Orchestrator使用重试**

```csharp
public async Task<OrchestrationResult> ExecuteToolCallWithRetryAsync(ToolCall toolCall)
{
    var policy = new RetryPolicy();
    
    return await policy.ExecuteWithRetryAsync(
        async () => await ExecuteToolCallAsync(toolCall),
        result => result.Success);
}
```

- [ ] **Step 3: 提交**

---

## P1验证清单

- [ ] 公式生成：输入"计算A1到A10的总和" → A1显示=SUM(A1:A10)
- [ ] VBA执行：多步VBA任务自动重试
- [ ] 数据清洗：统一日期格式、删除重复、标红缺失
- [ ] Python执行：生成并运行Python脚本
- [ ] 图表创建：根据数据生成柱状图
- [ ] 透视表：创建简单的透视表
- [ ] 多轮对话：上下文保持正确
- [ ] 流式输出：实时显示思考过程
