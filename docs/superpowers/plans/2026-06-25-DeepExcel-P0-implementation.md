# DeepExcel P0 架构打通实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 跑通"感知→Agent编排→VBA执行"管线，验证VSTO+WebView2+VBA三件套可协同工作

**Architecture:** 
- VSTO COM加载项嵌入Excel，持有Application对象
- WebView2承载React聊天UI，通过PostMessage与C#桥接
- C#执行引擎生成VBA代码，注入Excel VBA模块并执行
- 感知层导出Excel结构，Agent基于结构数据调用工具

**Tech Stack:** 
- C# / .NET Framework 4.7.2+ / VSTO
- Microsoft.Office.Interop.Excel
- WebView2 (Evergreen)
- React + TypeScript + Vite
- Claude API (本地API Key直连)

---

## Global Constraints

- Excel COM对象操作必须在UI线程(STA)执行
- COM引用释放必须使用Marshal.ReleaseComObject
- 原始.xlsx文件永不上传，只传任务相关的结构数据
- API Key存储在本地配置，不上传

---

## 阶段一：项目骨架搭建

### Task 1: 创建VSTO加载项项目结构

**Files:**
- Create: `src/DeepExcel.AddIn/DeepExcel.AddIn.csproj`
- Create: `src/DeepExcel.AddIn/Properties/AssemblyInfo.cs`
- Create: `src/DeepExcel.AddIn/ThisAddIn.cs`
- Create: `src/DeepExcel.AddIn/Ribbon/DeepExcelRibbon.cs`
- Create: `src/DeepExcel.AddIn/ThisAddIn.Designer.xml`

**Interfaces:**
- Produces: `ThisAddIn` 加载项入口，`DeepExcelRibbon` 功能区

- [ ] **Step 1: 创建.csproj项目文件**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <OutputType>Library</OutputType>
    <UseWindowsForms>true</UseWindowsForms>
    <Enable COMInterop>true</Enable>
    <RegisterForComInterop>true</RegisterForComInterop>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.Office.Interop.Excel" Version="15.0.4795.1001" />
    <PackageReference Include="Microsoft.Web.WebView2" Version="1.0.2420.47" />
  </ItemGroup>
</Project>
```

- [ ] **Step 2: 创建AssemblyInfo.cs**

```csharp
using System.Reflection;
using System.Runtime.InteropServices;
[assembly: AssemblyTitle("DeepExcel.AddIn")]
[assembly: AssemblyVersion("0.1.0.0")]
```

- [ ] **Step 3: 创建ThisAddIn.cs主入口**

```csharp
using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Web.WebView2;

public partial class ThisAddIn
{
    private ThisAddIn()
    {
    }
    
    protected override void Startup()
    {
        Globals.ThisAddIn.Application.WorkbookOpen += OnWorkbookOpen;
        CreateTaskPane();
    }

    private void OnWorkbookOpen(Workbook wb)
    {
    }

    private void CreateTaskPane()
    {
        var actionsPane = this.CustomTaskPanes.Add(
            new UserControl(), "DeepExcel");
        actionsPane.Visible = true;
    }

    protected override void Shutdown()
    {
    }
}
```

- [ ] **Step 4: 创建Ribbon功能区**

```csharp
using Microsoft.Office.Tools.Ribbon;
public class DeepExcelRibbon : RibbonBase
{
    public DeepExcelRibbon()
        : base(Globals.Factory.GetRibbonFactory())
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        var tab = Factory.CreateRibbonTab();
        tab.Label = "DeepExcel";
        
        var group = Factory.CreateRibbonGroup();
        group.Label = "工具";
        
        var button = Factory.CreateRibbonButton();
        button.Label = "打开面板";
        button.Click += (s, e) => ShowPanel();
        
        group.Items.Add(button);
        tab.Controls.Add(group);
        this.Tabs.Add(tab);
    }

    private void ShowPanel()
    {
        var pane = Globals.ThisAddIn.CustomTaskPanes[0];
        pane.Visible = true;
    }
}
```

- [ ] **Step 5: 提交**

```bash
git add src/DeepExcel.AddIn/
git commit -m "feat: 创建VSTO加载项骨架"
```

---

### Task 2: 创建WebView2 UI项目

**Files:**
- Create: `src/DeepExcel.UI/index.html`
- Create: `src/DeepExcel.UI/src/App.tsx`
- Create: `src/DeepExcel.UI/src/main.tsx`
- Create: `src/DeepExcel.UI/package.json`
- Create: `src/DeepExcel.UI/vite.config.ts`
- Create: `src/DeepExcel.UI/tsconfig.json`

**Interfaces:**
- Consumes: C# WebView2宿主提供的postMessage接口
- Produces: React聊天UI组件

- [ ] **Step 1: 创建package.json**

```json
{
  "name": "deepexcel-ui",
  "version": "0.1.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "tsc && vite build",
    "preview": "vite preview"
  },
  "dependencies": {
    "react": "^18.2.0",
    "react-dom": "^18.2.0"
  },
  "devDependencies": {
    "@types/react": "^18.2.0",
    "@types/react-dom": "^18.2.0",
    "@vitejs/plugin-react": "^4.0.0",
    "typescript": "^5.0.0",
    "vite": "^5.0.0"
  }
}
```

- [ ] **Step 2: 创建index.html**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <title>DeepExcel</title>
</head>
<body>
  <div id="root"></div>
  <script type="module" src="/src/main.tsx"></script>
</body>
</html>
```

- [ ] **Step 3: 创建App.tsx聊天组件**

```tsx
import { useState, useRef, useEffect } from 'react'

interface Message {
  role: 'user' | 'assistant'
  content: string
}

export default function App() {
  const [messages, setMessages] = useState<Message[]>([])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const messagesEndRef = useRef<HTMLDivElement>(null)

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }

  useEffect(() => {
    scrollToBottom()
  }, [messages])

  const sendMessage = async () => {
    if (!input.trim()) return
    
    const userMessage: Message = { role: 'user', content: input }
    setMessages(prev => [...prev, userMessage])
    setInput('')
    setLoading(true)

    try {
      // 通过WebView桥接发送消息到C#
      const response = await window.chrome.webview.postMessage({
        type: 'user_message',
        payload: { content: input }
      })
    } catch (err) {
      console.error(err)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="chat-container">
      <div className="messages">
        {messages.map((msg, idx) => (
          <div key={idx} className={`message ${msg.role}`}>
            {msg.content}
          </div>
        ))}
        {loading && <div className="message assistant">思考中...</div>}
        <div ref={messagesEndRef} />
      </div>
      <div className="input-area">
        <input
          value={input}
          onChange={e => setInput(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && sendMessage()}
          placeholder="描述你的Excel任务..."
        />
        <button onClick={sendMessage} disabled={loading}>发送</button>
      </div>
    </div>
  )
}
```

- [ ] **Step 4: 创建main.tsx**

```tsx
import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import './styles.css'

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
)
```

- [ ] **Step 5: 创建基础样式styles.css**

```css
* { box-sizing: border-box; margin: 0; padding: 0; }
.chat-container {
  display: flex;
  flex-direction: column;
  height: 100vh;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}
.messages {
  flex: 1;
  overflow-y: auto;
  padding: 16px;
}
.message {
  padding: 8px 12px;
  margin-bottom: 8px;
  border-radius: 8px;
  max-width: 85%;
}
.message.user {
  background: #0078d4;
  color: white;
  align-self: flex-end;
}
.message.assistant {
  background: #f1f1f1;
  align-self: flex-start;
}
.input-area {
  display: flex;
  padding: 16px;
  border-top: 1px solid #ddd;
}
.input-area input {
  flex: 1;
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 6px;
  font-size: 14px;
}
.input-area button {
  margin-left: 8px;
  padding: 8px 16px;
  background: #0078d4;
  color: white;
  border: none;
  border-radius: 6px;
  cursor: pointer;
}
.input-area button:disabled {
  opacity: 0.6;
}
```

- [ ] **Step 6: 提交**

```bash
git add src/DeepExcel.UI/
git commit -m "feat: 创建WebView2 UI项目骨架"
```

---

## 阶段二：桥接层实现

### Task 3: 实现C#与WebView2的双向消息桥接

**Files:**
- Modify: `src/DeepExcel.AddIn/ThisAddIn.cs` - 添加桥接逻辑
- Create: `src/DeepExcel.AddIn/Bridge/MessageBridge.cs`
- Create: `src/DeepExcel.AddIn/Bridge/IExcelActions.cs`

**Interfaces:**
- Consumes: WebView2发送的JSON消息
- Produces: Excel对象操作结果回传UI

- [ ] **Step 1: 创建消息类型定义**

```csharp
// Bridge/MessageBridge.cs
using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace DeepExcel.AddIn.Bridge
{
    public class MessageBridge
    {
        private readonly Application _excelApp;

        public MessageBridge(Application excelApp)
        {
            _excelApp = excelApp;
        }

        public object HandleMessage(string type, string payload)
        {
            return type switch
            {
                "get_selection" => GetSelection(),
                "read_range" => ReadRange(payload),
                "execute_vba" => ExecuteVBA(payload),
                _ => new { error = $"Unknown message type: {type}" }
            };
        }

        private object GetSelection()
        {
            var sel = _excelApp.Selection;
            if (sel is Range range)
            {
                return new {
                    type = "range",
                    address = range.Address(),
                    values = GetRangeValues(range),
                    formulas = GetRangeFormulas(range)
                };
            }
            return new { error = "No range selected" };
        }

        private string[,] GetRangeValues(Range range)
        {
            return range.Value as string[,] ?? new string[1, 1];
        }

        private string[,] GetRangeFormulas(Range range)
        {
            return range.Formula as string[,] ?? new string[1, 1];
        }

        private object ReadRange(string payload)
        {
            var range = _excelApp.Range[payload];
            return new {
                address = range.Address(),
                values = GetRangeValues(range),
                formulas = GetRangeFormulas(range)
            };
        }

        private object ExecuteVBA(string code)
        {
            try
            {
                // 生成临时VBA模块并执行
                var vbaProject = GetOrCreateVBAModule();
                vbaProject.CodeModule.AddFromString(code);
                
                // 执行宏（通过Application.Run）
                _excelApp.Run("Module1.TestMacro");
                
                return new { success = true };
            }
            catch (Exception ex)
            {
                return new { success = false, error = ex.Message };
            }
        }

        private VBProject GetOrCreateVBAModule()
        {
            var wb = _excelApp.ActiveWorkbook;
            return wb.VBProject;
        }
    }
}
```

- [ ] **Step 2: 更新ThisAddIn添加桥接**

```csharp
// ThisAddIn.cs 片段
private MessageBridge _bridge;

protected override void Startup()
{
    _bridge = new MessageBridge(Globals.ThisAddIn.Application);
    
    // WebView2消息处理
    var webview = GetWebView2();
    webview.WebMessageReceived += OnWebMessageReceived;
}

private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
{
    var msg = System.Text.Json.JsonSerializer.Deserialize<Message>(e.WebMessageAsJson);
    var result = _bridge.HandleMessage(msg.type, msg.payload);
    
    // 回传结果
    webview.PostMessageAsJson(System.Text.Json.JsonSerializer.Serialize(result));
}
```

- [ ] **Step 3: 提交**

```bash
git add src/DeepExcel.AddIn/
git commit -m "feat: 实现C#与WebView2消息桥接"
```

---

## 阶段三：感知层实现

### Task 4: 实现Excel结构感知

**Files:**
- Create: `src/DeepExcel.AddIn/Perception/WorkbookAnalyzer.cs`
- Create: `src/DeepExcel.AddIn/Perception/WorksheetAnalyzer.cs`
- Create: `src/DeepExcel.AddIn/Perception/RangeAnalyzer.cs`

**Interfaces:**
- Consumes: Excel对象模型
- Produces: 结构化JSON数据供Agent使用

- [ ] **Step 1: 创建WorkbookAnalyzer**

```csharp
namespace DeepExcel.AddIn.Perception
{
    public class WorkbookAnalyzer
    {
        private readonly Microsoft.Office.Interop.Excel.Application _app;

        public WorkbookAnalyzer(Microsoft.Office.Interop.Excel.Application app)
        {
            _app = app;
        }

        public WorkbookStructure Analyze()
        {
            var wb = _app.ActiveWorkbook;
            if (wb == null) return null;

            return new WorkbookStructure
            {
                Name = wb.Name,
                Worksheets = AnalyzeWorksheets(wb),
                NamedRanges = AnalyzeNamedRanges(wb),
                HasVBAProject = wb.HasVBProject
            };
        }

        private WorksheetInfo[] AnalyzeWorksheets(Workbook wb)
        {
            var sheets = wb.Worksheets;
            var result = new List<WorksheetInfo>();

            foreach (Worksheet sheet in sheets)
            {
                result.Add(new WorksheetInfo
                {
                    Name = sheet.Name,
                    Index = sheet.Index,
                    UsedRangeAddress = sheet.UsedRange.Address()
                });
                Marshal.ReleaseComObject(sheet);
            }
            Marshal.ReleaseComObject(sheets);

            return result.ToArray();
        }

        private NamedRangeInfo[] AnalyzeNamedRanges(Workbook wb)
        {
            var names = wb.Names;
            var result = new List<NamedRangeInfo>();

            foreach (Name name in names)
            {
                result.Add(new NamedRangeInfo
                {
                    Name = name.Name,
                    RefersTo = name.RefersTo
                });
                Marshal.ReleaseComObject(name);
            }
            Marshal.ReleaseComObject(names);

            return result.ToArray();
        }
    }

    public class WorkbookStructure
    {
        public string Name { get; set; }
        public WorksheetInfo[] Worksheets { get; set; }
        public NamedRangeInfo[] NamedRanges { get; set; }
        public bool HasVBAProject { get; set; }
    }

    public class WorksheetInfo
    {
        public string Name { get; set; }
        public int Index { get; set; }
        public string UsedRangeAddress { get; set; }
    }

    public class NamedRangeInfo
    {
        public string Name { get; set; }
        public string RefersTo { get; set; }
    }
}
```

- [ ] **Step 2: 创建RangeAnalyzer**

```csharp
public class RangeAnalyzer
{
    public RangeInfo Analyze(Range range)
    {
        return new RangeInfo
        {
            Address = range.Address(),
            RowCount = range.Rows.Count,
            ColumnCount = range.Columns.Count,
            Values = range.Value as object[,],
            Formulas = range.Formula as string[,],
            FormatConditions = AnalyzeFormatConditions(range),
            MergeCells = AnalyzeMergeCells(range)
        };
    }

    private FormatConditionInfo[] AnalyzeFormatConditions(Range range)
    {
        var formats = range.FormatConditions;
        var result = new List<FormatConditionInfo>();

        foreach (FormatCondition fc in formats)
        {
            result.Add(new FormatConditionInfo
            {
                Type = fc.Type.ToString(),
                Formula1 = fc.Formula1,
                Operator = fc.Operator.ToString()
            });
            Marshal.ReleaseComObject(fc);
        }
        Marshal.ReleaseComObject(formats);

        return result.ToArray();
    }

    private MergeCellInfo[] AnalyzeMergeCells(Range range)
    {
        var mergeAreas = range.MergeAreas;
        var result = new List<MergeCellInfo>();

        foreach (Range area in mergeAreas)
        {
            result.Add(new MergeCellInfo
            {
                Address = area.Address(),
                RowCount = area.Rows.Count,
                ColumnCount = area.Columns.Count
            });
            Marshal.ReleaseComObject(area);
        }
        Marshal.ReleaseComObject(mergeAreas);

        return result.ToArray();
    }
}

public class RangeInfo
{
    public string Address { get; set; }
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public object[,] Values { get; set; }
    public string[,] Formulas { get; set; }
    public FormatConditionInfo[] FormatConditions { get; set; }
    public MergeCellInfo[] MergeCells { get; set; }
}
```

- [ ] **Step 3: 提交**

```bash
git add src/DeepExcel.AddIn/Perception/
git commit -m "feat: 实现Excel结构感知层"
```

---

## 阶段四：执行引擎实现

### Task 5: 实现VBA执行引擎

**Files:**
- Create: `src/DeepExcel.AddIn/Executor/VBAExecutor.cs`
- Create: `src/DeepExcel.AddIn/Executor/SnapshotManager.cs`

**Interfaces:**
- Consumes: VBA代码字符串
- Produces: 执行结果或错误信息

- [ ] **Step 1: 创建VBAExecutor**

```csharp
public class VBAExecutor
{
    private readonly Application _app;
    private readonly SnapshotManager _snapshots;

    public VBAExecutor(Application app, SnapshotManager snapshots)
    {
        _app = app;
        _snapshots = snapshots;
    }

    public ExecutionResult Execute(string vbaCode, string macroName = "DeepExcel_TempMacro")
    {
        // 执行前快照
        var snapshotId = _snapshots.CreateSnapshot();
        
        try
        {
            // 注入VBA代码到活动工作簿
            var vbProject = _app.ActiveWorkbook.VBProject;
            var module = vbProject.Modules.Add();
            module.Name = "DeepExcelModule";
            module.CodeModule.AddFromString(vbaCode);

            // 执行宏
            _app.Run($"DeepExcelModule.{macroName}");

            // 清理临时模块
            vbProject.Modules.Remove(module);

            return new ExecutionResult { Success = true, SnapshotId = snapshotId };
        }
        catch (Exception ex)
        {
            // 执行失败，自动回滚
            _snapshots.Rollback(snapshotId);
            return new ExecutionResult { Success = false, Error = ex.Message };
        }
    }
}

public class ExecutionResult
{
    public bool Success { get; set; }
    public string Error { get; set; }
    public string SnapshotId { get; set; }
}
```

- [ ] **Step 2: 创建SnapshotManager**

```csharp
public class SnapshotManager
{
    private readonly Application _app;
    private readonly string _snapshotFolder;

    public SnapshotManager(Application app)
    {
        _app = app;
        _snapshotFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DeepExcel", "Snapshots");
        Directory.CreateDirectory(_snapshotFolder);
    }

    public string CreateSnapshot()
    {
        var wb = _app.ActiveWorkbook;
        if (wb == null) return null;

        var snapshotId = Guid.NewGuid().ToString();
        var backupPath = Path.Combine(_snapshotFolder, $"{snapshotId}.xlsx");
        
        wb.SaveCopyAs(backupPath);
        
        return snapshotId;
    }

    public bool Rollback(string snapshotId)
    {
        if (string.IsNullOrEmpty(snapshotId)) return false;

        var backupPath = Path.Combine(_snapshotFolder, $"{snapshotId}.xlsx");
        if (!File.Exists(backupPath)) return false;

        try
        {
            var wb = _app.ActiveWorkbook;
            wb.Save();
            wb.Close();
            
            _app.Workbooks.Open(backupPath);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
```

- [ ] **Step 3: 提交**

```bash
git add src/DeepExcel.AddIn/Executor/
git commit -m "feat: 实现VBA执行引擎和快照管理"
```

---

## 阶段五：模型适配层

### Task 6: 实现多模型支持（Claude/DeepSeek）

**Files:**
- Create: `src/DeepExcel.Models/IModelAdapter.cs`
- Create: `src/DeepExcel.Models/ClaudeAdapter.cs`
- Create: `src/DeepExcel.Models/DeepSeekAdapter.cs`
- Create: `src/DeepExcel.Models/ModelConfig.cs`

**Interfaces:**
- Consumes: 结构化Excel数据 + 用户意图
- Produces: 工具调用响应

- [ ] **Step 1: 创建模型接口**

```csharp
public interface IModelAdapter
{
    string ModelName { get; }
    Task<ModelResponse> SendMessageAsync(ModelRequest request);
    Task<bool> TestConnectionAsync(string apiKey);
}

public class ModelRequest
{
    public string UserMessage { get; set; }
    public ExcelContext Context { get; set; }
    public List<ToolDefinition> AvailableTools { get; set; }
}

public class ModelResponse
{
    public string Content { get; set; }
    public ToolCall ToolCall { get; set; }
    public bool NeedsMoreContext { get; set; }
}

public class ExcelContext
{
    public WorkbookStructure Workbook { get; set; }
    public RangeInfo CurrentSelection { get; set; }
}
```

- [ ] **Step 2: 创建ClaudeAdapter**

```csharp
public class ClaudeAdapter : IModelAdapter
{
    public string ModelName => "claude-3-5-sonnet";
    
    public async Task<ModelResponse> SendMessageAsync(ModelRequest request)
    {
        // 使用HTTPClient调用Claude API
        var client = new HttpClient();
        client.DefaultRequestHeaders.Add("x-api-key", Config.ApiKey);
        client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");

        var payload = new
        {
            model = ModelName,
            max_tokens = 1024,
            messages = new[] {
                new { role = "user", content = BuildPrompt(request) }
            },
            tools = request.AvailableTools.Select(t => new {
                name = t.Name,
                description = t.Description,
                input_schema = t.InputSchema
            })
        };

        var response = await client.PostAsync(
            "https://api.anthropic.com/v1/messages",
            new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json"));

        return ParseResponse(await response.Content.ReadAsStringAsync());
    }

    private string BuildPrompt(ModelRequest request)
    {
        var sb = new StringBuilder();
        sb.AppendLine("你是一个Excel AI助手。用户的需求:");
        sb.AppendLine(request.UserMessage);
        sb.AppendLine("\n当前Excel上下文:");
        sb.AppendLine($"工作簿: {request.Context.Workbook?.Name}");
        sb.AppendLine($"工作表: {request.Context.CurrentSelection?.Address}");
        return sb.ToString();
    }

    public async Task<bool> TestConnectionAsync(string apiKey)
    {
        // 测试API连接
        return true;
    }
}
```

- [ ] **Step 3: 创建DeepSeekAdapter**

```csharp
public class DeepSeekAdapter : IModelAdapter
{
    public string ModelName => "deepseek-chat";
    
    public async Task<ModelResponse> SendMessageAsync(ModelRequest request)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {Config.ApiKey}");

        var payload = new
        {
            model = ModelName,
            messages = new[] {
                new { role = "user", content = BuildPrompt(request) }
            }
        };

        var response = await client.PostAsync(
            "https://api.deepseek.com/v1/chat/completions",
            new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json"));

        return ParseResponse(await response.Content.ReadAsStringAsync());
    }

    public async Task<bool> TestConnectionAsync(string apiKey)
    {
        return true;
    }
}
```

- [ ] **Step 4: 提交**

```bash
git add src/DeepExcel.Models/
git commit -m "feat: 实现多模型适配层"
```

---

## 阶段六：集成与P0验证

### Task 7: 端到端集成测试

**Files:**
- Modify: `src/DeepExcel.AddIn/ThisAddIn.cs` - 集成所有组件
- Create: `src/DeepExcel.AddIn/Agent/Orchestrator.cs`

**Interfaces:**
- Consumes: 用户聊天消息
- Produces: 执行结果流式回传UI

- [ ] **Step 1: 创建Orchestrator编排器**

```csharp
public class Orchestrator
{
    private readonly WorkbookAnalyzer _workbookAnalyzer;
    private readonly RangeAnalyzer _rangeAnalyzer;
    private readonly VBAExecutor _vbaExecutor;
    private readonly IModelAdapter _modelAdapter;

    public Orchestrator(
        WorkbookAnalyzer workbookAnalyzer,
        RangeAnalyzer rangeAnalyzer,
        VBAExecutor vbaExecutor,
        IModelAdapter modelAdapter)
    {
        _workbookAnalyzer = workbookAnalyzer;
        _rangeAnalyzer = rangeAnalyzer;
        _vbaExecutor = vbaExecutor;
        _modelAdapter = modelAdapter;
    }

    public async Task<OrchestrationResult> ProcessAsync(string userMessage)
    {
        // 1. 感知层：收集上下文
        var context = new ExcelContext
        {
            Workbook = _workbookAnalyzer.Analyze(),
            CurrentSelection = _rangeAnalyzer.Analyze(
                Globals.ThisAddIn.Application.Selection as Range)
        };

        // 2. 发送给模型
        var request = new ModelRequest
        {
            UserMessage = userMessage,
            Context = context,
            AvailableTools = GetToolDefinitions()
        };

        var response = await _modelAdapter.SendMessageAsync(request);

        // 3. 执行工具调用
        if (response.ToolCall != null)
        {
            return await ExecuteToolCallAsync(response.ToolCall);
        }

        return new OrchestrationResult { Content = response.Content };
    }

    private async Task<OrchestrationResult> ExecuteToolCallAsync(ToolCall toolCall)
    {
        switch (toolCall.Name)
        {
            case "generate_vba":
                var vbaCode = ExtractVBAFromArgs(toolCall.Arguments);
                var result = _vbaExecutor.Execute(vbaCode);
                return new OrchestrationResult 
                { 
                    Success = result.Success,
                    Content = result.Success ? "VBA执行成功" : result.Error
                };
            default:
                return new OrchestrationResult { Content = $"Unknown tool: {toolCall.Name}" };
        }
    }
}
```

- [ ] **Step 2: 提交**

```bash
git add src/DeepExcel.AddIn/Agent/
git commit -m "feat: 实现Agent编排器"
```

---

## P0验证清单

完成以上所有Task后，验证以下能力：

- [ ] VSTO加载项成功加载到Excel功能区
- [ ] WebView2面板显示聊天UI
- [ ] 选中单元格后，"读取选区"返回正确结构
- [ ] 用户输入"在A1写入公式=SUM(B1:B10)"，Agent生成VBA并执行
- [ ] VBA执行后A1显示正确结果
- [ ] 执行前自动快照，失败后能回滚

---

## Plan文件位置

`docs/superpowers/plans/2026-06-25-DeepExcel-P0-implementation.md`
