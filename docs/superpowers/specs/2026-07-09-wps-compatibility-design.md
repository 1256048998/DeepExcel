# WPS Office 兼容性支持设计

> **日期**：2026-07-09
> **状态**：规划中（待用户审阅）
> **前提**：用户已休息，本设计基于调研结果和合理假设自主完成，用户醒来后审阅

## 1. 背景与目标

### 1.1 为什么需要 WPS 支持

DeepExcel 当前仅支持 Microsoft Excel。WPS Office 在中国市场占有率高（尤其政企用户），支持 WPS 可大幅扩展用户群体。

### 1.2 目标

- 让 DeepExcel AI Agent 在 WPS 表格中运行，体验与 Excel 端基本一致
- 最大化复用现有代码（AI Agent 层、前端 UI）
- 不牺牲 Excel 端的任何功能和体验

### 1.3 非目标

- 不支持 WPS 文字 / WPS 演示（仅表格 ET）
- 不追求 100% 功能对等（VBA 相关功能允许降级为 JSA）
- 不迁移 Excel 端到 Web Add-ins（保持现有 COM 架构）

### 1.4 目标 WPS 版本

| 版本 | 支持优先级 | 说明 |
|------|-----------|------|
| WPS 专业版 / 专业增强版 | P0 | 自带 VBA，COM 加载项无白名单限制 |
| WPS 企业版 / 党政专用版 | P0 | 与专业版一致 |
| WPS 个人版（最新） | P1 | 需引导用户加白名单；VBA 需装插件 |
| WPS 2019 SP2 以下 | 不支持 | COM 加载项兼容性问题 |

## 2. 调研发现

### 2.1 DeepExcel 架构解耦优势

```
当前架构（关键洞察：三层解耦）：
┌─ 前端 UI 层（React + WebView2）     ← 完全不依赖 Excel COM
├─ AI Agent 层（Python sidecar）       ← 完全不依赖 Excel COM
└─ Excel 操作层（IExcelActions 实现）   ← 唯一需要适配的层
```

- **Python sidecar**（sidecar.py / excel_tools.py）：所有工具通过 `call_csharp` IPC 让 C# 执行，不直接调用 `win32com`。原样可复用。
- **前端 UI**（React）：通过 `bridge.ts` 的 `window.chrome.webview.postMessage` 与 C# 通信。bridge 层抽象后可复用 95%+。
- **Excel 操作层**（ToolDispatcher → ExcelActionsImpl / VBAExecutor / ChartTool）：依赖 `Microsoft.Office.Interop.Excel`，需要用 WPS JS API 重写。

### 2.2 WPS 加载项机制

WPS 支持三种扩展机制：
1. **COM 加载项**（IDTExtensibility2）— 与 Office 共享 CLSID，互操作库可互换
2. **JS 加载项**（wpsjs，官方主推）— 基于 Chromium，原生支持 taskpane
3. **WPS 宏**（JSA）— JS 宏编辑器，VBA 的替代方案

**关键障碍**：
- WPS COM 加载项**没有 VSTO 风格的 `CustomTaskPaneCollection`** → 无法停靠任务窗格
- WPS 个人版**默认无 VBA** → `Microsoft.Vbe.Interop` 不可靠
- WPS 12.0.1.17xx+ COM 加载项需**白名单**
- WPS 个人版 12.1.0.16910+ 限制 oem.ini 方式加载 JS 加载项

**关键利好**：
- WPS JS 加载项的 taskpane **原生是嵌入式 Chromium 网页容器** → 直接替代 WebView2 + CTP
- WPS JS API 对象模型与 Excel COM **高度一致**（Application/Workbook/Worksheet/Range）
- WPS JSA 支持 ES6，对象模型与 VBA 一致 → VBA 的合理替代
- WPS 与 Office **共享 CLSID**，`Microsoft.Office.Interop.Excel.dll` 可互换

### 2.3 版本阻断问题

| 版本 | 影响 | 缓解 |
|------|------|------|
| 个人版 12.1.0.16910+ | JS 加载项 oem.ini 方式受限 | 用 `wpsjs publish` 重新发布 |
| 12.0.1.17xx+ | COM 加载项需白名单 | 引导用户在 WPS 选项中手动加白名单 |

## 3. 方案选型

### 方案 A：最小改动（保 COM 架构 + 独立窗口）❌

在现有 C# DLL 基础上加 WPS 注册路径，把 CTP 改为独立 WinForm 窗口。

- **致命缺陷**：无停靠任务窗格（体验严重降级）；VBA 不可靠；WPS 个人版白名单阻断
- **评价**：在"将就"，不是 AI Native 思维。用户在 AI 对话框操作，面板却弹成独立浮动窗口——体验割裂

### 方案 B：双轨架构（COM + WPS JS 加载项）✅ 推荐

Excel 端保持现有 COM 加载项完全不动；WPS 端新建 JS 加载项，用 taskpane 承载前端，用 JSA 替代 VBA，Python sidecar 完全复用。

- **优点**：两端各自最佳体验；AI Agent 层 100% 复用；前端 95% 复用；用各平台原生方式集成
- **代价**：Excel 操作层需用 WPS JS API 重写（但接口对齐 `IExcelActions`）

### 方案 C：完全 Web 化（Office Web Add-ins + WPS JS 加载项）❌

抛弃 COM，全面 Web 化。

- **缺陷**：Office.js API 表面受限（精细 Range 操作、VBA 全部丢失）；重写量最大；当前无跨平台需求

### 选型结论

**采用方案 B**。理由：
1. 不牺牲 Excel 端任何体验（COM 加载项原样保留）
2. WPS 端用官方推荐的 JS 加载项，taskpane 原生停靠
3. AI Agent（Python sidecar）100% 复用，是项目最核心的资产
4. 前端 UI 95%+ 复用，仅需适配 bridge 层
5. 符合 AI Native 思维：各平台用原生最佳方式集成，不做 hack

## 4. 架构设计

### 4.1 整体架构

```
┌─────────────────────────────────────────────────────────┐
│            Python Sidecar (100% 复用)                     │
│  Claude Agent SDK + MCP Tools + PreToolUse Hook          │
│  stdin/stdout JSON Lines IPC                             │
└──────────────┬────────────────────┬──────────────────────┘
               │                    │
     ┌─────────▼─────────┐  ┌──────▼────────────────┐
     │  Excel 端 (现有)    │  │  WPS 端 (新建)          │
     │  C# COM AddIn      │  │  WPS JS 加载项          │
     │  ├─ CTP + WebView2 │  │  ├─ taskpane (Chromium) │
     │  ├─ ToolDispatcher  │  │  ├─ ToolDispatcher.js   │
     │  │  → IExcelActions │  │  │  → WpsActions (JSA)  │
     │  └─ VBAExecutor    │  │  └─ JsaExecutor         │
     └───────────────────┘  └───────────────────────┘
               │                    │
     ┌─────────▼─────────┐  ┌──────▼────────────────┐
     │  React 前端 (95%复用)│  │  React 前端 (95%复用)   │
     │  bridge.ts → C#    │  │  bridge.ts → JS post   │
     └───────────────────┘  └───────────────────────┘
```

### 4.2 WPS 端组件设计

#### 4.2.1 WPS JS 加载项项目结构

```
src/DeepExcel.Wps/
├── ribbon.xml              # Ribbon UI 定义（复用 Excel 端设计）
├── jsplugins.xml           # WPS 加载项清单
├── package.json            # wpsjs 依赖
├── taskpane.html           # taskpane 入口（加载 React 构建产物）
├── main.js                 # 加载项入口（启动 Python sidecar + 注册 Ribbon）
├── sidecar-host.js         # Python sidecar 进程管理（对应 PythonSidecar.cs）
├── tool-dispatcher.js      # 工具调度（对应 ToolDispatcher.cs）
├── wps-actions.js          # WPS JS API 实现（对应 ExcelActionsImpl）
├── jsa-executor.js         # JSA 宏执行（对应 VBAExecutor）
└── web/                    # React 前端构建产物（与 Excel 端共享源码）
    ├── index.html
    └── assets/
```

#### 4.2.2 main.js — 加载项入口

```js
const { spawn } = require('child_process')
const SidecarHost = require('./sidecar-host')

let sidecar = null

// Ribbon 按钮回调：打开 taskpane
function OnAction(control) {
  wps.ShowTaskPane({
    url: 'taskpane.html',
    title: 'DeepExcel AI',
    width: 400,
    dockRight: true
  })
}

// 加载项启动
function OnPluginInit() {
  // 启动 Python sidecar 子进程
  sidecar = new SidecarHost({
    pythonPath: getPythonPath(),
    sidecarPath: getSidecarPath()
  })
  sidecar.start()
}

// 加载项卸载
function OnPluginDestroy() {
  if (sidecar) sidecar.stop()
}
```

#### 4.2.3 sidecar-host.js — Python sidecar 进程管理

对应 C# 端的 `PythonSidecar.cs`，负责：
- `child_process.spawn` 启动 Python sidecar
- stdin/stdout JSON Lines 协议（与 C# 端完全一致）
- `tool_call` 消息路由到 `tool-dispatcher.js`
- `stream_delta` / `stream_end` / `clarify` 消息转发到 taskpane
- 进程健康检查 + 自动重启
- 环境变量 `PYTHONUTF8=1`（与 C# 端一致，避免编码问题）

#### 4.2.4 tool-dispatcher.js — 工具调度

对应 C# 端的 `ToolDispatcher.cs`，负责：
- 接收来自 Python sidecar 的 `tool_call` 消息
- 路由到 `wps-actions.js`（对应 `IExcelActions`）或 `jsa-executor.js`
- 构建 Excel 快照（活动工作簿/工作表/选区信息）返回给 sidecar
- 截图工具（用 WPS 窗口句柄 + GDI，与 C# 端逻辑一致）
- 工具执行结果通过 IPC 返回 Python sidecar

#### 4.2.5 wps-actions.js — WPS JS API 实现

对应 C# 端的 `ExcelActionsImpl`，用 WPS JS API 实现所有 `IExcelActions` 方法：

```js
// 对齐 IExcelActions 接口
const WpsActions = {
  ReadRange(address) {
    const app = wps.Application
    const range = app.Range(address)
    return {
      address: range.Address,
      values: range.Value2,  // 2D 数组
      formulas: range.Formula,
      rowCount: range.Rows.Count,
      columnCount: range.Columns.Count
    }
  },

  WriteFormula(address, formula) {
    wps.Application.Range(address).Formula = formula
  },

  WriteValue(address, value) {
    wps.Application.Range(address).Value2 = value
  },

  SortData(rangeAddress, sortColumn, descending, hasHeader) {
    const range = wps.Application.Range(rangeAddress)
    range.Sort(range.Columns(sortColumn), descending ? 2 : 1, {
      Header: hasHeader ? 1 : 0
    })
  },

  // ... 其他方法对齐 IExcelActions
}
```

**API 对照表**：

| IExcelActions 方法 | WPS JS API | 兼容性 |
|---|---|---|
| ReadRange | `Range.Value2` | ✅ |
| WriteFormula | `Range.Formula =` | ✅ |
| WriteValue | `Range.Value2 =` | ✅ |
| WriteRange | `Range.Value2 = 2D array` | ✅ |
| ReadWorkbook | `Application.Workbooks` 遍历 | ✅ |
| GetSelection | `Application.Selection` | ✅ |
| SortData | `Range.Sort()` | ✅ |
| FilterData | `Range.AutoFilter()` | ✅ |
| MergeCells | `Range.Merge()` | ✅ |
| FreezePanes | `Application.ActiveWindow.FreezePanes` | ✅ |
| SetNumberFormat | `Range.NumberFormat =` | ✅ |
| InsertRows / DeleteRows | `Range.Insert() / Delete()` | ✅ |
| CopyRange / ClearRange | `Range.Copy() / Clear()` | ✅ |
| CreateChart | `Application.Charts.Add()` | ⚠ 需验证 |
| CreatePivotTable | `PivotTableWizard` | ⚠ 需验证 |
| ApplyConditionalFormat | `FormatConditions.Add()` | ⚠ 需验证 |
| ExecuteVBA | → JsaExecutor | ❌ 改为 JSA |
| ExecutePython | 独立 Python 进程 | ✅ |

#### 4.2.6 jsa-executor.js — JSA 宏执行

对应 C# 端的 `VBAExecutor`，用 WPS JSA 替代 VBA：

```js
const JsaExecutor = {
  ExecuteJsa(code) {
    // 方式 1：通过 Application.Run 调用预注册的 JSA 宏
    // 方式 2：通过 wps.Application.Macro.JSEval 直接执行 JS 代码
    try {
      wps.Application.Run(code)
      return { success: true }
    } catch (e) {
      return { success: false, error: e.message }
    }
  }
}
```

**VBA → JSA 转换策略**：
- AI 模型在 system_prompt 中被告知当前宿主是 WPS，应使用 `execute_jsa` 工具写 JS 宏而非 VBA
- JSA 语法与 VBA 差异（`Dim x` → `let x`，`Sub/End Sub` → `function`）由 AI 自动适配
- 对于复杂 VBA 代码，AI 可直接写等价的 JS 代码

#### 4.2.7 前端 bridge 层适配

现有 `bridge.ts`：
```ts
// Excel 端：通过 WebView2 与 C# 通信
export function sendToHost(msg) {
  window.chrome.webview.postMessage(msg)
}
```

适配后：
```ts
// 抽象为接口，按环境选择实现
interface IBridge {
  sendToHost(msg: any): void
  onHostMessage(handler: (data: any) => void): () => void
  sendToHostWithResponse(msg: any, responseType: string): Promise<any>
}

// Excel 端实现（现有）
class WebView2Bridge implements IBridge { ... }

// WPS 端实现（新增）
class WpsTaskpaneBridge implements IBridge {
  sendToHost(msg) {
    // taskpane 在 iframe 中，通过 postMessage 与父窗口通信
    window.parent.postMessage(msg, '*')
  }
  onHostMessage(handler) {
    window.addEventListener('message', e => handler(e.data))
    return () => window.removeEventListener('message', handler)
  }
}

// 编译时按环境选择
export const bridge: IBridge = ENV === 'wps'
  ? new WpsTaskpaneBridge()
  : new WebView2Bridge()
```

### 4.3 注册机制

新增 `scripts/register-wps.ps1`：

```powershell
# 1. COM CLSID 注册（与 Office 共用，调用现有 register-user.ps1 的逻辑）

# 2. WPS 加载项白名单
$wpsAddinKey = "HKCU:\Software\kingsoft\office\ET\AddinsWL\DeepExcel.AddIn"
New-Item -Path $wpsAddinKey -Force | Out-Null
Set-ItemProperty -Path $wpsAddinKey -Name "(default)" -Value "" -Force

# 3. JS 加载项发布（wpsjs publish）
# 或手动复制 JS 加载项文件到 WPS 加载项目录
```

### 4.4 Python sidecar 工具适配

sidecar.py 和 excel_tools.py **不需要修改**。所有 `call_csharp` 调用通过 IPC 发送 `tool_call` 消息，WPS 端的 `tool-dispatcher.js` 接收后路由到 `wps-actions.js`。

唯一需要新增的是 `execute_jsa` 工具（替代 `execute_vba`），在 `excel_tools.py` 中注册：

```python
@tool("execute_jsa", "执行 WPS JS 宏代码", {"code": str})
async def execute_jsa(args):
    result = await call_csharp("execute_jsa", {"code": args["code"]})
    return _wrap_result(result)
```

system_prompt.py 根据宿主类型（Excel/WPS）提示 AI 使用 `execute_vba` 或 `execute_jsa`。

### 4.5 system_prompt 适配

system_prompt.py 增加宿主类型感知：

```python
# C# 端在 BuildExcelSnapshot 中返回 host_type
host_type = context.get("host_type", "excel")  # "excel" or "wps"

if host_type == "wps":
    prompt += """
    注意：当前宿主是 WPS 表格。
    - 使用 execute_jsa 工具执行 JS 宏（而非 execute_vba）
    - JSA 语法为 ES6，对象模型与 VBA 一致（Application/Workbook/Worksheet/Range）
    - 截图和键盘操作行为与 Excel 一致
    """
```

## 5. 数据流

### 5.1 WPS 端完整数据流

```
用户在 taskpane 输入消息
  → React 前端 (WpsTaskpaneBridge.sendToHost)
  → main.js (postMessage 接收)
  → sidecar-host.js (stdin 写入 JSON)
  → Python sidecar (Claude Agent SDK 处理)
  → excel_tools.py (call_csharp 发送 tool_call)
  → sidecar-host.js (stdout 读取 tool_call)
  → tool-dispatcher.js (路由)
  → wps-actions.js (WPS JS API 调用)
  → WPS 表格 (执行操作)
  → 结果返回: wps-actions → tool-dispatcher → sidecar-host → stdin → Python
  → Claude Agent 继续推理
  → stream_delta → sidecar-host → main.js → postMessage → React 前端渲染
```

### 5.2 与 Excel 端数据流对比

| 步骤 | Excel 端 | WPS 端 |
|------|---------|--------|
| 前端→宿主 | `chrome.webview.postMessage` | `window.parent.postMessage` |
| 宿主→sidecar | `Process.StandardInput.WriteLine` | `child.stdin.write` |
| sidecar→宿主 | `Process.StandardOutput` 异步读取 | `child.stdout.on('data')` |
| 宿主→表格 | `Microsoft.Office.Interop.Excel` | `wps.Application` JS API |
| 宿主→前端 | `chrome.webview.postMessage` | `window.postMessage` |

**协议完全一致**（JSON Lines），只是宿主语言和通信通道不同。

## 6. 风险与缓解

| 风险 | 等级 | 缓解措施 |
|------|------|---------|
| WPS JS API 能力不足（某些 Excel 操作无等价 API） | 中 | 降级处理 + AI prompt 提示替代路径；逐工具验证 |
| JSA 与 VBA 语义差异 | 中 | AI 直接用 `execute_jsa` 写原生 JS；system_prompt 引导 |
| WPS 个人版白名单阻断 | 中 | 首次运行检测 + 引导用户加白名单 |
| 前端 bridge 层差异 | 低 | 抽象为 `IBridge` 接口，编译时选择实现 |
| Python sidecar 在 Node.js 环境稳定性 | 低 | 复用已验证的 IPC 协议；增加进程健康检查 |
| WPS taskpane 与 React 前端兼容性 | 低 | taskpane 是 Chromium 环境，React 原生支持 |
| WPS 版本差异导致 API 行为不一致 | 中 | 在 system_prompt 中提示 AI 注意版本差异 |

## 7. 实施优先级

### P0 — MVP（核心可用）

- WPS JS 加载项骨架（ribbon.xml + jsplugins.xml + taskpane.html）
- sidecar-host.js（Python sidecar 进程管理 + IPC）
- 前端 bridge 层适配（WpsTaskpaneBridge）
- 核心 10 个工具实现（read_range / write_value / write_formula / write_range / read_workbook / read_selection / sort_data / filter_data / add_sheet / freeze_panes）
- register-wps.ps1 注册脚本

### P1 — 功能完善

- 图表创建（create_chart）
- 透视表（create_pivot_table）
- 条件格式（apply_conditional_format）
- 数据清洗（clean_data）
- set_cell_style / set_number_format / insert_rows / delete_rows 等
- 截图工具（screenshot_excel）

### P2 — 高级功能

- JSA 执行器（execute_jsa）
- send_keys 工具
- 工作簿快照（create_snapshot / rollback）
- Python 执行器（execute_python）
- WPS 个人版白名单引导 UI

## 8. 测试策略

- **单元测试**：wps-actions.js 每个方法独立测试（mock `wps.Application`）
- **集成测试**：在真实 WPS 专业版中运行完整对话流程
- **兼容性测试**：WPS 专业版 + 个人版 + 企业版分别测试
- **对比测试**：同一任务在 Excel 端和 WPS 端执行，比较结果一致性

## 9. 待验证事项

以下事项需要在真实 WPS 环境中验证（用户醒来后安装 WPS 专业版测试）：

1. WPS JS 加载项的 taskpane 是否能加载 React 构建产物（SPA 路由兼容性）
2. `wps.Application` JS API 的 `Range.Value2` 是否返回 2D 数组（与 Excel COM 一致）
3. WPS JS 加载项（Node.js 环境）是否能用 `child_process.spawn` 启动 Python 子进程
4. `wps.Application.Run` 是否能执行动态注入的 JSA 代码
5. WPS taskpane 中的 `window.postMessage` 是否能与 React 前端正常通信
6. WPS 专用版/企业版的加载项白名单机制具体操作步骤

## 10. 参考资料

- [WPS 开放平台 - 加载项概述](http://open.wps.cn/previous/docs/client/wpsLoad)
- [C# 开发 Office 和 WPS COM 加载项](https://blog.csdn.net/weixin_40604069/article/147958219)
- [WPS 论坛 - COM 加载项白名单问题](https://forum.wps.cn/topic/53623)
- [WPS 社区 - CustomTaskPane 讨论](https://bbs.wps.cn/topic/2984)
- [WPS JS 宏 (JSA) 文档](https://bbs.wps.cn/topics/tag/11024)
- DeepExcel 项目内部调研报告（2026-07-09，两份子代理调研结果）
