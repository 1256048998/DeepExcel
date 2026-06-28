# DeepExcel 架构设计文档

> 一句话定位：**住在 Excel 里、能看着表格自己动手干活的 AI Agent 桌面插件。**
> 技术对标 Excel Master，支持多模型（Claude/DeepSeek等）可热切换。

文档版本：v0.1
最后更新：2026-06-25

---

## 0. 已拍板的决策

| 决策项 | 选择 | 含义 |
|---|---|---|
| **技术地基** | VSTO 桌面插件（C#/.NET） | 深度操作 Excel 对象模型；Windows + 桌面 only |
| **产品范围** | 执行型 AI Agent | 读取结构 → 写代码(VBA/Python) → 执行 → 验证 → 回滚 |
| **模型与后端** | 本地API Key直连（MVP）+ 多模型支持 | 优先Claude，支持DeepSeek/其他模型热切换 |
| **UI层** | WebView2 + React/TypeScript | 现代聊天界面，可复用到未来网页版 |

---

## 1. 产品定位与差异化

### 1.1 本质
- **Excel 的价值是"算对"**：AI 的手是 VBA/Python，操作**数据**，对错有标准答案（数字正确）。
- **核心差异化**：不是助手给建议，而是**自主执行**——读懂真实表格结构、生成完整代码、首次就能跑通、出错自动重试。

### 1.2 与Excel Master的架构映射

| 维度 | Excel Master | DeepExcel |
|---|---|---|
| 宿主 | Excel COM/VSTO 加载项 | 同 |
| UI | WebView2 聊天侧边栏 | 同 |
| 感知 | 读单元格/区域/合并单元格/命名区域/条件格式 | 同 |
| 执行 | VBA + Python | VBA + Python（扩展） |
| 安全 | 自动备份 + 回滚 | 快照备份 + 回滚 + 自动重试 |
| 后端 | LLM代理 + 积分 | 本地API Key直连（MVP）|
| 模型 | Claude主力 | **多模型支持（Claude/DeepSeek等可热切换）**|

---

## 2. 整体技术架构（分层）

```
┌──────────────────────────────────────────────────────────────┐
│  ① 宿主插件层  DeepExcel.AddIn  (C#/.NET, VSTO)             │
│     · Ribbon + 生命周期管理                                   │
│     · COM Interop ↔ Excel 对象模型                           │
│     · WebView2 任务面板承载                                  │
├──────────────────────────────────────────────────────────────┤
│  ② 交互 UI 层  DeepExcel.UI  (WebView2 + React/TypeScript)│
│     · 聊天侧边栏 · 流式输出 · 每步可视化 · 改动预览 · 撤销   │
├──────────────────────────────────────────────────────────────┤
│  ③ 桥接层  Bridge (JS ↔ C# 双向消息)                        │
│     · UI 发意图/工具调用 → C# 执行 → 流式回传结果           │
├──────────────────────────────────────────────────────────────┤
│  ④ 感知层  Perception                                       │
│     · 结构读取：工作表/单元格/区域/合并单元格/命名区域/条件格式│
│     · 视觉渲染：Range → PNG → 喂多模态模型（可选）           │
├──────────────────────────────────────────────────────────────┤
│  ⑤ Agent 编排层  Orchestrator (工具调用循环)               │
│     · 读 → 规划 → 改 → 验证 → 修 → 备份/回滚               │
├──────────────────────────────────────────────────────────────┤
│  ⑥ 执行引擎  Executor                                       │
│     · VBA 执行（首选）+ Python（需装库）+ OpenXML 兜底       │
│     · 操作前快照备份；失败自动回滚重试                         │
├──────────────────────────────────────────────────────────────┤
│  ⑦ 模型层  Model Adapter                                     │
│     · 多模型支持：Claude / DeepSeek / 其他（可热切换）      │
│     · 统一接口封装，模型可配置                                │
└──────────────────────────────────────────────────────────────┘
```

### 2.1 Agent 循环跑在哪里（关键架构取舍）
**本地为主（MVP）**：
- 客户端（本地）跑：Agent 循环 + 工具执行
- 低延迟、文件不离机，只把**相关结构数据**发给模型
- 模型API Key由用户本地配置

> 隐私承诺：**原始 .xlsx 永不上传**，只上传当前任务相关的结构数据。

### 2.2 多模型支持
- 通过统一的 `IModelAdapter` 接口封装不同模型
- 模型可在配置文件中切换（热切换，无需重发安装包）
- MVP阶段主力Claude，支持DeepSeek扩展

---

## 3. 核心模块详解

### ① 宿主插件层 DeepExcel.AddIn（C#/.NET, VSTO）
- **职责**：作为 COM 加载项嵌入 Excel，注册 Ribbon，管理任务面板生命周期，持有对 `Application / Workbook / Worksheet / Range` 等对象模型的引用。
- **关键点**：
  - 对象模型操作必须在 **UI 线程（STA）** 上跑，跨线程会抛 COM 异常 → 需要"调度到主线程"的封装
  - 释放 COM 引用要规范（`Marshal.ReleaseComObject`），否则 Excel 进程残留、卡死
  - 加载项被禁用/不加载是这类产品最大的售后坑 → 需要健康自检与一键修复

### ② 交互 UI 层 DeepExcel.UI（WebView2 + React/TS）
- **职责**：聊天界面、流式 token、Agent 每一步的可视化（"正在写入公式…"）、改动前后对比、一键撤销
- **为什么用 WebView2**：现代聊天 UI 用 Web 技术开发快、好看、可复用未来网页版的组件
- **关键点**：WebView2 需要 Evergreen Runtime（安装包要带/检测）

### ③ 桥接层 Bridge
- **机制**：`WebView2.CoreWebView2.PostMessage`（C#→JS）+ `window.chrome.webview.postMessage`（JS→C#）
- **协议**：定义一套 JSON 消息（`type` + `payload`），例如 `tool_call` / `tool_result` / `stream_delta`
- **关键点**：所有 Excel 操作的"真身"在 C#，UI 只发意图、收结果

### ④ 感知层 Perception
两条腿走路：
- **结构感知（文本）**：遍历对象模型，导出当前选区/工作表的结构化描述——单元格值、公式、格式、合并区域、命名区域、条件格式、数据透视表
- **视觉感知（图像，可选）**：`Range.CopyPicture()` → PNG → 作为图片输入给多模态模型

### ⑤ Agent 编排层 Orchestrator
- **循环**：`读结构 → 规划 → 调工具改 → 验证 → 不满意自动修 → 满意则备份落地`
- **工具调用（function calling）** 由模型驱动，工具清单见 §4
- **校验闭环**：改完读取结果校验，是 Excel 版的"执行后校验结果"

### ⑥ 执行引擎 Executor（三层逃生舱）
1. **VBA 执行（首选）**：生成 VBA 代码，注入到 Excel VBA 模块并执行
2. **Python（扩展）**：生成 Python 脚本，调用 pandas/openpyxl 执行（需用户环境有Python）
3. **OpenXML 兜底（保真）**：直接改 .xlsx 的 XML，保证保真
- **安全网**：每个**批量操作前**对当前工作簿（或受影响工作表）做快照备份；执行失败或用户拒绝 → 回滚

### ⑦ 模型层 Model Adapter
- **接口抽象**：定义统一 `IModelAdapter` 接口
- **实现**：ClaudeAdapter / DeepSeekAdapter / [扩展]
- **热切换**：通过配置文件指定模型，无需重发安装包

---

## 4. 工具集设计（Agent 的"手"）

### 4.1 感知类（共用）
- `read_workbook` 读整本工作簿结构概览（工作表名、命名区域）
- `read_worksheet` 读某工作表的详细结构
- `read_range` 读某区域的单元格值、公式、格式
- `get_selection` 读当前用户选中的区域
- `detect_merge_cells` 检测合并单元格布局
- `list_defined_names` 列出命名区域

### 4.2 公式类
- `generate_formula` 理解自然语言，生成并插入公式 ✅ P0

### 4.3 VBA执行类
- `generate_vba` 理解表格结构，生成完整VBA模块 ✅ P0
- `execute_vba` 执行生成的VBA代码
- `debug_vba` 出错自动分析并重试

### 4.4 Python执行类
- `generate_python` 生成Python自动化脚本 ✅ P1
- `execute_python` 执行Python脚本（需用户环境有Python）

### 4.5 数据处理类
- `clean_data` 统一格式、删除重复、标红缺失值 ✅ P1
- `create_pivot_table` 创建数据透视表 ✅ P1
- `create_chart` 根据数据生成图表 ✅ P1

### 4.6 安全类（共用）
- `snapshot` 操作前快照 ✅ P0
- `rollback` 回滚到快照 ✅ P0
- `preview_diff` 生成改动前后对比给用户确认

---

## 5. 数据流与隐私

```
用户在面板输入需求
   → UI 发意图给 C#
   → 感知层导出"相关结构数据"（不含整本原文件）
   → 本地模型适配器发送到对应模型（鉴权由API Key处理）
   → 模型返回工具调用
   → C# 执行引擎在本地 Excel 上操作（操作前已快照）
   → 流式把结果回传 UI
   → 用户确认/撤销
```

隐私红线：
- 原始 .xlsx **永不上传**，仅上传任务相关的结构数据
- API Key 存储在本地配置，不上传

---

## 6. 技术栈选型清单

| 层 | 选型 | 说明 |
|---|---|---|
| 宿主插件 | C# / .NET Framework 4.7.2+ / **VSTO** | Excel COM 加载项标准路线 |
| 对象模型 | Microsoft.Office.Interop.Excel | COM Interop |
| 保真兜底 | **Open XML SDK** (DocumentFormat.OpenXml) | 直接改 .xlsx XML |
| 面板宿主 | **WebView2** (Evergreen) | 承载 Web 聊天 UI |
| 前端 | React + TypeScript + Vite | 现代聊天界面 |
| Agent/模型 | **Claude / DeepSeek / 其他** via 本地API | 多模型可热切换 |
| 安装包 | **Inno Setup**（带 /Qp 静默） | 检测 .NET / WebView2 运行时 |

---

## 7. 分阶段路线图

| 阶段 | 目标 | 功能范围 | 验证点 |
|---|---|---|---|
| **P0 架构打通** | 跑通"感知→执行"管线 | 读取Excel结构 + VBA生成执行（最简闭环） | VSTO+WebView2+VBA执行 三件套打通 |
| **P1 MVP 闭环** | 单任务Agent闭环 | 公式生成 + VBA执行 + 数据清洗 + Python扩展 | 用户能在真实工作簿上看到 Agent 干活 |
| **P2 能力扩展** | 覆盖主流场景 | VBA + Python + 图表/透视表 + 多步工作流 | 真实财务/数据分析场景可用 |
| **P3 商业化** | 可上线收费 | 积分制 + 用户体系 + 提示词热更新 + 遥测 | 后端完整计费系统 |

---

## 8. 风险与对策

| 风险 | 影响 | 对策 |
|---|---|---|
| 加载项被 Excel 禁用/不加载 | 用户打不开，售后高发 | 健康自检 + 一键修复；注册表/COM 注册规范 |
| COM 线程/引用泄漏导致 Excel 卡死残留 | 体验崩、口碑差 | 统一主线程调度封装 + 严格释放 COM 引用 |
| VBA 代码执行破坏用户数据 | 损坏用户文件 | 操作前强制快照 + 失败回滚 + 用户确认 |
| 多模型切换后效果不一致 | 核心体验不稳定 | 统一工具定义 + 提示词模板化 |
| VSTO 属于微软"维护态"技术 | 长期演进受限 | UI 用 Web 技术解耦，未来可向网页版迁移 |

---

## 9. 建议代码目录结构

```
DeepExcel/
├─ docs/
│  └─ 架构设计.md                    ← 本文档
├─ src/
│  ├─ DeepExcel.AddIn/               C# VSTO 加载项（Ribbon / COM / WebView2 宿主）
│  ├─ DeepExcel.Core/               C# 核心（感知 / 执行引擎 / 工具实现 / 桥接）
│  ├─ DeepExcel.UI/                 React+TS 前端（聊天面板，构建产物注入 WebView2）
│  └─ DeepExcel.Models/             模型适配器（Claude / DeepSeek / 其他）
├─ installer/                        Inno Setup 脚本
└─ README.md
```
