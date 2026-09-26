# DeepExcel WPS 表格加载项部署指南

## 概述

DeepExcel 现支持 WPS Office 表格（ET）。WPS 端采用 JS 加载项架构，与 Excel 端的 COM 加载项完全独立，但共享：
- **AI Agent 层**（Python sidecar）100% 复用
- **前端 UI**（React）95%+ 复用
- **工具协议**（JSON Lines IPC）完全一致

## 系统要求

| 组件 | 要求 |
|------|------|
| 操作系统 | Windows 10/11 (64位) |
| WPS Office | 专业版/专业增强版/企业版（推荐）<br>个人版 12.0.1.17xx+（需手动加白名单） |
| Python | 3.10+（需配置在 PATH 中，或放在项目 `python-3.11-embed-amd64/` 目录） |
| Node.js | 16+（WPS JS 加载项运行时） |

## 部署步骤

### 1. 构建加载项

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build-wps.ps1
```

此脚本会：
- 构建 React 前端（Vite）
- 复制构建产物到 `src/DeepExcel.Wps/web/`
- 复制 Python sidecar 文件到 `src/DeepExcel.Wps/sidecar/`
- 验证文件完整性

### 2. 注册加载项到 WPS

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register-wps.ps1
```

此脚本会：
- 注册到 `HKCU\Software\kingsoft\office\ET\AddinsWL\DeepExcel`（WPS 白名单）
- 发布 `jsplugins.xml` 到 `%APPDATA%\kingsoft\wps\jsplugins\`
- 复制 JS 加载项文件到 WPS 加载项目录

> ⚠ **这条是开发期路径，和发给用户的安装器走的不是同一套机制。**
> `register-wps.ps1` 用的是 `jsplugins\jsplugins.xml`，而安装器用的是
> `jsaddons\publish.xml`（见 [DEPLOYMENT.md](DEPLOYMENT.md) 「关键防退化约束」：
> 新版 WPS 已限制 `jsplugins.xml/oem.ini` 部署路径）。
>
> 后果是：**在开发机上用本脚本验证通过，并不能说明用户装完能用。**
> 已有现场反馈是 Excel 端正常而 WPS 端装了用不了。验证 WPS 必须用真实安装器，
> 在一台干净机器上装完再看 Ribbon，见 DEPLOYMENT.md 的发布前检查第 4 条。

### 3. 启动 WPS 表格

重启 WPS 表格，应能在 Ribbon 中看到 "DeepExcel" 选项卡。点击 "打开面板" 按钮即可启动 AI 面板。

## 配置 AI 模型

首次使用前需配置 API：

1. 点击 Ribbon 中的 "DeepExcel" → "打开面板"
2. 在面板右上角点击齿轮图标（模型配置）
3. 填入 API Base URL、Model Name、API Key
4. 保存

## 已知限制

### WPS 个人版用户

- **VBA 不可用**：WPS 个人版默认不含 VBA，DeepExcel 会自动改用 JSA（JS 宏）
- **白名单限制**：WPS 12.0.1.17xx+ 需手动将 "DeepExcel" 加入信任的加载项列表
  - 路径：WPS → 选项 → 信任中心 → 加载项安全
- **oem.ini / jsplugins.xml 限制**：WPS 个人版 12.1.0.16910+ 限制 oem.ini 方式加载。
  新版对 `jsplugins.xml` 路径同样有限制，因此**安装器已改用 `jsaddons\publish.xml`**；
  上面的 `register-wps.ps1` 仍走 jsplugins，只适用于开发调试

### 功能差异（vs Excel 端）

| 功能 | Excel 端 | WPS 端 |
|------|---------|--------|
| 宿主容器 | COM CustomTaskPane + WebView2 | WPS taskpane（Chromium） |
| 宏执行 | VBA (`execute_vba`) | JSA (`execute_jsa`) |
| 截图 | screenshot_excel | P2 阶段实现 |
| 键盘模拟 | send_keys | P2 阶段实现 |
| 工作簿快照 | create_snapshot / rollback | P2 阶段实现 |
| 副本试跑 | execute_vba 先在隐藏副本上跑 | 无（execute_jsa 直接确认） |
| 先读后写 / 读后被改检测 | ReadLedger | read-ledger.js（同一套规则） |
| 其他工具 | 完整支持 | 完整支持 |

### 真 WPS 待实测（目前只在假对象模型上验证）

开发机没有装 WPS，下面这些调用只用 node 脚本里的假对象模型测过（`scripts/test-wps-*.js`），
下次在装了 WPS 的机器上逐条确认：

- 插删行列按块处理（`Range(...).EntireRow.Insert/Delete`）、`fill_formula_down` 的 `Resize + AutoFill`、
  `clear_range` 的三种类型、`write_table` 的 `ListObjects.Add`、`sort_data` 的 Header 位置参数
- 首次使用：`starter-host.js` 的 `Worksheets.Add(undefined, last)`（新表放到最后）、撇号写文本
- 先读后写：`Application.ApiEvent.AddApiEventListener('SheetChange', …)` 能否收到用户编辑、
  `WorksheetFunction.CountA`
- `execute_jsa` 的确认弹窗在 WPS 面板里是否正常出现、允许 / 拒绝能否回到侧车
- 选区条：`ApiEvent.AddApiEventListener('SheetSelectionChange', …)` 能否收到选区变化、`Selection.Areas.Item(1)` / `CountLarge` 是否可用（取不到时退回 `Count`）

### Excel COM 加载项直接跑在 WPS 上（共用一套代码）：未验证

设想是 WPS 经 `HKCU\Software\Kingsoft\Office\ET\AddIns` + `AddinsWL` 白名单加载 Excel 的
COM 加载项，两个宿主共用一套代码、JS 加载项退居兜底（`scripts/wps-com-test.ps1`、
`scripts/verify-install-sandbox.ps1 -WithWps -ComWps` 都已备好）。2026-09-25 没能得出结论：

- 开发机没有安装 WPS（只有残留的用户数据），不在主力机上装整套办公软件来做实验。
- Windows Sandbox 里 DeepExcel 自己的 COM 激活就失败（`0x80070002`，沙箱没有 Office，
  与 WPS 无关）。这种情况下「WPS 里没有 DeepExcel 选项卡」分不清是 WPS 不支持，还是我们的
  类没激活成功，结论不可信；WPS 首次启动还要扫码登录，需要人在场。
- 一个待确认的线索：`OFFICE.dll` 依赖 `stdole 7.0.3300.0`，安装包没有带它；开发机上的
  stdole 来自 Office 装进 GAC 的那份。没装 Office、只装 WPS 的机器上它很可能缺失。

结论出来之前维持现状：WPS 走 JS 加载项（`jsaddons\publish.xml`）。

## 架构说明

```
WPS 表格 (ET)
  └─ DeepExcel JS 加载项 (main.js)
       ├─ taskpane (Chromium) → React 前端
       │     └─ bridge.ts → window.parent.postMessage
       ├─ sidecar-host.js → Python sidecar (子进程)
       │     └─ stdin/stdout JSON Lines IPC
       ├─ tool-dispatcher.js
       │     └─ wps-actions.js → wps.Application JS API
       │     └─ jsa-executor.js → WPS JSA 宏
       └─ Python sidecar (sidecar.py)
             └─ Claude Agent SDK + MCP Tools
             └─ excel_tools.py → call_csharp (IPC)
```

## 故障排查

### 加载项不显示

先确认你是**怎么装的**，两条路径查的地方不一样：

- 用安装器装的（正式用户走这条）：检查 `%APPDATA%\kingsoft\wps\jsaddons\publish.xml`
  里是否有 `url="DeepExcel_<版本>"` 的节点，且 `jsaddons\DeepExcel_<版本>\` 目录存在
- 用 `register-wps.ps1` 装的（仅开发调试）：检查 `%APPDATA%\kingsoft\wps\jsplugins\jsplugins.xml`

两条都适用的：

1. 检查注册表：`HKCU\Software\kingsoft\office\ET\AddinsWL\DeepExcel` 是否存在
2. WPS 个人版用户：检查信任中心是否已加白名单
3. 重启 WPS 表格

### Python sidecar 启动失败

1. 检查 Python 是否在 PATH 中：`python --version`
2. 检查 `src/DeepExcel.Wps/sidecar/sidecar.py` 是否存在
3. 检查 Claude Agent SDK 是否安装：`pip list | grep claude-agent`
4. 查看诊断日志：WPS 加载项控制台输出（启动参数 `--enable-logging`）

### taskpane 白屏

1. 检查 `src/DeepExcel.Wps/web/index.html` 是否存在
2. 检查 taskpane.html 中的资源路径是否正确
3. 在 taskpane 中按 F12 打开开发者工具查看错误

### 工具调用失败

1. 查看 sidecar 日志（stderr 输出）
2. 确认 WPS 工作簿已打开
3. 确认工具名称正确（参考 `wps-actions.js` 支持的工具列表）

## 卸载

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register-wps.ps1 -Unregister
```

此脚本会移除 WPS 注册表项和发布的加载项文件。

## 开发调试

### 修改前端 UI

1. 编辑 `src/DeepExcel.UI/src/` 下的 React 源码
2. 运行 `build-wps.ps1` 重新构建
3. 重启 WPS 表格查看效果

### 修改工具实现

1. 编辑 `src/DeepExcel.Wps/wps-actions.js`（WPS JS API 实现）
2. 编辑 `src/DeepExcel.Wps/tool-dispatcher.js`（工具调度）
3. 无需重新构建，重启 WPS 即可生效

### 修改 AI 行为

1. 编辑 `src/DeepExcel.Sidecar/system_prompt.py`
2. 编辑 `src/DeepExcel.Sidecar/excel_tools.py`
3. 重启 WPS 表格（sidecar 子进程会重新加载）

## 版本历史

- **0.2.4** (2026-07-09)：WPS 兼容性首版（P0 MVP）
  - WPS JS 加载项骨架
  - Python sidecar 100% 复用
  - 核心 20+ 工具实现
  - JSA 宏执行器
  - 前端 bridge 层多宿主支持
