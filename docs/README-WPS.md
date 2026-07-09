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
- **oem.ini 限制**：WPS 个人版 12.1.0.16910+ 限制 oem.ini 方式加载，本加载项用 jsplugins.xml 方式不受影响

### 功能差异（vs Excel 端）

| 功能 | Excel 端 | WPS 端 |
|------|---------|--------|
| 宿主容器 | COM CustomTaskPane + WebView2 | WPS taskpane（Chromium） |
| 宏执行 | VBA (`execute_vba`) | JSA (`execute_jsa`) |
| 截图 | screenshot_excel | P2 阶段实现 |
| 键盘模拟 | send_keys | P2 阶段实现 |
| 工作簿快照 | create_snapshot / rollback | P2 阶段实现 |
| 其他工具 | 完整支持 | 完整支持 |

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

1. 检查注册表：`HKCU\Software\kingsoft\office\ET\AddinsWL\DeepExcel` 是否存在
2. 检查 jsplugins.xml 是否在 `%APPDATA%\kingsoft\wps\jsplugins\`
3. WPS 个人版用户：检查信任中心是否已加白名单
4. 重启 WPS 表格

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
