# DeepExcel

> 在 Microsoft Excel 和 WPS 表格中使用自然语言读取数据、生成公式、调整格式、制作图表并执行分析任务。

DeepExcel 是一款 Windows 表格 AI 加载项。它把模型对话、表格上下文和可审计的工具调用放在同一个侧边面板中，用户不需要在表格与聊天网页之间反复切换。

> [!IMPORTANT]
> 当前测试包使用自签名证书，仅适合受邀内测，不应公开分发。Microsoft Excel 端已完成真实安装和运行验证；WPS 端仍处于兼容性验证阶段，不同 WPS 版本的 JS 加载项能力可能存在差异。

## 快速开始

### 安装内测版

1. 获取 `DeepExcel-Internal-v0.4.17.zip`，删除旧的同名 ZIP 和旧解压目录。
2. 完全退出 Excel、WPS及其后台进程。
3. 解压 ZIP，首次使用时运行 `Install-Internal-Certificate.cmd`。
4. 运行 `DeepExcel.Setup.INTERNAL.exe`。
5. 重新打开 Excel 或 WPS，在功能区选择 **DeepExcel → 打开面板**。

安装器按当前 Windows 用户安装，会：

- 注册 32 位和 64 位 Excel COM 加载项；
- 检查 .NET Framework 4.8 与 WebView2 Runtime；
- 安装内置 Python 3.11 和 AI Sidecar；
- 写入 WPS `publish.xml` 节点，并保留其他 WPS 插件配置。

### 配置模型

1. 打开 DeepExcel 面板，点击右上角的模型设置按钮。
2. 选择模型供应商。
3. 填写 API Key；需要时修改 Base URL。
4. 点击 **刷新模型列表**，从供应商接口获取当前可用模型。
5. 选择模型并点击 **测试连接**。
6. 点击 **保存并应用**。

如果供应商不提供模型枚举接口，DeepExcel 会保留内置模型列表并显示提示，不会清空原有配置。

> [!NOTE]
> API Key 使用 Windows DPAPI 按当前用户加密，`config.json` 不保存明文密钥。模型请求仍会发送给用户选择的第三方模型供应商。

### 执行表格任务

在输入框中直接描述目标，例如：

```text
把 A1:F200 按销售额降序排列，并把前 10 名标成浅绿色。
```

```text
根据 Sheet2 的产品编号，用公式补齐当前表的产品名称。
```

```text
汇总各地区销售额，并生成一张柱状图。
```

DeepExcel 会优先调用结构化表格工具；只有复杂操作才使用 VBA、Python 或 WPS JSA。高风险操作会请求用户确认。

## 主要能力

| 能力 | Microsoft Excel | WPS 表格 |
| --- | --- | --- |
| AI 对话侧边面板 | 支持 | 兼容性验证中 |
| 读写单元格与区域 | 支持 | 支持的 WPS 版本可用 |
| 公式、排序、筛选、格式化 | 支持 | 支持的 WPS 版本可用 |
| 图表与复杂批量操作 | 支持 | 部分能力使用 JSA |
| VBA | 支持，需要 VBA 工程访问权限 | 不使用 VBA |
| 操作前快照 | 支持 | 能力受 WPS 宿主限制 |
| 多工作簿对话隔离 | 支持 | 逐步适配 |
| 模型列表刷新 | 支持 | 逐步适配 |

其他功能包括流式回复、工具调用状态、操作快照、手动回滚、对话历史、附件、提示词模板、VBA 中文字符串处理和 Sidecar 冷启动诊断。

## 系统要求

| 项目 | 要求 |
| --- | --- |
| 操作系统 | Windows 10 / Windows 11，x64 |
| Microsoft Office | Excel 2016 或更高版本，支持 32 位或 64 位 Office |
| WPS | 支持 `publish.xml` 和 JS 加载项的 WPS 表格版本 |
| .NET Framework | 4.8 或更高版本，Excel 端需要 |
| WebView2 Runtime | Excel 端需要，安装器会检测 |
| 网络 | 调用模型 API 时需要 |
| Python | 无需用户安装，安装包内置 Python 3.11 |

## 模型供应商

项目预置 Anthropic、DeepSeek、OpenAI、Kimi、通义千问、智谱、MiniMax、豆包和阶跃星辰等配置，并支持自定义兼容端点。

预置模型名称是发布时的默认目录，不代表供应商实时可用性。建议安装后使用 **刷新模型列表** 获取当前账号可访问的模型；最终可用模型、计费和限额以供应商返回结果为准。

## 安全与数据

| 路径 | 内容 |
| --- | --- |
| `%APPDATA%\DeepExcel\config.json` | 模型、界面和运行配置，不含明文 API Key |
| `%APPDATA%\DeepExcel\credentials\` | DPAPI 加密的 API Key |
| `%APPDATA%\DeepExcel\logs\` | 加载项和 Sidecar 日志 |
| `%LOCALAPPDATA%\DeepExcel\Snapshots\` | 操作快照 |
| `%LOCALAPPDATA%\DeepExcel\history\` | 工作簿对话历史 |
| `%LOCALAPPDATA%\DeepExcel\Attachments\` | 会话附件 |
| `%APPDATA%\kingsoft\wps\jsaddons\` | WPS 加载项文件和 `publish.xml` |

安全边界：

- API Key 只允许当前 Windows 用户通过 DPAPI 解密；
- VBA 和 Python 执行经过权限确认与受限操作检查；
- WPS 清单更新只修改 DeepExcel 节点；
- 内测自签名证书不等同于公共 CA 代码签名；
- 使用 AI 修改重要工作簿前，仍建议保留独立备份。

## 项目结构

```text
DeepExcel/
├── src/
│   ├── DeepExcel.AddIn/       # Excel COM 加载项、WebView2、表格工具
│   ├── DeepExcel.UI/          # React + TypeScript 侧边面板
│   ├── DeepExcel.Sidecar/     # Python AI Agent 与 IPC
│   ├── DeepExcel.Wps/         # WPS Ribbon、任务窗格与 JSA 工具
│   └── DeepExcel.Tests/       # 单元测试
├── scripts/                   # 编译、注册、签名和打包脚本
├── deploy/                    # Inno Setup 安装器定义
├── docs/                      # 部署、设计与审计文档
└── DeepExcel.sln
```

运行链路：

```text
用户输入
  → React 任务窗格
  → Excel C# Bridge / WPS JS Bridge
  → Python Sidecar
  → 模型供应商 API
  → 结构化工具调用
  → Excel 或 WPS 表格对象模型
```

## 本地开发

### 构建前端

```powershell
cd src\DeepExcel.UI
npm install
npm run build
```

### 编译并注册 Excel 加载项

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\_compile_only.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\register-user.ps1
```

运行 `_compile_only.ps1` 前应关闭 Excel，避免开发 DLL 被正在运行的 Excel 进程锁定。

### 构建 WPS 资源

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build-wps.ps1
```

该命令会构建共享前端、复制 WPS Sidecar，并运行 Ribbon/任务窗格回调冒烟测试。

### 构建内置 Python

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\package-python.ps1
```

Python、pip 和依赖版本由脚本及 `scripts/python-requirements.lock.txt` 固定，并在构建时校验哈希。

## 构建安装包

### 内部自签名测试包

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\new-internal-signing-cert.ps1
python scripts\package_release.py --version 0.4.17 --internal
```

输出：

```text
dist\DeepExcel-Internal-v0.4.17.zip
dist\DeepExcel.Setup.INTERNAL.exe
```

### 生产签名包

生产构建需要有效的 Authenticode PFX 或 Azure Trusted Signing：

```powershell
$env:DEEPEXCEL_PFX = 'C:\secure\deepexcel-code-signing.pfx'
$env:DEEPEXCEL_PFX_PASS = '<由安全环境提供>'
python scripts\package_release.py --version 0.4.17
```

没有生产证书时，只允许生成本机验证版本：

```powershell
python scripts\package_release.py --version 0.4.17 --allow-unsigned
```

无签名产物禁止发送给用户。完整流程见 [部署文档](docs/DEPLOYMENT.md)。

## 发布前验证

至少完成：

1. C# 加载项编译；
2. React 前端构建；
3. VBA 与模型列表解析测试；
4. Sidecar 自然冷启动导入；
5. Excel 32 位和 64 位 COM 激活检查；
6. WPS Ribbon/任务窗格回调测试；
7. 安装、升级、卸载及开发注册恢复；
8. 安装包 Authenticode 与 SHA-256 校验。

## 常见问题

### Excel 中没有 DeepExcel 选项卡

1. 完全退出所有 Excel 进程；
2. 重新运行安装器；
3. 在 `文件 → 选项 → 加载项` 中检查“禁用项目”和“COM 加载项”；
4. 确认 `DeepExcel.AddIn` 已启用；
5. 查看 `%APPDATA%\DeepExcel\logs\DeepExcel_Load.log`。

### Sidecar 显示 `code=1`

最新版本会附带启动错误摘要。请查看 `%APPDATA%\DeepExcel\logs\`，并确认以下文件存在：

```text
%LOCALAPPDATA%\DeepExcel\python\python.exe
%LOCALAPPDATA%\DeepExcel\sidecar\sidecar.py
```

### VBA 执行失败

- 入口必须是无参数 `Sub`；
- 过程名和变量名使用英文、数字和下划线；
- 中文可以出现在字符串和注释中；
- 不要使用 `MsgBox`、`InputBox` 或独立 `End`；
- 如果提示 VBA 工程未授权，只需启用“信任对 VBA 工程对象模型的访问”，不建议启用所有宏。

### WPS 没有图标或任务窗格打不开

1. 确认使用最新安装包并完全重启 WPS；
2. 检查 `%APPDATA%\kingsoft\wps\jsaddons\publish.xml` 是否包含 DeepExcel；
3. 检查同目录下是否存在 `DeepExcel_<版本>` 文件夹；
4. 按 `Alt+F12` 打开 WPS 加载项调试器，查看第一条红色 Console 错误；
5. 将 WPS 版本号、错误截图和 Console 错误一起反馈。

### 如何卸载

在 Windows“已安装的应用”中卸载 DeepExcel。卸载器会移除程序文件、Excel 注册和 WPS DeepExcel 节点，但会保留用户配置、日志、快照和对话历史。

如需彻底清除个人数据，请确认不再需要后手动删除：

```text
%APPDATA%\DeepExcel\
%LOCALAPPDATA%\DeepExcel\
```

## 文档

- [部署与签名](docs/DEPLOYMENT.md)
- [WPS 说明](docs/README-WPS.md)
- [架构设计](docs/superpowers/specs/2026-06-25-DeepExcel-architecture-design.md)
- [经验总结](docs/lessons-learned.md)
