# DeepExcel

> 在 Microsoft Excel 和 WPS 表格中使用自然语言读取数据、生成公式、调整格式、制作图表并执行分析任务。

DeepExcel 是一款 Windows 表格 AI 加载项。它把模型对话、表格上下文和可审计的工具调用放在同一个侧边面板中，用户不需要在表格与聊天网页之间反复切换。

> [!IMPORTANT]
> 当前安装包**未进行代码签名**，首次运行会触发 Windows SmartScreen 提示，属于正常现象（见下）。Microsoft Excel 端已完成真实安装和运行验证；WPS 端仍处于兼容性验证阶段，不同 WPS 版本的 JS 加载项能力可能存在差异。

## 快速开始

### 安装

1. 完全退出 Excel、WPS 及其后台进程。
2. 运行 `DeepExcel.Setup.exe`。出现「Windows 已保护你的电脑」时，点击 **更多信息 → 仍要运行**。
3. 重新打开 Excel 或 WPS，在功能区选择 **DeepExcel → 打开面板**。

安装完成后如果功能区没有 DeepExcel 选项卡，运行安装目录下的 `DeepExcel.Repair.exe`。

<details>
<summary>关于 SmartScreen 提示与完整性校验</summary>

DeepExcel 尚未购买代码签名证书，因此 Windows 无法显示发布者信息。安装包完整性通过 SHA-256 公布，可在安装前自行核对：

```powershell
certutil -hashfile DeepExcel.Setup.exe SHA256
```

把输出与下载页的 `SHA256SUMS.txt` 比对，一致即说明文件未被篡改。

早期内测版曾要求运行 `Install-Internal-Certificate.cmd` 安装自签名根证书，该做法已于 2026-09-12 移除——让用户信任一个私有根 CA 的风险高于 SmartScreen 提示带来的不便。如果你装过旧内测版，建议清理残留证书：

```powershell
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -eq 'CN=DeepExcel Internal Testing' } | Remove-Item
```

</details>

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

### 工作簿结构摘要

打开工作簿时会在后台构建轻量结构摘要——每列的名称、类型、取值范围、空值位置、公式列以及跨表引用——随每条消息提供给模型。这让模型不必为了搞清"C 列是不是数字"而反复读取数据；大表上它本来也读不全。摘要会在编辑后自动失效重建，采样得出的结论会明确标注为采样结论。

### 执行前变更预览

删行、批量写入、清空区域等操作会先算出**具体会改哪些单元格**，以改前/改后对照展示，确认后才执行。覆盖已有公式的单元格会单独标出。

VBA 和 Python 无法预演，这种情况会**明确说明"无法预览"并自动创建快照**，不会假装模拟过。改动很小或根本没有变化的操作会直接执行，不打扰。

### 技能库

一次成功的多步任务可以保存为技能：系统会把其中的数据区域、日期、工作表名识别成参数，下次填入新值即可重放。技能是骨架而非宏——重放仍由 AI 执行，数据结构变化时以原始意图为准，并且同样经过上述确认与预览流程。

技能按单文件存储在 `%APPDATA%\DeepExcel\skills\`，可直接发送文件分享。登录账号后还可以同步到云端、用分享码分享给他人。

**同步与分享是分开的**：同步只是备份，技能仍然私有；只有显式点“分享”才会生成分享码，且随时可以撤销。上传前会自动移除本地文件路径与工作簿名——技能的参数默认值来自真实运行，里面可能有真实的文件路径。客户端和服务端各脱敏一次。

### 账号（可选）

连接 DeepExcel 服务器后可使用账号登录，用于内测准入与使用统计。**不登录也能完整使用**——在模型设置里填写自己的 API Key 即可。

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
- 安装包未经代码签名，请用 `SHA256SUMS.txt` 核对完整性；
- 使用 AI 修改重要工作簿前，仍建议保留独立备份。

## 项目结构

```text
DeepExcel/
├── src/
│   ├── DeepExcel.AddIn/       # Excel COM 加载项、WebView2、表格工具
│   │   ├── Account/           # 账号会话、出口路由解析、遥测
│   │   ├── Perception/        # 工作簿语义索引（类型推断、采样、缓存）
│   │   ├── Preview/           # 执行前变更预览
│   │   └── Skills/            # 技能参数化与本地技能库
│   ├── DeepExcel.UI/          # React + TypeScript 侧边面板
│   ├── DeepExcel.Sidecar/     # Python AI Agent 与 IPC
│   ├── DeepExcel.Wps/         # WPS Ribbon、任务窗格与 JSA 工具
│   ├── DeepExcel.Repair/      # 诊断与修复工具（GUI + CLI）
│   ├── DeepExcel.Probe32/     # 32 位 COM 激活探针
│   └── DeepExcel.Tests/       # 单元测试
├── server/                    # 账号 / 权益 / 出口路由 / 遥测服务端
├── scripts/                   # 编译、测试、打包与安装验收脚本
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

```powershell
python scripts\package_release.py --version 0.5.0
```

输出：

```text
dist\DeepExcel.Setup.exe        # 面向用户的唯一交付物
dist\DeepExcel-v0.5.0.zip      # 支持/诊断 payload
dist\SHA256SUMS.txt             # 必须与安装包一起发布
dist\update.json                # 签名更新清单，放到更新源上
```

当前发布形态是**未签名安装包 + 公布 SHA-256**。签名链路完整保留：设置 `DEEPEXCEL_PFX`、`DEEPEXCEL_CERT_THUMBPRINT` 或 `USE_AZURE_TRUSTED_SIGNING` 任一即可自动签名，无需改代码。自签名证书会被打包脚本主动拒绝。

完整流程见 [部署文档](docs/DEPLOYMENT.md)。

## 自动更新

**SmartScreen 只拦从浏览器下载、带"来源标记"（MOTW）的文件。**由已安装程序自己下载的更新包没有这个标记，所以首次安装痛一次，之后每个版本都静默完成——这是不买代码签名证书时摩擦降幅最大的一项。

安全模型只有一句话：**服务端永远拿不到私钥，所以拿下服务端也推不了更新。**

```
私钥（离线保管）      签名 update.json
   │
   └─ 公钥 → 编译进客户端 → 客户端验签 → 校验 SHA-256 → 执行

服务端只是原样转发已签名的清单，不签名、不持有私钥
```

### 一次性准备：生成签名密钥

```bash
python scripts/update_signing.py genkey --out D:/offline/deepexcel-update.pem
```

它会把公钥写进 `src/DeepExcel.AddIn/Updates/UpdateSigning.cs`，因此公钥和私钥不可能被手工改到不一致。**私钥生成后立刻挪到离线存储**：拿到它的人可以向所有已安装客户端推送并执行任意代码，而收回的唯一办法是发一个换了公钥的新客户端。脚本拒绝把私钥写进仓库内，`.gitignore` 再兜一层。

改动公钥后必须重新编译并发布客户端——在此之前构建的客户端会拒绝新密钥签的清单。

### 每次发版

```bash
export DEEPEXCEL_UPDATE_KEY=D:/offline/deepexcel-update.pem
export DEEPEXCEL_UPDATE_BASE_URL=https://github.com/<owner>/<repo>/releases/download/v0.6.0
python scripts/package_release.py --version 0.6.0
```

打包脚本会用客户端里那份公钥验证刚生成的 `update.json`，**签名密钥与客户端公钥不匹配时直接构建失败**。这条守卫的价值在于：不匹配不会有任何症状，所有客户端只会悄悄停止升级，看起来和"还没人升级"一模一样。

可选：`DEEPEXCEL_UPDATE_CHANNEL`（默认 `stable`）、`DEEPEXCEL_UPDATE_NOTES`、`DEEPEXCEL_UPDATE_MINIMUM`（低于该版本的客户端不自动升级，提示手动重装）。

发布时把 `DeepExcel.Setup.exe` 放到 `DEEPEXCEL_UPDATE_BASE_URL` 下，把 `update.json` 放到更新源上。

### 服务端：更新源

放文件即发布，无需重启：

```bash
UPDATE_MANIFEST_DIR=/srv/deepexcel/updates
cp dist/update.json /srv/deepexcel/updates/stable.json
```

`GET /api/v1/updates/latest?channel=stable` 原样返回该文件（**刻意不鉴权**：登录过期或从未登录的用户同样需要拿到修复，清单的可信度来自签名而不是来自谁在请求）。

### 客户端：指向更新源

`%APPDATA%\DeepExcel\config.json`：

```json
{ "Update": { "Enabled": true, "FeedUrl": "https://api.example.com/api/v1/updates/latest", "Channel": "stable" } }
```

`FeedUrl` 默认为空，公钥默认也为空——**没人配置过的构建不会联系任何地址**，不存在一个写死的默认域名可以被抢注。

流程：Excel 启动 1 分钟后开始后台检查、之后每 6 小时一次（失败静默）→ 下载并验签、校验 SHA-256 → 面板出现一行「新版本已就绪」→ 用户点「重启安装」→ Excel 按正常流程关闭（未保存内容照常提示，用户取消也没关系，更新留到下次）→ `DeepExcel.Updater.exe` 重新验一遍签名和摘要 → 静默安装 → 重新打开 Excel。

**为什么不是只在启动时查一次：** Excel 常常一开就是几天。只查一次意味着早上那次网络不好，这一周都拿不到更新；中午发的版本，没重启过的人永远收不到。

被拒绝的清单一律不安装，且每种拒绝都有独立原因：签名不符、未知密钥、版本不更新（防降级）、通道不符、非 https 地址、摘要或大小不符、低于最低可升级版本。

### 装不上时会怎样

同一个版本连续 3 次启动安装都没成功，就不再自动提示，面板改成黄色的「多次安装未成功，请手动下载安装」。

这不是保守，是必须：更新包还在、验签也过，所以每次重开 Excel 都会重新判定为「已就绪」并再提示一次。国内最常见的原因是 360 / 火绒直接拦掉安装器，而且**用户完全看不到发生了什么**——没有计数的话，这就是一个永远点不完的弹窗。计数记在启动更新程序的那一刻，不是等它回报，因为那时候 Excel 已经退出，没人还能听到失败。

### 可观测性

更新是唯一从内部看不见的功能：它在后台跑，而知道「装成功了」的那个进程装完就退出了。所以更新程序在安装成功后往暂存根目录写一份回执，新版本启动时读取、上报、删除——**这是「到底有多少人真升上去了」的唯一来源**。

上报的事件是 `update_event`，字段只有：阶段（check/download/launch/apply）、结果（up_to_date/ready/failed/blocked/installed/started）、分类码、来源与目标版本号、耗时。

**刻意不包含更新源地址、暂存路径和异常原文。**这三样是把用户数据夹带进遥测的常见途径，而这个模块的每条错误消息都至少含其中之一——所以分类码在抛异常的那一行就定好，不是事后从消息里猜。服务端白名单是第二道保证，客户端漏传也进不了库。

未登录时更新照常工作（更新源本来就不鉴权），只是统计不到。

## 发布前验证

至少完成：

1. C# 加载项编译（`scripts\_compile_only.ps1`，需先关闭 Excel/WPS）；
2. 原生工具编译（`scripts\build-repair.ps1`：诊断修复 + 32 位探针 + 更新程序）；
3. React 前端构建；
4. C# 单元测试（`scripts\run-tests-csharp.ps1`）；
5. 打包守卫测试（`python scripts\test_package_release.py`）；
6. 更新链路端到端测试（`python scripts\test_updater_e2e.py`）；
7. Sidecar 自然冷启动导入；
8. WPS Ribbon/任务窗格回调测试；
9. **全新机器安装验收**（`scripts\verify-install-sandbox.ps1 -Launch`）——覆盖 SHA-256 校验、静默安装、32/64 位 COM 注册与激活、卸载清理；
10. `SHA256SUMS.txt` 与安装包一同发布；`update.json` 发布到更新源。

第 9 步是硬性要求。开发机永远装得上——它已经有注册表项、有运行时、没有下载来源标记。v0.4.11 → v0.4.17 连续七个版本栽在这里，就是因为没有在干净机器上验证过。

## 常见问题

### Excel 中没有 DeepExcel 选项卡

1. 完全退出所有 Excel 进程；
2. 运行安装目录（`%LOCALAPPDATA%\DeepExcel`）下的 **`DeepExcel.Repair.exe`**，或开始菜单的「DeepExcel 诊断与修复」；
3. 窗口会自动诊断并给出结论。若显示「发现可自动修复的问题」，点击 **修复**；
4. 重新打开 Excel。

如果提示存在无法自动修复的问题，点击 **导出诊断包**，把桌面上生成的 ZIP 发给支持人员。该 ZIP 只含日志、环境信息和脱敏后的配置，不含单元格内容、工作簿路径或 API Key。

命令行用法：

```powershell
%LOCALAPPDATA%\DeepExcel\DeepExcel.Repair.exe --verify
```

常见诊断码：

| 诊断码 | 含义 |
| --- | --- |
| `E-REG-002` | 32 位视图缺少 COM 注册，32 位 Excel 无法加载 |
| `E-REG-005` | 存在机器级（HKLM）注册残留，会覆盖用户级注册，需管理员清理 |
| `E-REG-006` | 注册指向旧版本 DLL |
| `E-ACT-001` / `E-ACT-002` | 64 位 / 32 位进程中 COM 激活失败 |
| `E-ENV-001` / `E-ENV-002` | 缺少 .NET Framework 4.8 / WebView2 运行时 |
| `E-RES-001` | 加载项被 Excel 加入禁用列表 |
| `E-LOAD-001` | 上次 Excel 启动时加载项初始化失败（附具体异常） |

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

## 运行测试

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run-tests-csharp.ps1
python scripts\test_package_release.py
python scripts\test_installer_script.py
python scripts\test_sidecar_routing.py
python scripts\test_sidecar_permissions.py
python scripts\test_updater_e2e.py
```

`test_updater_e2e.py` 用一次性密钥现编 `DeepExcel.Updater.exe`（编到临时目录，不碰仓库里那份刻意留空的公钥），再逐条验证被篡改的包、被伪造的签名、陌生密钥、降级、错通道各自的退出码。

服务端测试：

```powershell
cd server
.venv\Scripts\python.exe -m pytest tests\ -q
```

本项目没有 dotnet CLI，C# 测试通过 Roslyn + xunit.runner.console 直接编译运行。首次运行需要还原引用程序集：

```powershell
.\nuget.exe install Microsoft.NETFramework.ReferenceAssemblies.net48 -Version 1.0.3 -OutputDirectory packages
```

## 文档

- [服务端（账号 / 权益 / 出口路由 / 遥测）](server/README.md)
- [部署与签名](docs/DEPLOYMENT.md)
- [WPS 说明](docs/README-WPS.md)
- [架构设计](docs/superpowers/specs/2026-06-25-DeepExcel-architecture-design.md)
- [经验总结](docs/lessons-learned.md)
