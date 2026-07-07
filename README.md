# DeepExcel

> Excel AI 智能助手 — 让 AI 帮你处理表格数据、生成图表、写公式、做数据分析。

DeepExcel 是一款 Excel COM 加载项（Add-In），在 Excel 中直接集成 AI 对话能力。选中数据范围，告诉 AI 你想做什么，它会自动操作表格，无需切换窗口。

---

## ✨ 功能特性

- **💬 自然语言操作表格**：用中文描述需求，AI 自动读写单元格、设置格式、生成公式
- **📊 智能图表**：一句话生成柱状图、折线图、饼图、甘特图等多种图表
- **🧹 数据清洗**：自动去重、填充缺失值、格式转换、拆分/合并列
- **🔢 公式助手**：根据需求生成 Excel 公式，解释已有公式
- **📋 多工作簿隔离**：每个工作簿有独立的 AI 对话上下文，互不干扰
- **💾 操作快照**：AI 修改前自动创建快照，随时一键回滚
- **🔐 安全可控**：所有 API Key 使用 DPAPI 加密本地存储，不上传第三方服务器
- **🌓 亮/暗主题**：支持浅色/深色主题切换
- **📎 附件支持**：支持上传图片、文档等附件，AI 可识别图片内容

---

## 🖥️ 系统要求

| 项目 | 要求 |
|------|------|
| 操作系统 | Windows 10 / Windows 11（x64） |
| Excel 版本 | Excel 2016 或更高版本（推荐 Office 365 / 2021） |
| .NET Framework | 4.8 或更高版本（Win10/11 通常已预装） |
| WebView2 Runtime | Win11 自带；Win10 需[手动安装](https://developer.microsoft.com/microsoft-edge/webview2/) |
| Python | 3.8 或更高版本（用于 AI sidecar，安装包已内置） |
| 网络连接 | 需要（用于调用 AI 模型 API） |

---

## 📥 下载与安装

### 第一步：下载安装包

从 [GitHub Releases](https://github.com/qinjuanchen/DeepExcel/releases) 下载最新版本的 ZIP 包（如 `DeepExcel-v0.2.1.zip`）。

### 第二步：解压

将 ZIP 包解压到任意目录，例如：

```
C:\Program Files\DeepExcel\
```

> 建议解压到非系统盘或用户目录，避免权限问题。

### 第三步：注册加载项

1. 进入解压后的目录，找到 `register-user.ps1` 文件
2. **右键 → 使用 PowerShell 运行**，或在 PowerShell 中执行：

```powershell
cd "解压后的目录路径"
powershell -ExecutionPolicy Bypass -File register-user.ps1
```

3. 看到 `Registration successful!` 即表示注册成功

> 💡 **无需管理员权限**：`register-user.ps1` 使用 HKCU 注册表项，只为当前用户注册。

### 第四步：启动 Excel

打开 Excel，在顶部功能区找到 **DeepExcel** 选项卡，点击 **打开面板** 按钮即可开始使用。

---

## 🚀 快速开始

### 1. 配置 AI 模型

首次使用需要配置 API Key：

1. 点击 Excel 功能区的 **模型配置** 按钮
2. 在左侧选择你的模型提供商（Anthropic / DeepSeek / 智谱 / 通义千问 / Kimi / 豆包 / Stepfun / OpenAI / Minimax / 自定义）
3. 输入 API Key 和 Base URL（部分提供商已预置默认地址）
4. 选择模型版本
5. 点击 **测试连接** 验证配置
6. 点击 **保存并应用**

> 🔒 **安全提示**：API Key 使用 Windows DPAPI 加密存储在本地，仅当前用户可解密，不会明文保存。

### 2. 开始对话

1. 在 Excel 中打开或创建一个工作簿
2. 选中你想操作的数据范围（可选）
3. 点击 **DeepExcel → 打开面板**
4. 在底部输入框输入你的需求，例如：
   - "把销售额按降序排列"
   - "生成一张各地区销售额对比的柱状图"
   - "用 VLOOKUP 把 Sheet2 的产品名称匹配过来"
   - "删除重复行，缺失值用 0 填充"
5. 按回车发送，AI 会自动操作表格

### 3. 回滚操作

如果 AI 修改的结果不满意：

- 点击面板顶部的 **📷 快照** 按钮查看历史快照
- 选择对应快照点击 **回滚** 即可恢复

> ⚠️ AI 每次操作前会自动创建快照，但建议重要数据先手动备份。

---

## 🤖 支持的模型提供商

| 提供商 | 预置模型 | 支持视觉 | 备注 |
|--------|----------|----------|------|
| Anthropic (Claude) | claude-sonnet-5 / claude-opus-4.8 / claude-haiku-5 | ✅ | 官方接口 |
| DeepSeek | deepseek-v4-pro / deepseek-v4-flash | ❌ | Anthropic 兼容 |
| 阶跃星辰 Stepfun | step-3.7-flash / step-3.5-flash | ✅ | 视觉能力强 |
| OpenAI | gpt-5.5 / gpt-5.5-pro / gpt-5 | ✅ | 官方接口 |
| Kimi (月之暗面) | kimi-k2.7-code / kimi-k2.6 / kimi-k2-thinking | ✅ | Anthropic 兼容 |
| 通义千问 (阿里) | qwen3.7-max / qwen3-max / qwen3-coder-plus | ✅ | Anthropic 兼容 |
| 智谱 GLM | glm-5.2 / glm-5.1 / glm-4.7-flash | ✅ | Anthropic 兼容 |
| Minimax | MiniMax-M2.5 / MiniMax-M2 | ❌ | Anthropic 兼容 |
| 豆包 (火山引擎) | doubao-seed-2.1-pro / doubao-seed-2.1 / doubao-seed-1.6 | ✅ | Anthropic 兼容 |
| 自定义 | 自定义模型 | ❌ | OpenAI 兼容格式 |

> 💡 上传图片附件时，如果当前模型不支持视觉，会自动切换到支持视觉的模型。

---

## ⚙️ 配置与数据位置

所有配置和数据都存储在用户目录下，卸载时不会残留系统文件：

```
%APPDATA%\DeepExcel\
├── config.json              # 配置文件（不含 API Key 明文）
├── credentials\             # 加密的 API Key（DPAPI）
│   └── key_{provider}.crypt
├── logs\                    # 运行日志
│   └── deepexcel-YYYYMMDD.log
├── snapshots\               # 操作快照（用于回滚）
└── conversations\           # 对话历史
```

---

## 🔧 常见问题

### Q: Excel 里找不到 DeepExcel 选项卡？

**A:** 尝试以下方法：

1. 重启 Excel
2. 检查是否被禁用：`文件 → 选项 → 加载项 → 管理: 禁用项目 → 转到 → 取消禁用 DeepExcel.AddIn`
3. 重新运行 `register-user.ps1`
4. 确认 .NET Framework 4.8 已安装

### Q: 面板打开是空白/白屏？

**A:** 安装 [WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/)。Win11 已自带，Win10 可能需要手动安装。

### Q: AI 回复很慢或一直转圈？

**A:**

1. 检查网络连接和代理设置
2. 确认 API Key 正确（点击 **模型配置 → 测试连接**）
3. 查看日志文件 `%APPDATA%\DeepExcel\logs\` 获取详细错误信息

### Q: 如何卸载？

**A:**

1. 运行 `register-user.ps1 -Unregister` 注销加载项
2. 删除解压的文件夹
3. （可选）删除 `%APPDATA%\DeepExcel\` 目录清除所有配置和数据

### Q: API Key 安全吗？

**A:** 安全。API Key 使用 Windows DPAPI（Data Protection API）加密存储，加密密钥绑定当前 Windows 用户账户，只有登录用户才能解密。配置文件中不保存明文密钥。

---

## 🛠️ 开发指南

### 项目结构

```
DeepExcel/
├── src/
│   ├── DeepExcel.AddIn/        # C# Excel 加载项（COM 互操作 + WebView2）
│   ├── DeepExcel.Sidecar/      # Python sidecar（AI 对话 + 工具调用）
│   ├── DeepExcel.UI/           # React + TypeScript 前端界面
│   └── DeepExcel.Tests/        # C# 单元测试
├── scripts/                    # 构建、打包、注册脚本
├── docs/                       # 设计文档
├── test-data/                  # 测试数据
└── DeepExcel.sln               # Visual Studio 解决方案
```

### 本地开发

```powershell
# 1. 前端依赖安装与开发
cd src\DeepExcel.UI
npm install
npm run dev

# 2. C# 编译
cd ..\..\scripts
powershell -ExecutionPolicy Bypass -File _compile_only.ps1
powershell -ExecutionPolicy Bypass -File register-user.ps1

# 3. 运行 Python sidecar（自动启动，无需手动运行）
```

### 构建发布包

```powershell
python scripts\package_release.py --version 0.2.1
```

输出目录：`dist\DeepExcel-v0.2.1.zip`

---

## 📄 许可证

MIT License

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

**Enjoy using DeepExcel! 🎉**
