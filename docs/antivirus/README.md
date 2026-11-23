# 杀软误报：自查与申诉

国内用户装不上 DeepExcel，最常见的拦路虎不是 SmartScreen，而是 360 / 火绒 / 腾讯电脑管家
把未签名安装包或其中某个文件报成可疑。这个目录放三样东西：

| 文件 | 用途 |
| --- | --- |
| [hash-lookup-template.md](hash-lookup-template.md) | 每查一次就复制一份的「按 SHA-256 查询结果」记录模板 |
| [detections.md](detections.md) | 每个发布版本一行的检出记录总表 |
| [appeal-drafts.md](appeal-drafts.md) | 给 360、火绒、腾讯、Microsoft WDSI 的申诉材料草稿，外加一份各家共用的软件说明 |

## 红线

1. **绝不把安装包上传到 VirusTotal 一类公开多引擎服务。** 上传等于公开发布该文件，任何付费
   用户都能从那里下载。按 SHA-256 **查询**不涉及上传，可以做——但见第 3 条。
2. **向厂商提交申诉、提交样本是对外动作，每一次都要产品负责人点头。** 本目录里的申诉材料
   全是草稿，没有发出过任何一份。提交样本给厂商（WDSI、360 开放平台等）是把文件交给该厂商，
   不是公开发布，但仍然要他决定。
3. 按哈希去网上查检出结果也是联网动作，由他本人或他明确授权的会话来做；查完把结果填进模板。
4. 申诉材料里不写服务器 IP、密钥、内部账号；联系人和邮箱由他提交时填写（仓库是公开的）。

## 每次发版怎么做

1. 发版产物出来后，从 `dist/SHA256SUMS.txt` 抄下 `DeepExcel.Setup.exe` 的 SHA-256 和字节数，
   在 [detections.md](detections.md) 加一行，状态写「未查」。
2. 按哈希查询（见红线第 3 条），每个来源复制一份 [hash-lookup-template.md](hash-lookup-template.md)
   记录到 `lookups/<版本>-<日期>.md`，然后把汇总结论回填到总表。
3. 有检出：按 [appeal-drafts.md](appeal-drafts.md) 准备对应厂商的申诉，**交他确认后**再提交；
   提交时间、工单号回填总表。
4. 每次发版都要重查一遍：未签名文件每个版本的哈希都不同，上一版申诉过的白名单不会自动覆盖下一版。
   这也是买代码签名证书最实际的收益——签名后多数厂商按证书信任，不再逐个哈希申诉（2026-09-12
   定的是暂不买证书，重启条件见记忆 / 决策记录）。

## 为什么容易被报

不是猜测，下面每一条都对过代码，也是申诉材料里需要主动说明的点：

- 安装包和 DeepExcel 自己的四个 EXE（`DeepExcel.Setup.exe`、`DeepExcel.Repair.exe`、
  `DeepExcel.Probe32.exe`、`DeepExcel.Updater.exe`）**都没有代码签名**；随包的第三方文件
  `claude.exe`（Anthropic）、`python.exe`（Python Software Foundation）、WebView2 引导程序（Microsoft）
  签名有效。
- 安装器是 Inno Setup，装到 `%LOCALAPPDATA%\DeepExcel`，写当前用户的 COM 注册（`HKCU\Software\Classes`
  的 CLSID / ProgID，32 / 64 两个视图）、Excel 加载项键、Excel 的 Resiliency 白名单，以及 WPS 的
  `ET\AddIns` / `AddinsWL`，并从 WPS 的禁用名单 `AddinsCL` 删掉自己。
- 加载项运行时临时把 `HKCU\Software\Microsoft\Office\<版本>\Excel\Security\AccessVBOM` 设为 1
  （运行 VBA 需要），关闭时恢复原值——「修改 Office 安全设置」是常见的启发式规则。
- 用户确认后可以往工作簿里写入并执行 VBA、在子进程里执行 Python、模拟按键（`send_keys`）、
  截取 Excel 窗口——都是高风险能力，代码里逐次确认，但静态特征上像宏注入 / 键盘模拟。
- 内置 Python 3.11 运行时和约 220 MB 的 `claude.exe`，运行时会拉起子进程并联网（模型接口、账号服务）。
- 自动更新器会下载安装包并以 `/VERYSILENT` 静默运行——「下载器 + 静默安装」是典型特征。
  **目前发出的版本里自动更新是关着的**（内置更新签名公钥为空），但这段代码在包里。
