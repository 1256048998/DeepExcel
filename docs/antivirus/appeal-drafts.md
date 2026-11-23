# 误报申诉材料（草稿，一份都没提交过）

**每一份提交前都要产品负责人点头。** 提交入口、要不要附样本、用哪个联系人，都是他的决定。
下面的「提交入口」是我按印象写的渠道类型，**没有上网核对过**，提交时以厂商当时的页面为准。

尖括号 `<…>` 是提交时要填的；其余文字可以直接用。

---

## 共用：软件说明（各家都附这一段）

**软件名称：** DeepExcel
**发布者：** DeepExcel（个人开发者，暂无企业主体 / 代码签名证书）
**源码：** https://github.com/1256048998/DeepExcel （公开仓库）
**用途：** Microsoft Excel 与 WPS 表格的 AI 助手加载项。用户在 Excel 侧栏里用中文描述需求，
由 AI 模型生成并在用户确认后执行表格操作（读写单元格、公式、图表、数据清洗、VBA 宏等）。

**安装行为（逐条可在源码 `deploy/DeepExcel.Setup.iss` 核对）：**
- Inno Setup 安装包，按当前用户安装到 `%LOCALAPPDATA%\DeepExcel`，正常情况下不需要管理员权限；
  只有检测到旧版本残留的机器级注册时，才经用户同意弹一次 UAC，仅删除 DeepExcel 自己的旧注册项。
- 在 `HKCU\Software\Classes` 下注册 DeepExcel 自己的 COM 类（32 / 64 位两个视图），写入 Excel 的
  COM 加载项键（`HKCU\Software\Microsoft\Office\Excel\Addins\DeepExcel.AddIn`）和 Excel 的
  `Resiliency\DoNotDisableAddinList`；若装有 WPS，写入 `HKCU\Software\Kingsoft\Office\ET\AddIns`、
  `AddinsWL`，并从 `AddinsCL`（WPS 的已禁用列表）删除 DeepExcel 自己这一项。不修改其他软件的键。
- 缺少 WebView2 运行时时，运行随包附带的、Microsoft 签名的 WebView2 引导安装程序。
- 卸载时删除以上注册和安装的程序文件。

**运行时行为：**
- 作为 Excel / WPS 的 COM 加载项运行，在侧栏显示 WebView2 页面。
- 使用随包的 Python 3.11（Python Software Foundation 签名）运行本地辅助进程，并通过随包的
  `claude.exe`（Anthropic 签名）调用用户自己配置的 AI 模型接口；可选登录账号服务
  `https://deepexcel.657890.xyz`。工作簿内容只发送给用户自己选择的模型服务，DeepExcel 不保存单元格内容、
  用户输入和 AI 回复。
- 运行 VBA 功能需要 Excel 的「信任对 VBA 工程对象模型的访问」：加载项启动时临时把当前用户的
  `AccessVBOM` 设为 1，Excel 关闭时恢复为原来的值；不改动宏警告级别（`VBAWarnings`）。
- 执行 VBA、Python、按键模拟、截取 Excel 窗口这类操作，每一次都先在界面上向用户展示并等待确认。
- 内置自动更新器：从更新源下载经 RSA 签名清单校验的安装包，校验 SHA-256 后静默安装；
  当前发布的版本中该功能未启用。

**文件签名情况：** `DeepExcel.Setup.exe`、`DeepExcel.AddIn.dll`、`DeepExcel.Repair.exe`、
`DeepExcel.Probe32.exe`、`DeepExcel.Updater.exe` 未签名；随包的 `claude.exe`、`python.exe`、
`MicrosoftEdgeWebview2Setup.exe` 为各自厂商的有效签名。

**本次申诉的文件：**

| 文件 | 版本 | SHA-256 | 字节数 | 检出名称 |
| --- | --- | --- | --- | --- |
| DeepExcel.Setup.exe | <v0.x.y> | <sha256> | <字节数> | <对方报的名称，原文> |

**联系人：** <姓名> · <邮箱> （提交时由产品负责人填写）

---

## 360 安全卫士

- **提交入口：** 360 的软件开发者 / 误报申诉渠道（印象中是「360 开放平台」的软件认证与误报申诉，
  未核对）。
- **附件：** 安装包本体（是否上传样本由他定）；上面的「软件说明」；被拦截时的截图。
- **申诉正文：**

  > 您好，我们是表格 AI 加载项 DeepExcel 的开发者。DeepExcel 安装包 `DeepExcel.Setup.exe`
  > （版本 <v0.x.y>，SHA-256 <sha256>）在 360 安全卫士中被识别为 <检出名称>，导致用户无法安装。
  >
  > 该软件为 Excel / WPS 表格的 COM 加载项，安装到当前用户目录、只写入自身的 COM 注册和加载项登记，
  > 源码完整公开于 https://github.com/1256048998/DeepExcel 。我们推测命中的是以下特征：安装包未做
  > 代码签名；加载项为运行 VBA 功能临时开启「信任对 VBA 工程对象模型的访问」（退出时恢复）；包内
  > 带有自动更新组件。详细行为见附件说明。
  >
  > 恳请复核并解除误报。如需样本或进一步信息，请联系 <邮箱>。谢谢！

## 火绒安全

- **提交入口：** 火绒的误报反馈渠道（客户端隔离区 / 官网反馈，未核对）。
- **附件：** 同上。火绒常针对单个文件报，若报的是包内文件（例如 `DeepExcel.Updater.exe`），
  在表格里写包内相对路径和该文件自己的 SHA-256（安装后在 `%LOCALAPPDATA%\DeepExcel` 里算）。
- **申诉正文：**

  > 您好，DeepExcel（Excel / WPS 表格 AI 加载项）的文件 <文件名>（SHA-256 <sha256>）被火绒报为
  > <检出名称>。这是我们自己开发的程序，作用是 <按文件填写：安装程序 / 安装修复工具 / 32 位
  > 注册探针 / 更新程序>，源码公开于 https://github.com/1256048998/DeepExcel 。该文件未做代码签名，
  > 行为说明见附件。请帮忙复核，谢谢。联系方式：<邮箱>。

  按文件填写的作用说明（对过源码）：
  - `DeepExcel.Setup.exe`：Inno Setup 安装程序。
  - `DeepExcel.Repair.exe`：「诊断与修复」工具，检查并修复当前用户的 COM 注册与加载项登记，
    替代了早期版本里调用 powershell.exe 的做法。
  - `DeepExcel.Probe32.exe`：32 位探针，只用于验证 32 位 Excel 能否激活 COM 类。
  - `DeepExcel.Updater.exe`：更新程序，安装前重新校验签名清单与安装包 SHA-256，只在 Excel 关闭后安装。

## 腾讯电脑管家

- **提交入口：** 腾讯电脑管家的软件误报申诉渠道（未核对）。
- **附件 / 正文：** 与 360 那份相同，把「360 安全卫士」换成「腾讯电脑管家」即可。

## Microsoft Defender（WDSI）

- **提交入口：** Microsoft Security Intelligence 的文件提交页（Submit a file for malware analysis），
  身份选「Software developer」，结论选「Incorrectly detected as malware/malicious」。
  这条渠道要求上传文件本体给 Microsoft 分析——不是公开发布，但仍需他决定。
- **申诉正文（英文）：**

  > DeepExcel is an AI assistant add-in for Microsoft Excel and WPS Spreadsheets. Its installer
  > `DeepExcel.Setup.exe` (version <v0.x.y>, SHA-256 <sha256>) is detected as <detection name>.
  >
  > The installer is an Inno Setup package that installs per user into `%LOCALAPPDATA%\DeepExcel`
  > without administrator rights, registers its own COM add-in under `HKCU\Software\Classes` and
  > the Excel add-in key, and removes these registrations and its program files on uninstall. The full source code is public at
  > https://github.com/1256048998/DeepExcel . The add-in temporarily sets the current user's
  > `AccessVBOM` to 1 so it can run VBA macros the user has approved, and restores the original
  > value when Excel closes. It bundles Python 3.11 (signed by the PSF), `claude.exe` (signed by
  > Anthropic) and the WebView2 bootstrapper (signed by Microsoft). Our own binaries are not yet
  > code-signed.
  >
  > We believe this is a false positive and would appreciate a review. Contact: <email>.

---

## 提交前检查

- [ ] 产品负责人同意提交这一家（写明日期）
- [ ] 哈希、字节数与 `dist/SHA256SUMS.txt` / GitHub Release 一致
- [ ] 检出名称是对方原文
- [ ] 联系人、邮箱已填，没有带服务器 IP、密钥
- [ ] 提交后把时间、工单号回填 [detections.md](detections.md)
