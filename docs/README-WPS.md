# DeepExcel 在 WPS 表格中的部署与排查

## 概述

WPS 表格（ET）加载的是**和 Excel 同一个 COM 加载项**（`DeepExcel.AddIn`）：同一个 DLL、同一套 WebView2
面板、同一个 Python sidecar、同一套表格工具。没有单独的 WPS 版本，也没有单独的 WPS 构建步骤。

2026-11-23 在真机（WPS 12.1.0.29690 个人版，32 位 et.exe）上验证：功能区、WebView2 面板、sidecar、
模型对话、读写工具、建表画图整条链路都通。

## 系统要求

| 组件 | 要求 |
|------|------|
| 操作系统 | Windows 10/11（64 位） |
| WPS Office | 能加载 COM 加载项的 WPS 表格；已验证 12.1 个人版 |
| .NET Framework | 4.8（COM 加载项本身需要） |
| WebView2 Runtime | 安装器检测到 WPS 时会一并安装 |

Python 和 sidecar 由安装器内置，用户不需要装 Python 或 Node。

## 安装器做了什么

一个 `DeepExcel.Setup.exe` 同时服务 Excel 和 WPS。WPS 部分在安装器里单独一个 try 块，Excel 那半失败
（例如机器上根本没有 Excel）不会跳过它：

- 登记 `HKCU\Software\Kingsoft\Office\ET\AddIns\DeepExcel.AddIn`：`Description`、`FriendlyName`、
  `LoadBehavior=3`；
- 在 `HKCU\Software\Kingsoft\Office\ET\AddinsWL` 下写一个**名为 `DeepExcel.AddIn` 的字符串值**（白名单是
  值列表，不是子键；写成子键 WPS 不认）；
- 删掉 `ET\AddinsCL` 里的 `DeepExcel.AddIn` 值。WPS 第一次见到陌生 COM 加载项会把它放进这个禁用列表，
  弹「已禁用的COM加载项」，用户点「启用」做的也只是删掉这个值。实测装时删掉后不再弹；
- 没检测到 Excel 时，Excel 激活失败会回滚 COM 类注册，WPS 步骤会把 COM 类重新注册回来——WPS 也要靠它们；
- 清理旧版本留下的 WPS JS 加载项，见下一节。

卸载时移除上面这些登记，并同样清理 JS 加载项残留。

## 旧版 JS 加载项的清理

早期版本给 WPS 另装过一个 JS 加载项：`%APPDATA%\kingsoft\wps\jsaddons\DeepExcel_<版本>\` 目录，加上
`jsaddons\publish.xml` 里一个 `name="DeepExcel"` 的节点。它在 WPS 里跑不起来——WPS 的 JS 加载项宿主是
纯浏览器页面，没有 Node（`require is not defined`），而那个加载项的 sidecar、配置、凭据、历史都依赖 Node。
所以它已从仓库和安装包里删除。

留着它，WPS 里会多出第二个点了没反应的 DeepExcel 选项卡。安装和卸载都会：

1. 从 `publish.xml` 删掉 DeepExcel 自己的节点（其他 WPS 插件的节点不动；同目录临时文件原子替换，旁边留
   `.bak`）。机器上没有 `publish.xml` 就什么也不做，不会凭空创建；
2. 再删掉 `jsaddons\DeepExcel_*` 目录。

先删节点再删目录：中途失败最多留下一个没人引用的目录，不会留下指向不存在目录的节点。清理只是收尾，
失败（例如别的插件把 `publish.xml` 写坏了）只记进安装日志，不会让安装或卸载失败。

## 已修的 WPS 兼容差异

COM 加载项是同一份代码，WPS 和 Excel 对象模型不一致的地方在 C# 里兜：

- `Application.Range["表名!A1:G42"]` 在 WPS 抛 `0x8FE3080B` → 回退为先取表再取区域（`ExcelTarget.ResolveOnSheet`）；
- CustomTaskPane 显示前设的 Width 在 WPS 不生效，面板只有约 80 px → 显示后再纠正；
- `imageMso="WindowUnhide"` WPS 没有 → 功能区图标改为自绘（`RibbonIcons.cs`）。

## 仍未验证

- **只装 WPS、没装 Office 的机器。** Windows Sandbox 里 DeepExcel 自己的 COM 激活报 `0x80070002`（沙箱没有
  Office），这在纯 WPS 机器上可能同样成立。另一个线索：`OFFICE.dll` 依赖 `stdole 7.0.3300.0`，安装包没有带它，
  开发机上的 stdole 来自 Office 装进 GAC 的那份。拿到一台这样的机器验证之前，不能宣称支持纯 WPS 机器。
- **逐项工具。** 真机跑通的是面板、对话、读写、建表画图；透视表、快照回滚、VBA（WPS 个人版默认不带 VBA 环境）
  等还没有在 WPS 上逐个走过。
- **侧车里的 WPS 宿主模式已经用不到。** `sidecar.py` / `excel_tools.py` 里按 `DEEPEXCEL_HOST=wps` 切换的那套
  （`WPS_HOST_TOOLS` 工具子集、`execute_jsa`、`wps-jsa` 知识技能）只有旧 JS 加载项会打开。COM 加载项在 WPS 里
  启动 sidecar 时不设这个变量，WPS 会话拿到的是和 Excel 一样的完整工具集。这部分代码待清理。

## 本地开发

和 Excel 完全一样：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build-ui.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\_compile_only.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\register-user.ps1
```

`register-user.ps1` 把 COM 类指向开发版 DLL，但只写 Excel 的加载项登记。WPS 按 ProgID 找 COM 类，所以只要
WPS 那三处登记在（装过一次安装包就有，或按上面「安装器做了什么」手动写），WPS 加载的就是 `register-user.ps1`
指向的那个 DLL。编译前关闭 Excel 和 WPS，否则 DLL 被锁。

### 在沙箱里试用（不在本机装 WPS）

```powershell
powershell -ExecutionPolicy Bypass -File scripts\try-wps-sandbox.ps1
```

开一个 Windows Sandbox（默认不限内存，`-MemoryMB` 可设上限），装好最新公开发布的 DeepExcel（从 GitHub
Release 下载并核对 SHA-256；`-SetupPath dist\DeepExcel.Setup.exe` 换成本地构建），确认它已登记给 WPS，
然后打开 WPS 表格。关掉沙箱窗口，WPS 和 DeepExcel 全部消失，本机不留任何东西。

- 第一次：沙箱里的浏览器打开 wps.cn，手动下载并安装 WPS。装好后安装包会被缓存到
  `dist\wps-sandbox\stage\`，之后每次自动启动安装，只需点「立即安装」。
- 本机 `Downloads` 里已有 `WPS*.exe`，或用 `-WpsSetupPath` 指定，可以跳过下载。
- 日志在 `dist\wps-sandbox\out\`（`guest.log` 末行 `RESULT: READY / FAILED`）。
- ⚠️ 2026-11-23 在开发机上，沙箱**每次启动约 4 分钟就崩**（`The virtual machine or container exited
  unexpectedly. (0x80370106)`），原因未定位。这台机器上 WPS 改为直接装本机来测。

发布验收脚本 `scripts\verify-install-sandbox.ps1 -WithWps` 也会检查上面的 WPS 登记、旧 JS 加载项是否清理干净，
并拉起 WPS 让人确认选项卡和面板。

## 故障排查

### WPS 里没有 DeepExcel 选项卡

1. 完全退出 WPS（任务管理器里没有残留的 `et.exe` / `wps.exe`）后重开；
2. 弹出「已禁用的COM加载项」时，点 DeepExcel.AddIn 旁的「启用」；
3. 检查 `HKCU\Software\Kingsoft\Office\ET\AddIns\DeepExcel.AddIn` 存在且 `LoadBehavior` 为 3，
   `ET\AddinsWL` 下有名为 `DeepExcel.AddIn` 的值，`ET\AddinsCL` 下没有；缺了就重新运行安装包；
4. COM 类注册坏了时，从「开始菜单 → DeepExcel → DeepExcel 诊断与修复」修复。注意它目前只修 COM 类和 Excel
   登记，不写上面的 WPS 登记。

### 出现两个 DeepExcel 选项卡，其中一个点了没反应

旧版 JS 加载项没清掉。重新运行最新安装包即可；或手动删掉 `%APPDATA%\kingsoft\wps\jsaddons\publish.xml`
里 `name="DeepExcel"` 的节点和同目录下的 `DeepExcel_*` 文件夹。

### 面板白屏或对话没反应

和 Excel 端同一套排查：查看 `%APPDATA%\DeepExcel\logs\` 下的加载项与 sidecar 日志，反馈时附上 WPS 版本号和截图。
