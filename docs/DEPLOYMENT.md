# DeepExcel 发布与安装

DeepExcel 对外只分发单文件 `DeepExcel.Setup.exe`。同一个安装器同时支持 Microsoft Excel 和 WPS 表格：

- Excel：按当前用户写入真实的 32/64 位 COM 注册表视图，安装后分别用 32/64 位 CLR 做激活验证，同时安装 x86/x64/ARM64 WebView2 loader。
- WPS：使用新版 WPS 支持的 `publish.xml` 模式，只合并/删除 DeepExcel 节点，不覆盖其他 WPS 插件。
- 两个宿主共享安装器内置的 Python 3.11 和 sidecar，用户不需安装 Python。
- 安装器检测 WebView2 Runtime，缺失时调用已验证 Microsoft 签名的 Evergreen bootstrapper。

## 生产构建

```powershell
# 1. 构建共享前端和 WPS 资源
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build-wps.ps1

# 2. 构建 Excel COM 加载项和依赖
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\_compile_only.ps1

# 3. 生成内置 Python（首次需下载依赖）
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\package-python.ps1

# 4. 配置 Authenticode 代码签名
$env:DEEPEXCEL_PFX = 'C:\secure\deepexcel-code-signing.pfx'
$env:DEEPEXCEL_PFX_PASS = '<secret>'

# 5. 生成 ZIP 支持包和签名的单文件安装器
python scripts\package_release.py --version 0.4.17
```

输出：

- `dist\DeepExcel.Setup.exe`：唯一面向用户的交付物。
- `dist\DeepExcel-v0.4.17.zip`：支持/诊断 payload，不作为用户安装方式。

没有签名证书时，只允许本地验证：

```powershell
python scripts\package_release.py --version 0.4.17 --allow-unsigned
```

这种产物输出为 `dist\DeepExcel.Setup.UNSIGNED-LOCAL.exe`，只允许本机验证，不能发给用户。生产流程在未配置签名时会直接失败，避免误发未签名安装包。

## 发布前检查

1. 关闭 Excel 和 WPS 表格。
2. 在一台未安装开发工具的 Windows 10/11 虚拟机运行已签名安装器。
3. Excel 中检查 DeepExcel Ribbon、任务窗格和 AI 对话。
4. WPS 表格中检查 DeepExcel Ribbon、任务窗格和 AI 对话。
5. 确认 `%APPDATA%\kingsoft\wps\jsaddons\publish.xml` 中其他插件节点仍然存在。
6. 执行卸载，确认 Excel COM 注册和 WPS DeepExcel 节点被清理，其他 WPS 插件不受影响。
7. 若安装器报告旧版机器级注册残留，让管理员先卸载/清理旧版，再由目标用户正常运行安装器；不要用另一个管理员账户代装用户级插件。

## 关键防退化约束

- 缺少 `Extensibility.dll`、内置 Python、WPS web/sidecar 或任一 WebView2 位数 loader 时，打包立即失败。
- AnyCPU 包不允许根目录出现 x86 `WebView2Loader.dll`，避免 64 位 Excel 白屏。
- Excel 加载项写入 Microsoft 文档指定的 `HKCU\Software\Microsoft\Office\Excel\Addins\DeepExcel.AddIn`。
- Excel 注册不能依赖 `App Paths\excel.exe` 是否存在；企业/MSI Office 可能没有该键，但仍必须注册加载项。
- 32 位 COM 必须写入 Windows `Registry32`/Inno `HKCU32` 视图，禁止在 64 位视图下手拼 `WOW6432Node` 路径。
- WPS 不再使用新版已限制的 `jsplugins.xml/oem.ini` 部署路径。
- WPS `publish.xml` 由安装包内置 Python 标准库更新，通过同目录临时文件原子替换，并在替换前校验文件未被其他进程改写；该步骤不依赖用户的 Windows PowerShell 环境。
- WPS 离线节点必须使用官方 `wpsjs` 格式：`url="DeepExcel_<version>"`，并与 `jsaddons` 下的目录名完全一致；禁止写成 `file://`。
- 内置 Python、pip 与全部 Python 依赖均固定版本和 SHA-256；依赖变化必须同步更新 `scripts\python-requirements.lock.txt`。

## 内部测试版（免费自签名）

内部测试版不会改变本机 Excel/WPS 的开发注册，也不会把私钥写入仓库或安装包。私钥以不可导出形式保存在构建电脑的当前用户证书存储区：

```powershell
# 首次执行或证书到期后执行
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\new-internal-signing-cert.ps1

# 构建明确标注的内测包
python scripts\package_release.py --version 0.4.17 --internal
```

发送给受邀测试用户的文件只有 `dist\DeepExcel-Internal-v0.4.17.zip`。用户先运行包内 `Install-Internal-Certificate.cmd`，再运行 `DeepExcel.Setup.INTERNAL.exe`。

此证书不是公共 CA 证书。它只适用于明确知情的内部/受邀测试用户，不得把 `*.INTERNAL.exe` 或 `DeepExcel-Internal-*.zip` 放到公开下载页。正式发布命令不会把该证书视为生产签名，缺少 PFX/Azure 正式凭据时仍会失败。
