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

# 3. 构建三个原生工具：诊断修复 / 32 位探针 / 更新程序
#    漏掉这步，第 5 步会直接失败（它们都在 EXCEL_REQUIRED_FILES 里）
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build-repair.ps1

# 4. 生成内置 Python（首次需下载依赖）
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\package-python.ps1

# 5. 生成 ZIP 支持包、单文件安装器、SHA-256 校验值和签名更新清单
python scripts\package_release.py --version 0.5.0
```

> 第 3 步是 2026-09-19 补进本文档的。此前这里少了它，而 `DeepExcel.Repair.exe`、
> `DeepExcel.Probe32.exe`、`DeepExcel.Updater.exe` 只由 `build-repair.ps1` 生成——
> 照旧文档走，第 5 步会停在 `Required release item is missing`。

输出：

- `dist\DeepExcel.Setup.exe`：唯一面向用户的交付物。
- `dist\DeepExcel-v0.5.0.zip`：支持/诊断 payload，不作为用户安装方式。
- `dist\SHA256SUMS.txt`：两个产物的 SHA-256,**必须与安装包一起发布**。
- `dist\update.json`：签名更新清单，**必须发布到更新源**，否则已安装的客户端收不到这一版。详见下节。

## 自动更新：发版时必须做的两件事

已安装的客户端靠 `update.json` 发现新版本。**发版时漏掉它，不会有任何报错，只是从此没人再升级**——看起来和"还没人升级"完全一样。

打包前设两个环境变量：

```powershell
$env:DEEPEXCEL_UPDATE_KEY = 'D:\offline\deepexcel-update.pem'   # 离线保管的私钥
$env:DEEPEXCEL_UPDATE_BASE_URL = 'https://github.com/<owner>/<repo>/releases/download/v0.5.0'
python scripts\package_release.py --version 0.5.0
```

`DEEPEXCEL_UPDATE_BASE_URL` 是安装器最终被下载的目录，打包脚本会拼成 `<base>/DeepExcel.Setup.exe` 写进签名清单。可选：`DEEPEXCEL_UPDATE_CHANNEL`（默认 `stable`）、`DEEPEXCEL_UPDATE_NOTES`、`DEEPEXCEL_UPDATE_MINIMUM`（低于该版本不自动升级，提示手动重装）。

发布时：

1. 把 `DeepExcel.Setup.exe` 放到 `DEEPEXCEL_UPDATE_BASE_URL` 指向的位置；
2. 把 `dist\update.json` 复制到更新源服务器的 `UPDATE_MANIFEST_DIR` 下，命名为 `<channel>.json`（通常是 `stable.json`）。放文件即生效，不需要重启服务。

**尚未生成正式签名密钥**时，打包会跳过 `update.json` 并打印生成命令——这是刻意的：客户端内置公钥也为空，两边一致地"自动更新未启用"。生成密钥见 README 的「自动更新」一节，生成后必须重新编译并发布一次客户端，之后的版本才能自动升级。

若确实要发一版不带更新清单的：`--no-update-manifest`。它会打印醒目提示，不会静默跳过。

**只改了知识技能**（`src/DeepExcel.Sidecar/knowledge/`）不必发版：用同一把私钥
`python scripts\knowledge_pack.py build --key <pem> --out knowledge_pack.json`，把它放进 `UPDATE_MANIFEST_DIR`
即可，客户端下次更新检查时验签替换。细节见 `server/README.md` 的「知识包」一节。

## 签名状态：当前不签名

DeepExcel 目前**没有购买代码签名证书**，发布形态就是未签名安装包 + 公布 SHA-256。

2026-09-12 移除了原先的自签名内测方案。原方案要求用户运行 `Install-Internal-Certificate.cmd`，把我们自签发的 CA 装进系统根证书存储——该私钥一旦泄露，持有者可为任意软件签名并被用户机器无条件信任。为省掉一次 SmartScreen 放行点击而引入系统级信任风险，不划算。未签名 + 校验值是更干净的选择。

用户侧的表现与应对：

- 首次运行出现「Windows 已保护你的电脑」→ 点击「更多信息」→「仍要运行」。
- 校验完整性：`certutil -hashfile DeepExcel.Setup.exe SHA256`，与 `SHA256SUMS.txt` 比对。
- 下载页必须放 SmartScreen 对话框截图并标注点击位置。纯文字说明的首装完成率显著更低。

### 将来购买证书后

签名链路完整保留，届时只是环境变量变化，不需要改代码：

```powershell
# 方式一：PFX
$env:DEEPEXCEL_PFX = 'C:\secure\deepexcel-code-signing.pfx'
$env:DEEPEXCEL_PFX_PASS = '<secret>'

# 方式二：证书存储（EV USB token 私钥不可导出时用这个）
$env:DEEPEXCEL_CERT_THUMBPRINT = '<40 位 SHA-1 指纹>'

# 方式三：Azure Trusted Signing
$env:USE_AZURE_TRUSTED_SIGNING = '1'

python scripts\package_release.py --version 0.4.17
```

三种来源必须**只配置一种**，否则打包直接失败。任何来源的证书都会校验：必须含私钥、具备代码签名 EKU、在有效期内、**非自签名**（`Subject -eq Issuer` 直接拒绝）、且能构建到受信任根。这条校验是为了让被移除的自签名流程无法通过设环境变量复活。

本机验证时若不想动用已配置的证书：

```powershell
python scripts\package_release.py --version 0.4.17 --force-unsigned
```

## 发布前检查

1. 关闭 Excel 和 WPS 表格。
2. 在一台未安装开发工具的 Windows 10/11 虚拟机运行已签名安装器。
3. Excel 中检查 DeepExcel Ribbon、任务窗格和 AI 对话。
4. WPS 表格中检查 DeepExcel Ribbon、任务窗格和 AI 对话。
5. 确认 `%APPDATA%\kingsoft\wps\jsaddons\publish.xml` 中其他插件节点仍然存在。
6. 执行卸载，确认 Excel COM 注册和 WPS DeepExcel 节点被清理，其他 WPS 插件不受影响。
7. 若安装器报告旧版机器级注册残留，让管理员先卸载/清理旧版，再由目标用户正常运行安装器；不要用另一个管理员账户代装用户级插件。
8. **确认 `dist\update.json` 已发布到更新源**，且 `DeepExcel.Setup.exe` 真的在清单里那个 URL 上。验证方式：`curl https://<更新源>/api/v1/updates/latest` 应返回刚生成的那份清单。
9. **在一台装着上一版的机器上验证自动升级**：等它后台检查（启动 1 分钟后，之后每 6 小时；也可在面板里手动"检查更新"），看是否出现"新版本已就绪"，点"重启安装"后能否装上。这是唯一能证明更新链路真的通了的检查。
10. 发版后看运营后台的「7 天成功升级」和「更新装不上（已放弃）」。后者非 0 意味着有用户的客户端下载验签都成功却装不上——多半被安全软件拦掉，需要主动联系。

无人值守的干净机器验收：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\verify-install-sandbox.ps1 -Launch -NoPause
```

注意 Windows Sandbox 里**没有 Office**，所以它覆盖不了任何依赖 Excel 的步骤（上面第 3 条）。它能验的是 SHA-256、静默安装、注册表写入和卸载残留。

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
- 更新签名**私钥绝不出现在服务端或仓库里**。服务端只原样转发已签名的清单，所以拿下服务端也推不了更新——这是整个更新安全模型的全部内容。`update_signing.py genkey` 会拒绝把私钥写进仓库，`.gitignore` 再兜一层。
- 客户端 `EmbeddedUpdateKey` 里的公钥常量只能由 `update_signing.py genkey` 改写，禁止手改。手改会让公钥与私钥悄悄不一致，**所有客户端静默停止升级**；打包时会用客户端那份公钥回验刚生成的清单，不匹配直接构建失败。
- 任何被 git 跟踪的 `src\DeepExcel.Wps` 源文件都必须在 `WPS_ITEMS` 里。`jsplugins.xml` 曾因此漏打包，用户装到的加载项选项卡在、所有 JS 回调是死的。守卫以 `git ls-files` 为唯一真相。

## 遗留清理（2026-09-12 一次性）

移除自签名方案后，构建机上仍可能残留旧的私钥和已分发的旧包。私钥留在证书存储里就仍然是可被滥用的签名能力，应当清掉：

```powershell
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq 'CN=DeepExcel Internal Testing' } | Remove-Item
```

装过内测包的机器上，还应让用户从受信任根存储中移除该证书：

```powershell
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -eq 'CN=DeepExcel Internal Testing' } | Remove-Item
```

同时下架所有 `dist\DeepExcel-Internal-*.zip` 与 `dist\DeepExcel.Setup.INTERNAL.exe`。
