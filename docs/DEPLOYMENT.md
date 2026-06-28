# DeepExcel 安装包发布说明

## 系统要求
- Windows 10/11 (x64)
- Excel 2016 或更高版本 (Office 365)
- .NET Framework 4.8
- WebView2 Runtime (Win11自带，Win10需安装)
- 网络连接（用于调用AI模型API）

## 构建步骤

### 1. 开发模式运行
```powershell
# 安装依赖
cd src\DeepExcel.UI
npm install

# 启动UI开发服务器
npm run dev

# 另开终端编译AddIn
cd src\DeepExcel.AddIn
dotnet build
```

### 2. 生产构建
```powershell
# 完整构建（前端+后端+发布）
.\scripts\build.ps1 -Configuration Release

# 注册到Excel（需管理员权限）
.\scripts\register.ps1
```

### 3. 注销
```powershell
.\scripts\register.ps1 -Unregister
```

## 发布目录结构
```
dist\publish\
├── DeepExcel.AddIn.dll           # 主加载项
├── DeepExcel.AddIn.tlb           # 类型库
├── WebViewAssets\                # 前端资源
│   ├── index.html
│   └── assets\
├── Microsoft.Office.Interop.Excel.dll
├── Microsoft.Web.WebView2.*
└── ...
```

## Excel注册表项

加载项会在以下位置写入注册表：
- `HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\DeepExcel.AddIn`
- `HKEY_CLASSES_ROOT\CLSID\{...}` (COM组件)

如需手动添加Excel信任，可通过：
- 文件 → 选项 → 加载项 → 管理: 禁用项目 → 转到 → 取消勾选DeepExcel.AddIn → 重启Excel → 再次勾选

## 配置文件位置
- 配置: `%APPDATA%\DeepExcel\config.json`
- 日志: `%APPDATA%\DeepExcel\logs\deepexcel-YYYYMMDD.log`
- 快照: `%APPDATA%\DeepExcel\snapshots\`

## 故障排除

1. **加载项未出现在Excel中**
   - 检查注册表HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins
   - 确认LoadBehavior = 3
   - 重启Excel

2. **WebView2白屏**
   - 安装WebView2 Runtime: https://developer.microsoft.com/microsoft-edge/webview2/
   - 检查WebViewAssets目录是否完整

3. **API调用失败**
   - 在%APPDATA%\DeepExcel\config.json中确认API Key
   - 检查网络连接与代理设置

4. **日志查看**
   - 实时日志: %APPDATA%\DeepExcel\logs\
   - 调试模式: 在config.json中设置 MinLevel = 0
