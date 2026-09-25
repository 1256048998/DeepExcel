# DeepExcel WPS 加载项构建脚本
# 构建 React 前端 + 复制到 WPS 加载项目录
#
# 用法：powershell -ExecutionPolicy Bypass -File scripts\build-wps.ps1

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$projectRoot = Split-Path -Parent $scriptDir

$uiDir = Join-Path $projectRoot "src\DeepExcel.UI"
$wpsDir = Join-Path $projectRoot "src\DeepExcel.Wps"
$webDestDir = Join-Path $wpsDir "web"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "DeepExcel WPS Addin Build" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# 1. 构建 React 前端
Write-Host "[1/3] Building React frontend..." -ForegroundColor Yellow
$uiPkg = Join-Path $uiDir "package.json"
if (-not (Test-Path $uiPkg)) {
    Write-Host "ERROR: DeepExcel.UI/package.json not found at $uiPkg" -ForegroundColor Red
    exit 1
}

# ★ 复用现有 Vite 构建（与 Excel 端共用源码）
Push-Location $uiDir
try {
    # 确保 node_modules 存在
    if (-not (Test-Path (Join-Path $uiDir "node_modules"))) {
        Write-Host "  Installing npm dependencies..." -ForegroundColor Gray
        npm install --silent
        if ($LASTEXITCODE -ne 0) { throw "npm install failed" }
    }

    # 构建
    Write-Host "  Running vite build..." -ForegroundColor Gray
    npm run build --silent
    if ($LASTEXITCODE -ne 0) { throw "vite build failed" }
} finally {
    Pop-Location
}
Write-Host "  React frontend built successfully." -ForegroundColor Green
Write-Host ""

# 2. 复制构建产物到 WPS 加载项的 web/ 目录
Write-Host "[2/3] Copying build artifacts to WPS addin..." -ForegroundColor Yellow
$webSrcDir = Join-Path $projectRoot "src\DeepExcel.AddIn\WebViewAssets"
if (-not (Test-Path $webSrcDir)) {
    Write-Host "ERROR: dist directory not found at $webSrcDir" -ForegroundColor Red
    exit 1
}

# 清空旧产物
if (Test-Path $webDestDir) {
    Remove-Item -Path $webDestDir -Recurse -Force
}
New-Item -Path $webDestDir -ItemType Directory -Force | Out-Null

# 复制所有文件
Copy-Item -Path (Join-Path $webSrcDir "*") -Destination $webDestDir -Recurse -Force
Write-Host "  Copied to: $webDestDir" -ForegroundColor Green

# ★ 复制 sidecar 文件到 WPS 加载项目录（Python sidecar 共用）
$sidecarSrc = Join-Path $projectRoot "src\DeepExcel.Sidecar"
$sidecarDest = Join-Path $wpsDir "sidecar"
if (Test-Path $sidecarSrc) {
    if (Test-Path $sidecarDest) {
        Remove-Item -Path $sidecarDest -Recurse -Force
    }
    New-Item -Path $sidecarDest -ItemType Directory -Force | Out-Null
    # 与 _compile_only.ps1 一致：整棵侧车源码树（含子包），排除测试与缓存
    Get-ChildItem -LiteralPath $sidecarSrc -Force | Where-Object {
        $_.Name -notin @('tests', '__pycache__', '.pytest_cache') -and
        ($_.PSIsContainer -or $_.Extension -in @('.py', '.md', '.json'))
    } | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $sidecarDest -Recurse -Force
    }
    Get-ChildItem -LiteralPath $sidecarDest -Recurse -Directory -Force |
        Where-Object { $_.Name -in @('__pycache__', 'tests', '.pytest_cache') } |
        Remove-Item -Recurse -Force
    Write-Host "  Copied sidecar Python files to: $sidecarDest" -ForegroundColor Green
}

Write-Host ""

# 3. 验证文件完整性
Write-Host "[3/3] Verifying WPS addin files..." -ForegroundColor Yellow
$requiredFiles = @(
    "main.js",
    "ribbon.xml",
    "jsplugins.xml",
    "taskpane.html",
    "sidecar-host.js",
    "tool-dispatcher.js",
    "wps-actions.js",
    "range-paging.js",
    # 工作簿记忆的面板入口（与 Excel 端共用 workbooks/<hash>/NOTES.md）
    "workbook-memory-store.js",
    "jsa-executor.js",
    # 首次使用：工作簿结构 / 插入示例数据
    "starter-host.js",
    # 先读后写 + 读后被改检测
    "read-ledger.js",
    # ★ 模型配置（厂商 / 模型优先级 / API Key），与 Excel 端共用 config.json + DPAPI 凭据
    "config-store.js",
    "credential-store.js",
    "model-service.js",
    "dpapi-cli.py",
    # ★ 对话历史 + 附件
    "conversation-store.js",
    "attachment-store.js",
    "images\panel.svg",
    "images\help.svg"
)

$allPresent = $true
foreach ($file in $requiredFiles) {
    $path = Join-Path $wpsDir $file
    if (Test-Path $path) {
        Write-Host "  [OK] $file" -ForegroundColor Green
    } else {
        Write-Host "  [MISSING] $file" -ForegroundColor Red
        $allPresent = $false
    }
}

# 检查 web/ 目录
$requiredWebFiles = @(
    "index.html",
    "assets\index.js",
    "assets\index.css"
)
foreach ($file in $requiredWebFiles) {
    $path = Join-Path $webDestDir $file
    if ((Test-Path -LiteralPath $path -PathType Leaf) -and ((Get-Item -LiteralPath $path).Length -gt 0)) {
        Write-Host "  [OK] web/$($file.Replace('\', '/'))" -ForegroundColor Green
    } else {
        Write-Host "  [MISSING/EMPTY] web/$($file.Replace('\', '/'))" -ForegroundColor Red
        $allPresent = $false
    }
}

# 检查 sidecar
$sidecarPy = Join-Path $sidecarDest "sidecar.py"
if (Test-Path $sidecarPy) {
    Write-Host "  [OK] sidecar/sidecar.py" -ForegroundColor Green
} else {
    Write-Host "  [MISSING] sidecar/sidecar.py" -ForegroundColor Red
    $allPresent = $false
}

if (-not $allPresent) {
    Write-Host ""
    Write-Host "ERROR: Some files are missing. Build incomplete." -ForegroundColor Red
    exit 1
}

& node (Join-Path $scriptDir 'test-wps-main.js')
if ($LASTEXITCODE -ne 0) { throw 'WPS main.js smoke test failed' }

& node (Join-Path $scriptDir 'test-wps-model-service.js')
if ($LASTEXITCODE -ne 0) { throw 'WPS model-service test failed' }

& node (Join-Path $scriptDir 'test-wps-range-paging.js')
if ($LASTEXITCODE -ne 0) { throw 'WPS range-paging test failed' }

& node (Join-Path $scriptDir 'test-wps-memory-store.js')
if ($LASTEXITCODE -ne 0) { throw 'WPS workbook memory store test failed' }

& node (Join-Path $scriptDir 'test-wps-starter.js')
if ($LASTEXITCODE -ne 0) { throw 'WPS starter test failed' }

& node (Join-Path $scriptDir 'test-wps-tools.js')
if ($LASTEXITCODE -ne 0) { throw 'WPS tool dispatch test failed' }

& node (Join-Path $scriptDir 'test-wps-ledger.js')
if ($LASTEXITCODE -ne 0) { throw 'WPS read ledger test failed' }

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "Build successful!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host ""
Write-Host "WPS addin directory: $wpsDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Run register-wps.ps1 to register the addin in WPS" -ForegroundColor Gray
Write-Host "  2. Restart WPS Spreadsheets" -ForegroundColor Gray
Write-Host "  3. Find the 'DeepExcel' tab in the ribbon" -ForegroundColor Gray
