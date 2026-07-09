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
$webSrcDir = Join-Path $uiDir "dist"
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
    Copy-Item -Path (Join-Path $sidecarSrc "*.py") -Destination $sidecarDest -Force
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
    "jsa-executor.js"
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
$indexHtml = Join-Path $webDestDir "index.html"
if (Test-Path $indexHtml) {
    Write-Host "  [OK] web/index.html" -ForegroundColor Green
} else {
    Write-Host "  [MISSING] web/index.html" -ForegroundColor Red
    $allPresent = $false
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
