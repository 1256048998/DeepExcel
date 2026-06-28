# DeepExcel 构建与发布脚本 (PowerShell)
# 用法: .\build.ps1 -Configuration Release -Platform AnyCPU

param(
    [string]$Configuration = "Release",
    [string]$Platform = "AnyCPU",
    [switch]$Clean = $false
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $rootDir "src"
$addInDir = Join-Path $srcDir "DeepExcel.AddIn"
$uiDir = Join-Path $srcDir "DeepExcel.UI"
$outputDir = Join-Path $rootDir "dist"
$publishDir = Join-Path $outputDir "publish"

Write-Host "==> DeepExcel 构建脚本" -ForegroundColor Cyan
Write-Host "    配置: $Configuration | 平台: $Platform" -ForegroundColor Gray

# 1. 清理
if ($Clean) {
    Write-Host "[1/5] 清理输出..." -ForegroundColor Yellow
    if (Test-Path $outputDir) { Remove-Item $outputDir -Recurse -Force }
}

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
New-Item -ItemType Directory -Force -Path $publishDir | Out-Null

# 2. 构建UI
Write-Host "[2/5] 构建前端..." -ForegroundColor Yellow
Set-Location $uiDir
if (-not (Test-Path "node_modules")) {
    Write-Host "  安装依赖..." -ForegroundColor Gray
    npm install --no-audit --no-fund
}
npm run build
if ($LASTEXITCODE -ne 0) { throw "UI构建失败" }

# 复制UI构建产物到AddIn
$uiDist = Join-Path $uiDir "dist"
$uiTarget = Join-Path $addInDir "WebViewAssets"
if (Test-Path $uiTarget) { Remove-Item $uiTarget -Recurse -Force }
Copy-Item $uiDist $uiTarget -Recurse
Write-Host "  UI资源已复制到: $uiTarget" -ForegroundColor Green

# 3. 还原NuGet包
Write-Host "[3/5] 还原NuGet包..." -ForegroundColor Yellow
Set-Location $addInDir
dotnet restore
if ($LASTEXITCODE -ne 0) { throw "NuGet还原失败" }

# 4. 构建AddIn
Write-Host "[4/5] 构建AddIn..." -ForegroundColor Yellow
dotnet build -c $Configuration -p:Platform=$Platform --no-restore
if ($LASTEXITCODE -ne 0) { throw "AddIn构建失败" }

# 5. 发布
Write-Host "[5/5] 发布..." -ForegroundColor Yellow
dotnet publish -c $Configuration -p:Platform=$Platform -o $publishDir --no-build
if ($LASTEXITCODE -ne 0) { throw "发布失败" }

Write-Host ""
Write-Host "==> 构建完成!" -ForegroundColor Green
Write-Host "    发布目录: $publishDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "后续步骤:" -ForegroundColor Yellow
Write-Host "  1. 打开 Excel" -ForegroundColor Gray
Write-Host "  2. 文件 → 选项 → 加载项 → 管理: COM加载项 → 转到" -ForegroundColor Gray
Write-Host "  3. 取消勾选'DeepExcel.AddIn'后重新勾选（或运行regsvr32注册）" -ForegroundColor Gray
