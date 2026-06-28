# DeepExcel 注册脚本（管理员权限运行）
# 用于将AddIn注册到Excel

param(
    [switch]$Unregister = $false
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$publishDir = Join-Path $rootDir "dist\publish"
$dllPath = Join-Path $publishDir "DeepExcel.AddIn.dll"

if (-not (Test-Path $dllPath)) {
    Write-Error "未找到 $dllPath，请先运行 build.ps1"
}

# 检查管理员权限
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warning "需要管理员权限注册COM加载项！请以管理员身份运行PowerShell"
    exit 1
}

# 找到合适的regasm
$regasmPaths = @(
    "${env:ProgramFiles}\dotnet\sdk\8.0.*\Roslyn\bincore\regasm.exe",
    "${env:ProgramFiles}\dotnet\sdk\7.0.*\Roslyn\bincore\regasm.exe",
    "${env:ProgramFiles}\dotnet\sdk\6.0.*\Roslyn\bincore\regasm.exe",
    "${env:WINDIR}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
)

$regasm = $null
foreach ($pattern in $regasmPaths) {
    $found = Get-Item $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { $regasm = $found.FullName; break }
}

if (-not $regasm) {
    Write-Error "未找到 regasm.exe，请先安装 .NET Framework 4.8 SDK 或 .NET SDK 6+"
}

Write-Host "使用: $regasm" -ForegroundColor Cyan
Write-Host "加载项: $dllPath" -ForegroundColor Cyan

if ($Unregister) {
    Write-Host "[注销] DeepExcel.AddIn..." -ForegroundColor Yellow
    & $regasm /u $dllPath
} else {
    Write-Host "[注册] DeepExcel.AddIn..." -ForegroundColor Yellow
    & $regasm /codebase $dllPath
}

if ($LASTEXITCODE -eq 0) {
    Write-Host "完成！请重启Excel使加载项生效。" -ForegroundColor Green
} else {
    Write-Error "注册失败，退出码: $LASTEXITCODE"
}
