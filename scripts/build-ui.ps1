# DeepExcel task-pane front end build.
#
# Builds src\DeepExcel.UI with Vite. The output goes straight to
# src\DeepExcel.AddIn\WebViewAssets (see vite.config.ts), which
# _compile_only.ps1 then copies to bin\Release. Excel and WPS load the same
# COM add-in, so this one build serves both.
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build-ui.ps1
#
# Keep this file ASCII-only: Windows PowerShell 5.1 reads a BOM-less .ps1 with
# the system ANSI codepage, and non-ASCII text breaks parsing on CJK systems.

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$uiDir = Join-Path $projectRoot 'src\DeepExcel.UI'
$assetsDir = Join-Path $projectRoot 'src\DeepExcel.AddIn\WebViewAssets'

if (-not (Test-Path (Join-Path $uiDir 'package.json'))) {
    throw "DeepExcel.UI\package.json not found under $uiDir"
}

Push-Location $uiDir
try {
    if (-not (Test-Path (Join-Path $uiDir 'node_modules'))) {
        Write-Host 'Installing npm dependencies...' -ForegroundColor Gray
        npm install --silent
        if ($LASTEXITCODE -ne 0) { throw 'npm install failed' }
    }
    Write-Host 'Running vite build...' -ForegroundColor Gray
    npm run build --silent
    if ($LASTEXITCODE -ne 0) { throw 'vite build failed' }
} finally {
    Pop-Location
}

# The packager fails on these too, but only at the very end of a release run.
foreach ($file in 'index.html', 'assets\index.js', 'assets\index.css') {
    $path = Join-Path $assetsDir $file
    if (-not ((Test-Path -LiteralPath $path -PathType Leaf) -and ((Get-Item -LiteralPath $path).Length -gt 0))) {
        throw "Front-end build output missing or empty: $path"
    }
}

Write-Host "Front end built into $assetsDir" -ForegroundColor Green
