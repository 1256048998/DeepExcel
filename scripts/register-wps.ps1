# DeepExcel WPS 加载项注册脚本
# 将 DeepExcel JS 加载项注册到 WPS 表格（ET）
#
# ★ WPS 加载项注册路径与 Office 完全不同：
#   Office: HKCU\Software\Microsoft\Office\16.0\Excel\Addins\<ProgID>
#   WPS:    HKCU\Software\kingsoft\office\ET\AddinsWL\<ProgID>  (AddinsWL = Addins White List)
#
# ★ WPS 不区分 32/64 位，不需要 Wow6432Node
#
# 用法：
#   注册：powershell -ExecutionPolicy Bypass -File scripts\register-wps.ps1
#   注销：powershell -ExecutionPolicy Bypass -File scripts\register-wps.ps1 -Unregister

param(
    [switch]$Unregister = $false
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

# ★ JS 加载项目录（包含 main.js, ribbon.xml, jsplugins.xml, taskpane.html, web/）
$possiblePaths = @(
    (Join-Path $scriptDir "DeepExcel.Wps"),
    (Join-Path (Split-Path -Parent $scriptDir) "src\DeepExcel.Wps")
)

$wpsAddinDir = $null
foreach ($p in $possiblePaths) {
    if ((Test-Path $p) -and (Test-Path (Join-Path $p "main.js"))) {
        $wpsAddinDir = $p
        break
    }
}

if (-not $wpsAddinDir) {
    Write-Host "ERROR: DeepExcel.Wps directory not found (must contain main.js)" -ForegroundColor Red
    Write-Host "Searched paths:" -ForegroundColor Yellow
    foreach ($p in $possiblePaths) { Write-Host "  $p" -ForegroundColor Gray }
    exit 1
}

$wpsAddinDir = Resolve-Path $wpsAddinDir
Write-Host "WPS Addin Directory: $wpsAddinDir" -ForegroundColor Cyan
Write-Host ""

# ★ WPS 加载项 ProgID（与 JS 加载项的 name 字段一致）
$progId = "DeepExcel"

# ★ WPS 表格加载项注册路径
# HKCU\Software\kingsoft\office\ET\AddinsWL\<ProgID>
$wpsAddinKey = "HKCU:\Software\kingsoft\office\ET\AddinsWL\$progId"

# ★ WPS 加载项配置文件路径（jsplugins.xml 的发布位置）
# WPS 启动时会扫描此目录下的 jsplugins.xml
$wpsPluginConfigDir = Join-Path $env:APPDATA "kingsoft\wps\jsplugins"

if ($Unregister) {
    Write-Host "[Unregister] DeepExcel WPS Addin..." -ForegroundColor Yellow

    # 移除 WPS 加载项白名单注册
    if (Test-Path $wpsAddinKey) {
        Remove-Item -Path $wpsAddinKey -Recurse -Force
        Write-Host "  Removed WPS addin whitelist entry: $wpsAddinKey" -ForegroundColor Gray
    }

    # 移除 jsplugins.xml 发布副本
    $publishedConfig = Join-Path $wpsPluginConfigDir "jsplugins.xml"
    if (Test-Path $publishedConfig) {
        Remove-Item -Path $publishedConfig -Force
        Write-Host "  Removed published jsplugins.xml" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "Unregistration complete!" -ForegroundColor Green
    Write-Host "Please restart WPS Spreadsheets to take effect." -ForegroundColor Yellow
} else {
    Write-Host "[Register] DeepExcel WPS Addin..." -ForegroundColor Yellow

    # 1. 注册 WPS 加载项白名单
    # ★ WPS 12.0.1.17xx+ 要求 COM 加载项在白名单中才能加载
    # ★ JS 加载项也需要在此注册（值为空字符串，键名即 ProgID）
    if (-not (Test-Path $wpsAddinKey)) {
        New-Item -Path $wpsAddinKey -Force | Out-Null
    }
    Set-ItemProperty -Path $wpsAddinKey -Name "(default)" -Value "" -Force
    Write-Host "  Registered WPS addin whitelist: $wpsAddinKey" -ForegroundColor Gray

    # 2. 发布 jsplugins.xml 到 WPS 扫描目录
    $sourceConfig = Join-Path $wpsAddinDir "jsplugins.xml"
    if (-not (Test-Path $sourceConfig)) {
        Write-Host "ERROR: jsplugins.xml not found in $wpsAddinDir" -ForegroundColor Red
        exit 1
    }

    if (-not (Test-Path $wpsPluginConfigDir)) {
        New-Item -Path $wpsPluginConfigDir -ItemType Directory -Force | Out-Null
        Write-Host "  Created WPS plugin config dir: $wpsPluginConfigDir" -ForegroundColor Gray
    }

    Copy-Item -Path $sourceConfig -Destination $wpsPluginConfigDir -Force
    Write-Host "  Published jsplugins.xml to $wpsPluginConfigDir" -ForegroundColor Gray

    # 3. 发布 JS 加载项文件到 WPS 加载项目录
    # WPS JS 加载项通常放在 %APPDATA%\kingsoft\wps\jsplugins\<ProgID>\
    $pluginDir = Join-Path $wpsPluginConfigDir $progId
    if (-not (Test-Path $pluginDir)) {
        New-Item -Path $pluginDir -ItemType Directory -Force | Out-Null
    }

    # 复制所有 JS 加载项文件（main.js, ribbon.xml, taskpane.html, web/ 等）
    $filesToCopy = @("main.js", "ribbon.xml", "jsplugins.xml", "taskpane.html", "package.json", "sidecar-host.js", "tool-dispatcher.js", "wps-actions.js", "jsa-executor.js")
    foreach ($file in $filesToCopy) {
        $src = Join-Path $wpsAddinDir $file
        if (Test-Path $src) {
            Copy-Item -Path $src -Destination $pluginDir -Force
        }
    }

    # 复制 web/ 目录（React 构建产物）
    $webDir = Join-Path $wpsAddinDir "web"
    if (Test-Path $webDir) {
        $destWebDir = Join-Path $pluginDir "web"
        if (Test-Path $destWebDir) {
            Remove-Item -Path $destWebDir -Recurse -Force
        }
        Copy-Item -Path $webDir -Destination $pluginDir -Recurse -Force
        Write-Host "  Copied web/ directory (React build)" -ForegroundColor Gray
    } else {
        Write-Host "  WARNING: web/ directory not found. Run build-wps.ps1 first." -ForegroundColor Yellow
    }

    Write-Host "  Published JS addin files to $pluginDir" -ForegroundColor Gray

    Write-Host ""
    Write-Host "Registration successful!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Restart WPS Spreadsheets" -ForegroundColor Gray
    Write-Host "  2. Find the 'DeepExcel' tab in the ribbon" -ForegroundColor Gray
    Write-Host "  3. Click 'Open Panel' button to start" -ForegroundColor Gray
    Write-Host ""
    Write-Host "If you don't see the DeepExcel tab:" -ForegroundColor Yellow
    Write-Host "  WPS -> Options -> Add-Ins -> COM Add-ins" -ForegroundColor Gray
    Write-Host "  Or check: $wpsAddinKey" -ForegroundColor Gray
    Write-Host ""
    Write-Host "★ WPS Personal Edition (12.0.1.17xx+) may require manual whitelist:" -ForegroundColor Yellow
    Write-Host "  WPS -> Options -> Trust Center -> Add-in Security" -ForegroundColor Gray
    Write-Host "  Add 'DeepExcel' to the trusted add-ins list" -ForegroundColor Gray
}
