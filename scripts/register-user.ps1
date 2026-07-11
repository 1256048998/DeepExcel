# DeepExcel User-Scope Registration Script
# Registers the AddIn for the current user only (no admin required)

param(
    [switch]$Unregister = $false
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

# ★ 优先查找脚本同目录下的 DLL（发布包场景：脚本和 DLL 在同一目录）
#   回退到开发场景的目录结构
$possiblePaths = @(
    (Join-Path $scriptDir "DeepExcel.AddIn.dll"),
    (Join-Path $scriptDir "bin\Release\DeepExcel.AddIn.dll"),
    (Join-Path $scriptDir "bin\Debug\DeepExcel.AddIn.dll"),
    (Join-Path (Split-Path -Parent $scriptDir) "src\DeepExcel.AddIn\bin\Release\DeepExcel.AddIn.dll"),
    (Join-Path (Split-Path -Parent $scriptDir) "src\DeepExcel.AddIn\bin\Debug\DeepExcel.AddIn.dll")
)

$dllPath = $null
foreach ($p in $possiblePaths) {
    if (Test-Path $p) { $dllPath = $p; break }
}

if (-not $dllPath) {
    Write-Host "ERROR: DeepExcel.AddIn.dll not found. Please build the project first." -ForegroundColor Red
    Write-Host "Searched paths:" -ForegroundColor Yellow
    foreach ($p in $possiblePaths) { Write-Host "  $p" -ForegroundColor Gray }
    exit 1
}

$dllPath = Resolve-Path $dllPath
Write-Host "AddIn DLL: $dllPath" -ForegroundColor Cyan

# ★ 解除 Mark of the Web (MOTW) 标记
#   从互联网下载的 ZIP 解压后，DLL 会带"来自互联网"标记，
#   Excel 信任中心会静默阻止加载此类未签名加载项，导致功能区选项卡不出现。
$scriptRoot = Split-Path -Parent $dllPath
Write-Host "Unblocking files (removing Mark of the Web)..." -ForegroundColor Gray
$unblocked = 0
Get-ChildItem -Path $scriptRoot -Recurse -File -Include *.dll,*.exe,*.ps1,*.config,*.py,*.html,*.js,*.css | ForEach-Object {
    if (Get-Item $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue) {
        try {
            Unblock-File -Path $_.FullName -ErrorAction Stop
            $unblocked++
        } catch {
            Write-Host "  WARN: failed to unblock $($_.Name)" -ForegroundColor Yellow
        }
    }
}
if ($unblocked -gt 0) {
    Write-Host "  Unblocked $unblocked file(s) downloaded from internet." -ForegroundColor Green
} else {
    Write-Host "  No blocked files found (OK)." -ForegroundColor Gray
}
Write-Host ""

$addInClass = "DeepExcel.AddIn.ThisAddIn"
$progId = "DeepExcel.AddIn"
$addinName = "DeepExcel.AddIn"

function Register-ComClass {
    param([string]$clsid, [string]$progId, [string]$dllPath, [string]$className)

    # ★ 同时写入 64 位和 32 位 (WOW6432Node) 视图
    #   AnyCPU 托管 DLL 可被 32/64 位 Excel 加载，但 32 位 Excel 读取
    #   HKCU\Software\Classes\WOW6432Node\CLSID，64 位读取 CLSID。
    #   若只写一处，另一位 Excel 在 COM 加载项列表里看不到此项。
    $clsidViews = @(
        "HKCU:\Software\Classes\CLSID\$clsid",
        "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsid"
    )
    $progIdViews = @(
        "HKCU:\Software\Classes\$progId",
        "HKCU:\Software\Classes\WOW6432Node\$progId"
    )

    foreach ($hkcuClsid in $clsidViews) {
        if (-not (Test-Path $hkcuClsid)) {
            New-Item -Path $hkcuClsid -Force | Out-Null
        }
        Set-ItemProperty -Path $hkcuClsid -Name "(default)" -Value $className -Force

        $inproc32 = Join-Path $hkcuClsid "InprocServer32"
        if (-not (Test-Path $inproc32)) {
            New-Item -Path $inproc32 -Force | Out-Null
        }
        Set-ItemProperty -Path $inproc32 -Name "(default)" -Value "mscoree.dll" -Force
        Set-ItemProperty -Path $inproc32 -Name "Assembly" -Value "DeepExcel.AddIn, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null" -Force
        Set-ItemProperty -Path $inproc32 -Name "Class" -Value $className -Force
        Set-ItemProperty -Path $inproc32 -Name "CodeBase" -Value $dllPath -Force
        Set-ItemProperty -Path $inproc32 -Name "RuntimeVersion" -Value "v4.0.30319" -Force
        Set-ItemProperty -Path $inproc32 -Name "ThreadingModel" -Value "Both" -Force
    }

    foreach ($hkcuProgId in $progIdViews) {
        if (-not (Test-Path $hkcuProgId)) {
            New-Item -Path $hkcuProgId -Force | Out-Null
        }
        Set-ItemProperty -Path $hkcuProgId -Name "(default)" -Value $className -Force

        $clsIdKey = Join-Path $hkcuProgId "CLSID"
        if (-not (Test-Path $clsIdKey)) {
            New-Item -Path $clsIdKey -Force | Out-Null
        }
        Set-ItemProperty -Path $clsIdKey -Name "(default)" -Value $clsid -Force
    }
}

function Unregister-ComClass {
    param([string]$clsid, [string]$progId)

    $clsidViews = @(
        "HKCU:\Software\Classes\CLSID\$clsid",
        "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsid"
    )
    $progIdViews = @(
        "HKCU:\Software\Classes\$progId",
        "HKCU:\Software\Classes\WOW6432Node\$progId"
    )
    foreach ($p in $clsidViews) {
        if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force }
    }
    foreach ($p in $progIdViews) {
        if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force }
    }
}

function Register-ExcelAddIn {
    param([string]$progId, [string]$friendlyName, [string]$dllPath)
    
    # Excel 16.0 (2016/2019/365) uses versioned path
    $addinKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Addins\$progId"
    
    if (-not (Test-Path $addinKey)) {
        New-Item -Path $addinKey -Force | Out-Null
    }
    
    Set-ItemProperty -Path $addinKey -Name "Description" -Value $friendlyName -Force
    Set-ItemProperty -Path $addinKey -Name "FriendlyName" -Value $friendlyName -Force
    Set-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -Force
    Set-ItemProperty -Path $addinKey -Name "CommandLineSafe" -Value 0 -Force
    Set-ItemProperty -Path $addinKey -Name "Location" -Value $dllPath -Force
}

function Unregister-ExcelAddIn {
    param([string]$progId)
    
    $addinKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Addins\$progId"
    
    if (Test-Path $addinKey) {
        Remove-Item -Path $addinKey -Recurse -Force
    }
}

$clsid = "{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}"
$taskPaneClsid = "{B2C3D4E5-F6A7-404B-9A5F-9B3C1D2E3F4B}"
$taskPaneProgId = "DeepExcel.AddIn.TaskPaneControl"
$taskPaneClass = "DeepExcel.AddIn.TaskPaneControl"

if ($Unregister) {
    Write-Host "[Unregister] DeepExcel.AddIn..." -ForegroundColor Yellow
    Unregister-ExcelAddIn -progId $progId
    Unregister-ComClass -clsid $clsid -progId $progId
    Unregister-ComClass -clsid $taskPaneClsid -progId $taskPaneProgId
    Write-Host "Unregistration complete!" -ForegroundColor Green
} else {
    Write-Host "[Register] DeepExcel.AddIn..." -ForegroundColor Yellow

    Register-ComClass -clsid $clsid -progId $progId -dllPath $dllPath -className $addInClass
    Register-ExcelAddIn -progId $progId -friendlyName "DeepExcel AI AddIn" -dllPath $dllPath

    # 注册 TaskPaneControl（CustomTaskPane 需要）
    Register-ComClass -clsid $taskPaneClsid -progId $taskPaneProgId -dllPath $dllPath -className $taskPaneClass
    Write-Host "TaskPaneControl registered (ProgID: $taskPaneProgId)" -ForegroundColor Gray

    # ★ 注册后验证：确认关键注册表项真实写入
    Write-Host ""
    Write-Host "=== Verification ===" -ForegroundColor Yellow
    $verifyOk = $true
    $addinKeyCheck = "HKCU:\Software\Microsoft\Office\16.0\Excel\Addins\$progId"
    if (Test-Path $addinKeyCheck) {
        $lb = (Get-ItemProperty $addinKeyCheck -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior
        Write-Host "  [OK] Excel Addin key (LoadBehavior=$lb)" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Excel Addin key missing!" -ForegroundColor Red
        $verifyOk = $false
    }
    $clsidPaths = @(
        "HKCU:\Software\Classes\CLSID\$clsid",
        "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsid"
    )
    foreach ($cp in $clsidPaths) {
        $tag = if ($cp -match "WOW6432Node") { "32-bit (WOW6432Node)" } else { "64-bit" }
        if (Test-Path $cp) {
            Write-Host "  [OK] CLSID $tag" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] CLSID $tag missing (may be OK if not needed)" -ForegroundColor Yellow
        }
    }
    if (Test-Path $dllPath) {
        Write-Host "  [OK] DLL exists: $dllPath" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] DLL missing!" -ForegroundColor Red
        $verifyOk = $false
    }

    # ★ 环境诊断信息
    Write-Host ""
    Write-Host "=== Environment Diagnostics ===" -ForegroundColor Yellow
    # .NET Framework 4.8 检测
    try {
        $ndpKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
        $release = (Get-ItemProperty $ndpKey -Name Release -ErrorAction Stop).Release
        if ($release -ge 462834) {
            Write-Host "  [OK] .NET Framework 4.8+ (release=$release)" -ForegroundColor Green
        } else {
            Write-Host "  [FAIL] .NET Framework 4.8 required (found release=$release)" -ForegroundColor Red
            $verifyOk = $false
        }
    } catch {
        Write-Host "  [WARN] Cannot determine .NET Framework version" -ForegroundColor Yellow
    }
    # 检测已安装 Excel 位数
    $excelPaths = @(
        @{ Name = "Excel 64-bit (Click-to-Run)"; Path = "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE" },
        @{ Name = "Excel 32-bit (Click-to-Run)"; Path = "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE" },
        @{ Name = "Excel 64-bit (MSI)"; Path = "${env:ProgramFiles}\Microsoft Office\Office16\EXCEL.EXE" },
        @{ Name = "Excel 32-bit (MSI)"; Path = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\EXCEL.EXE" }
    )
    $excelFound = $false
    foreach ($ep in $excelPaths) {
        if (Test-Path $ep.Path) {
            $excelFound = $true
            Write-Host "  [OK] $($ep.Name): $($ep.Path)" -ForegroundColor Green
        }
    }
    if (-not $excelFound) {
        Write-Host "  [WARN] Excel 2016+ not found in default paths (other version/location?)" -ForegroundColor Yellow
        Write-Host "         DeepExcel only supports Excel 2016/2019/365 (version 16.0)" -ForegroundColor Gray
    }
    # PowerShell 位数
    $psBitness = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
    Write-Host "  [INFO] PowerShell: $psBitness" -ForegroundColor Gray

    Write-Host ""
    if ($verifyOk) {
        Write-Host "Registration successful!" -ForegroundColor Green
    } else {
        Write-Host "Registration completed with ERRORS! See [FAIL] items above." -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Restart Excel (close ALL Excel windows)" -ForegroundColor Gray
    Write-Host "  2. Find the 'DeepExcel' tab in the ribbon" -ForegroundColor Gray
    Write-Host "  3. Click 'Open Panel' button to start" -ForegroundColor Gray
    Write-Host ""
    Write-Host "If you don't see the DeepExcel tab:" -ForegroundColor Yellow
    Write-Host "  File -> Options -> Add-Ins -> Manage: COM Add-ins -> Go" -ForegroundColor Gray
    Write-Host "  Check 'DeepExcel.AddIn' -> OK" -ForegroundColor Gray
    Write-Host ""
    Write-Host "If it's NOT in the COM Add-ins list:" -ForegroundColor Yellow
    Write-Host "  - Make sure Excel is 2016/2019/365 (version 16.0)" -ForegroundColor Gray
    Write-Host "  - Close all Excel windows and rerun this script" -ForegroundColor Gray
    Write-Host "  - Check %APPDATA%\DeepExcel\logs\ for errors" -ForegroundColor Gray
}
