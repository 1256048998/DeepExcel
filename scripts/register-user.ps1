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
Write-Host ""

$addInClass = "DeepExcel.AddIn.ThisAddIn"
$progId = "DeepExcel.AddIn"
$addinName = "DeepExcel.AddIn"

function Register-ComClass {
    param([string]$clsid, [string]$progId, [string]$dllPath, [string]$className)
    
    $hkcuClsid = "HKCU:\Software\Classes\CLSID\$clsid"
    $hkcuProgId = "HKCU:\Software\Classes\$progId"
    
    if (-not (Test-Path $hkcuClsid)) {
        New-Item -Path $hkcuClsid -Force | Out-Null
    }
    Set-ItemProperty -Path $hkcuClsid -Name "(default)" -Value $className -Force
    
    $inproc32 = Join-Path $hkcuClsid "InprocServer32"
    if (-not (Test-Path $inproc32)) {
        New-Item -Path $inproc32 -Force | Out-Null
    }
    Set-ItemProperty -Path $inproc32 -Name "(default)" -Value "mscoree.dll" -Force
    Set-ItemProperty -Path $inproc32 -Name "Assembly" -Value "DeepExcel.AddIn, Version=0.2.1.0, Culture=neutral, PublicKeyToken=null" -Force
    Set-ItemProperty -Path $inproc32 -Name "Class" -Value $className -Force
    Set-ItemProperty -Path $inproc32 -Name "CodeBase" -Value $dllPath -Force
    Set-ItemProperty -Path $inproc32 -Name "RuntimeVersion" -Value "v4.0.30319" -Force
    Set-ItemProperty -Path $inproc32 -Name "ThreadingModel" -Value "Both" -Force
    
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

function Unregister-ComClass {
    param([string]$clsid, [string]$progId)
    
    $hkcuClsid = "HKCU:\Software\Classes\CLSID\$clsid"
    $hkcuProgId = "HKCU:\Software\Classes\$progId"
    
    if (Test-Path $hkcuClsid) {
        Remove-Item -Path $hkcuClsid -Recurse -Force
    }
    if (Test-Path $hkcuProgId) {
        Remove-Item -Path $hkcuProgId -Recurse -Force
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
    
    Write-Host ""
    Write-Host "Registration successful!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Restart Excel" -ForegroundColor Gray
    Write-Host "  2. Find the 'DeepExcel' tab in the ribbon" -ForegroundColor Gray
    Write-Host "  3. Click 'Open Panel' button to start" -ForegroundColor Gray
    Write-Host ""
    Write-Host "If you don't see the DeepExcel tab:" -ForegroundColor Yellow
    Write-Host "  File -> Options -> Add-Ins -> Manage: COM Add-ins -> Go" -ForegroundColor Gray
    Write-Host "  Check 'DeepExcel.AddIn' -> OK" -ForegroundColor Gray
}
