# DeepExcel User-Scope Registration Script
# Registers the AddIn for the current user only (no admin required)

param(
    [switch]$Unregister = $false
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

# Wrap everything in try/catch so errors are displayed before the window
# closes (when launched via "Run with PowerShell" the window auto-closes
# on exit, hiding any error messages from the user).
try {

# STEP 0: Kill any running Excel processes BEFORE registering.
# If Excel is running during registration, it will NOT pick up the new
# add-in key until restarted. Many users forget to close Excel first,
# which is the #1 cause of "registered but not visible" issues.
$excelProcs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Host "Closing running Excel processes (required for registration)..." -ForegroundColor Yellow
    foreach ($ep in $excelProcs) {
        try { $ep.CloseMainWindow() | Out-Null } catch {}
    }
    Start-Sleep -Seconds 2
    # Force kill any that didn't close gracefully
    $remaining = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
    if ($remaining) {
        foreach ($ep in $remaining) {
            try { Stop-Process -Id $ep.Id -Force -ErrorAction SilentlyContinue } catch {}
        }
        Start-Sleep -Seconds 1
    }
    Write-Host "  Excel closed." -ForegroundColor Green
}

# Find DLL next to this script first (release package layout),
# then fall back to the dev folder structure.
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

# Remove Mark of the Web (MOTW) from downloaded files.
# DLLs extracted from a ZIP downloaded online carry an internet-zone mark;
# the Excel Trust Center silently blocks unsigned add-ins with this mark,
# so the ribbon tab never appears even when registration succeeds.
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

    # Write both 64-bit and 32-bit (WOW6432Node) registry views.
    # The AnyCPU managed DLL loads in both 32/64-bit Excel, but 32-bit Excel
    # reads HKCU\Software\Classes\WOW6432Node\CLSID while 64-bit reads CLSID.
    # Writing only one view makes the add-in invisible in the other Excel's
    # COM Add-ins list.
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

function Clear-ExcelResiliency {
    # Excel maintains a "blacklist" of add-ins that failed to load at
    # HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems.
    # Once an add-in is listed there, Excel will NOT attempt to reload it
    # even after the underlying DLL is fixed. This is the #1 reason users
    # see "registered successfully but not in COM Add-ins list".
    # We clear all entries that reference DeepExcel so Excel gets a fresh start.
    $disabledKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems"
    if (Test-Path $disabledKey) {
        $cleared = 0
        $items = Get-ChildItem $disabledKey -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            $val = (Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue)
            # DisabledItems entries store binary data; check all string properties for DeepExcel
            $isDeepExcel = $false
            foreach ($prop in $val.PSObject.Properties) {
                if ($prop.Value -is [string] -and $prop.Value -match "DeepExcel") {
                    $isDeepExcel = $true
                    break
                }
            }
            if ($isDeepExcel) {
                Remove-Item -Path $item.PSPath -Force -ErrorAction SilentlyContinue
                $cleared++
            }
        }
        if ($cleared -gt 0) {
            Write-Host "  Cleared $cleared DisabledItems entry(ies) for DeepExcel." -ForegroundColor Green
        }
    }
}

if ($Unregister) {
    Write-Host "[Unregister] DeepExcel.AddIn..." -ForegroundColor Yellow
    Unregister-ExcelAddIn -progId $progId
    Unregister-ComClass -clsid $clsid -progId $progId
    Unregister-ComClass -clsid $taskPaneClsid -progId $taskPaneProgId
    Clear-ExcelResiliency
    Write-Host "Unregistration complete!" -ForegroundColor Green
} else {
    Write-Host "[Register] DeepExcel.AddIn..." -ForegroundColor Yellow

    # Clear any prior "disabled" status before registering, so Excel
    # gets a clean attempt at loading the (now-fixed) add-in.
    Clear-ExcelResiliency

    Register-ComClass -clsid $clsid -progId $progId -dllPath $dllPath -className $addInClass
    Register-ExcelAddIn -progId $progId -friendlyName "DeepExcel AI AddIn" -dllPath $dllPath

    # Register TaskPaneControl (required by CustomTaskPane)
    Register-ComClass -clsid $taskPaneClsid -progId $taskPaneProgId -dllPath $dllPath -className $taskPaneClass
    Write-Host "TaskPaneControl registered (ProgID: $taskPaneProgId)" -ForegroundColor Gray

    # Add DoNotDisableAddinList entry: tells Excel "never auto-disable this add-in
    # even if it crashes on load". Without this, if the add-in fails to load once
    # (e.g. missing dependency), Excel will silently add it to DisabledItems and
    # it will never appear in the COM Add-ins list again. This is the #2 cause
    # of "registered but not visible" issues, especially on enterprise Office 365
    # which is more aggressive about disabling add-ins.
    $doNotDisableKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList"
    if (-not (Test-Path $doNotDisableKey)) {
        New-Item -Path $doNotDisableKey -Force | Out-Null
    }
    Set-ItemProperty -Path $doNotDisableKey -Name $progId -Value 1 -Type DWord -Force
    Write-Host "DoNotDisableAddinList entry added (prevents Excel from auto-disabling)" -ForegroundColor Gray

    # Post-registration verification: confirm key registry entries were written.
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

    # Environment diagnostics.
    Write-Host ""
    Write-Host "=== Environment Diagnostics ===" -ForegroundColor Yellow
    # .NET Framework 4.8 check.
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
    # Detect installed Excel bitness.
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
    # PowerShell bitness.
    $psBitness = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
    Write-Host "  [INFO] PowerShell: $psBitness" -ForegroundColor Gray

    # Check for Group Policy that might block add-ins.
    $gpKey = "HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Addins"
    if (Test-Path $gpKey) {
        Write-Host "  [WARN] Group Policy add-in restrictions detected at: $gpKey" -ForegroundColor Yellow
        Write-Host "         This may block user-installed add-ins. Contact your IT admin." -ForegroundColor Gray
    }
    $gpDisableKey = "HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Disabled"
    if (Test-Path $gpDisableKey) {
        Write-Host "  [WARN] Group Policy disabled-items key detected at: $gpDisableKey" -ForegroundColor Yellow
    }

    Write-Host ""
    if ($verifyOk) {
        Write-Host "Registration successful!" -ForegroundColor Green
    } else {
        Write-Host "Registration completed with ERRORS! See [FAIL] items above." -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  Excel has been closed. Launch Excel now to see the DeepExcel tab." -ForegroundColor Gray

    Write-Host ""
    Write-Host "If you don't see the DeepExcel tab:" -ForegroundColor Yellow
    Write-Host "  File -> Options -> Add-Ins -> Manage: COM Add-ins -> Go" -ForegroundColor Gray
    Write-Host "  Check 'DeepExcel.AddIn' -> OK" -ForegroundColor Gray
    Write-Host ""
    Write-Host "If it's NOT in the COM Add-ins list:" -ForegroundColor Yellow
    Write-Host "  - Make sure Excel is 2016/2019/365 (version 16.0)" -ForegroundColor Gray
    Write-Host "  - Close all Excel windows and rerun this script" -ForegroundColor Gray
    Write-Host "  - Check %APPDATA%\DeepExcel\logs\ for errors" -ForegroundColor Gray

    # Offer to launch Excel automatically.
    Write-Host ""
    $launchExcel = Read-Host "Launch Excel now? (Y/N, default Y)"
    if ($launchExcel -ne "N" -and $launchExcel -ne "n") {
        $excelExe = "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE"
        if (-not (Test-Path $excelExe)) {
            $excelExe = "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
        }
        if (Test-Path $excelExe) {
            Write-Host "Launching Excel..." -ForegroundColor Cyan
            Start-Process $excelExe
        } else {
            Write-Host "Excel not found in default path. Please launch manually." -ForegroundColor Yellow
        }
    }
}

} catch {
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "Details: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Full error:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
}

# Keep the window open when launched via "Run with PowerShell" right-click,
# which auto-closes the window on script exit. Pause so user can read output.
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Read-Host "Press Enter to close this window"

