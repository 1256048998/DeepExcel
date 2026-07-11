# DeepExcel Diagnostic Script
# Run: powershell -ExecutionPolicy Bypass -File diagnose.ps1
# This script checks why DeepExcel is not showing in Excel's COM Add-ins list.
# It does NOT modify anything - it only reads and reports.

$ErrorActionPreference = "SilentlyContinue"
$progId = "DeepExcel.AddIn"
$clsid = "{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}"

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " DeepExcel Diagnostic Report" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# 1. All Addin Registry Paths
Write-Host "=== 1. Addin Registry Paths ===" -ForegroundColor Yellow
$paths = @(
    @{ P = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"; D = "HKCU generic" },
    @{ P = "HKCU:\Software\Microsoft\Office\16.0\Excel\Addins\$progId"; D = "HKCU 16.0" },
    @{ P = "HKLM:\Software\Microsoft\Office\16.0\Excel\Addins\$progId"; D = "HKLM 64-bit" },
    @{ P = "HKLM:\Software\WOW6432Node\Microsoft\Office\16.0\Excel\Addins\$progId"; D = "HKLM 32-bit" }
)
foreach ($item in $paths) {
    if (Test-Path $item.P) {
        $lb = (Get-ItemProperty $item.P -Name LoadBehavior).LoadBehavior
        Write-Host "  [EXISTS] $($item.D) (LoadBehavior=$lb)" -ForegroundColor Green
    } else {
        Write-Host "  [NONE]   $($item.D)" -ForegroundColor Gray
    }
}

# 2. Resiliency\DisabledItems
Write-Host ""
Write-Host "=== 2. Excel Disabled Items (Resiliency) ===" -ForegroundColor Yellow
$disabledKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems"
if (Test-Path $disabledKey) {
    $items = Get-ChildItem $disabledKey
    if ($items) {
        foreach ($i in $items) {
            $name = (Get-ItemProperty $i.PSPath)."(default)"
            Write-Host "  [DISABLED] $name" -ForegroundColor Red
        }
        Write-Host "  >>> Excel has DISABLED some items (see above) <<<" -ForegroundColor Red
    } else {
        Write-Host "  [OK] DisabledItems key is empty" -ForegroundColor Green
    }
} else {
    Write-Host "  [OK] No DisabledItems key (good)" -ForegroundColor Green
}

# 3. Group Policy
Write-Host ""
Write-Host "=== 3. Group Policy Check ===" -ForegroundColor Yellow
$policyPaths = @(
    "HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Addins",
    "HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Addins",
    "HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Options",
    "HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Options"
)
$policyFound = $false
foreach ($p in $policyPaths) {
    if (Test-Path $p) {
        $props = Get-ItemProperty $p
        $customProps = $props.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" }
        if ($customProps) {
            $policyFound = $true
            Write-Host "  [POLICY] $p" -ForegroundColor Red
            foreach ($prop in $customProps) {
                Write-Host "    $($prop.Name) = $($prop.Value)" -ForegroundColor Red
            }
        }
    }
}
if (-not $policyFound) {
    Write-Host "  [OK] No Office group policies found" -ForegroundColor Green
}

# 4. Trust Center Security Settings
Write-Host ""
Write-Host "=== 4. Trust Center Security (Addin-related) ===" -ForegroundColor Yellow
$secPaths = @(
    @{ P = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"; D = "User Security" },
    @{ P = "HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Security"; D = "Policy Security" },
    @{ P = "HKCU:\Software\Microsoft\Office\16.0\Common\Security"; D = "Common Security" }
)
$secIssue = $false
foreach ($item in $secPaths) {
    if (Test-Path $item.P) {
        $props = Get-ItemProperty $item.P
        $customProps = $props.PSObject.Properties | Where-Object { $_.Name -match "Addin|Sig|Macro|Disable|Require" -and $_.Name -notmatch "^PS" }
        if ($customProps) {
            Write-Host "  [$($item.D)]:" -ForegroundColor Yellow
            foreach ($prop in $customProps) {
                $color = if ($prop.Value -ne 0) { "Red" } else { "Gray" }
                if ($prop.Value -ne 0) { $secIssue = $true }
                Write-Host "    $($prop.Name) = $($prop.Value)" -ForegroundColor $color
            }
        }
    }
}
if (-not $secIssue) {
    Write-Host "  [OK] No restrictive security settings found" -ForegroundColor Green
} else {
    Write-Host "  >>> Some security settings are NON-ZERO (may block addin) <<<" -ForegroundColor Red
}

# 5. COM Instantiation Test (THE CRITICAL TEST)
Write-Host ""
Write-Host "=== 5. COM Instantiation Test (CRITICAL) ===" -ForegroundColor Yellow
Write-Host "  Trying to create COM object: $progId" -ForegroundColor Gray
$comOk = $false
try {
    $obj = New-Object -ComObject $progId -ErrorAction Stop
    $comOk = $true
    Write-Host "  [OK] COM instantiation SUCCESS!" -ForegroundColor Green
    Write-Host "  >>> DLL loads fine. Problem is Excel policy/trust. <<<" -ForegroundColor Green
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
} catch {
    Write-Host "  [FAIL] COM instantiation FAILED!" -ForegroundColor Red
    Write-Host "  >>> DLL cannot be loaded. See error below. <<<" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "  InnerException: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
}

# 6. Excel Process Info
Write-Host ""
Write-Host "=== 6. Excel Process ===" -ForegroundColor Yellow
$excelProc = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Host "  [RUNNING] Excel is currently running (PID: $($excelProc.Id -join ', '))" -ForegroundColor Yellow
    Write-Host "  >>> Close ALL Excel windows before testing <<<" -ForegroundColor Yellow
} else {
    Write-Host "  [OK] Excel is not running" -ForegroundColor Green
}

# 7. DLL Dependencies Check
Write-Host ""
Write-Host "=== 7. DLL File Check ===" -ForegroundColor Yellow
$scriptDir = $PSScriptRoot
$dllPath = Join-Path $scriptDir "DeepExcel.AddIn.dll"
if (Test-Path $dllPath) {
    $size = (Get-Item $dllPath).Length
    Write-Host "  [OK] DeepExcel.AddIn.dll exists ($([math]::Round($size/1024,1)) KB)" -ForegroundColor Green
    # Check MOTW
    $zone = Get-Item $dllPath -Stream Zone.Identifier -ErrorAction SilentlyContinue
    if ($zone) {
        Write-Host "  [WARN] DLL has Mark of the Web (internet zone mark)!" -ForegroundColor Red
        Write-Host "  >>> Run register-user.ps1 to unblock, or right-click ZIP -> Properties -> Unblock <<<" -ForegroundColor Red
    } else {
        Write-Host "  [OK] DLL is not blocked (no MOTW)" -ForegroundColor Green
    }
} else {
    Write-Host "  [FAIL] DeepExcel.AddIn.dll NOT FOUND in script directory!" -ForegroundColor Red
}

# 8. DeepExcel Log
Write-Host ""
Write-Host "=== 8. DeepExcel Log ===" -ForegroundColor Yellow
$logDir = "$env:APPDATA\DeepExcel\logs"
if (Test-Path $logDir) {
    $latest = Get-ChildItem $logDir -Filter *.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latest) {
        Write-Host "  [FOUND] Latest log: $($latest.Name) ($($latest.LastWriteTime))" -ForegroundColor Gray
        Write-Host "  Last 15 lines:" -ForegroundColor Gray
        Get-Content $latest.FullName -Tail 15 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    } else {
        Write-Host "  [NONE] No log files (Excel may never have loaded the addin)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  [NONE] No log directory (Excel never started the addin)" -ForegroundColor Yellow
}

# Summary
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Summary" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Please send a SCREENSHOT of this entire output back." -ForegroundColor White
Write-Host "The section '5. COM Instantiation Test' is the most important." -ForegroundColor White
Write-Host ""
Write-Host "  If section 5 = [OK]    -> DLL is fine, Excel policy is blocking" -ForegroundColor Gray
Write-Host "  If section 5 = [FAIL]  -> DLL cannot load, see error message" -ForegroundColor Gray
Write-Host "  If section 2 = [DISABLED] -> Excel disabled the addin (needs cleanup)" -ForegroundColor Gray
Write-Host "  If section 4 = non-zero  -> Trust Center is blocking unsigned addins" -ForegroundColor Gray
Write-Host ""
