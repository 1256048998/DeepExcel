# DeepExcel User-Scope Registration Script
# Registers the AddIn for the current user only (no admin required)

param(
    [switch]$Unregister = $false,
    [switch]$RepairOnly = $false,
    [switch]$ActivationOnly = $false
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

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

# Read Assembly version from DLL (replaces hardcoded 0.2.4.0)
try {
    $asmName = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath)
    $asmVersion = $asmName.Version.ToString()
    $assemblyValue = "DeepExcel.AddIn, Version=$asmVersion, Culture=neutral, PublicKeyToken=null"
    Write-Host "Assembly version: $asmVersion" -ForegroundColor Cyan
} catch {
    throw "Cannot read the DeepExcel assembly version from '$dllPath': $($_.Exception.Message)"
}

# Remove Mark of the Web (MOTW) from downloaded files.
# DLLs extracted from a ZIP downloaded online carry an internet-zone mark;
# the Excel Trust Center silently blocks unsigned add-ins with this mark,
# so the ribbon tab never appears even when registration succeeds.
$scriptRoot = Split-Path -Parent $dllPath
Write-Host "Unblocking files (removing Mark of the Web)..." -ForegroundColor Gray
$unblocked = 0
Get-ChildItem -Path $scriptRoot -Recurse -File -Include *.dll,*.exe,*.ps1,*.config,*.py,*.html,*.js,*.css,*.txt | ForEach-Object {
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
    param([string]$clsid, [string]$progId, [string]$dllPath, [string]$className, [string]$assemblyValue)

    # Open the actual registry views. A literal WOW6432Node path written by a
    # 64-bit process is not equivalent to Registry32 for HKCU\Software\Classes;
    # 32-bit Office then reports REGDB_E_CLASSNOTREG.
    $views = @(
        [Microsoft.Win32.RegistryView]::Registry64,
        [Microsoft.Win32.RegistryView]::Registry32
    )
    $dotNetCat = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
    $codeBase = ([Uri]$dllPath).AbsoluteUri
    $versionKeyName = ([System.Reflection.AssemblyName]::GetAssemblyName($dllPath)).Version.ToString()

    foreach ($view in $views) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $clsidPath = "Software\Classes\CLSID\$clsid"
            $inprocPath = "$clsidPath\InprocServer32"
            try { $base.DeleteSubKeyTree($inprocPath, $false) } catch { }

            $key = $base.CreateSubKey($clsidPath, $true)
            try { $key.SetValue('', $className, [Microsoft.Win32.RegistryValueKind]::String) } finally { $key.Dispose() }

            $key = $base.CreateSubKey($inprocPath, $true)
            try {
                $key.SetValue('', 'mscoree.dll', [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('Assembly', $assemblyValue, [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('Class', $className, [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('CodeBase', $codeBase, [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('RuntimeVersion', 'v4.0.30319', [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('ThreadingModel', 'Both', [Microsoft.Win32.RegistryValueKind]::String)
            } finally { $key.Dispose() }

            $key = $base.CreateSubKey("$inprocPath\$versionKeyName", $true)
            try {
                $key.SetValue('Assembly', $assemblyValue, [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('Class', $className, [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('CodeBase', $codeBase, [Microsoft.Win32.RegistryValueKind]::String)
                $key.SetValue('RuntimeVersion', 'v4.0.30319', [Microsoft.Win32.RegistryValueKind]::String)
            } finally { $key.Dispose() }

            $key = $base.CreateSubKey("$clsidPath\Implemented Categories\$dotNetCat", $true)
            $key.Dispose()
            $key = $base.CreateSubKey("$clsidPath\ProgId", $true)
            try { $key.SetValue('', $progId, [Microsoft.Win32.RegistryValueKind]::String) } finally { $key.Dispose() }

            $progIdPath = "Software\Classes\$progId"
            $key = $base.CreateSubKey($progIdPath, $true)
            try { $key.SetValue('', $className, [Microsoft.Win32.RegistryValueKind]::String) } finally { $key.Dispose() }
            $key = $base.CreateSubKey("$progIdPath\CLSID", $true)
            try { $key.SetValue('', $clsid, [Microsoft.Win32.RegistryValueKind]::String) } finally { $key.Dispose() }
        } finally {
            $base.Dispose()
        }
    }
}

function Unregister-ComClass {
    param([string]$clsid, [string]$progId)

    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            try { $base.DeleteSubKeyTree("Software\Classes\CLSID\$clsid", $false) } catch { }
            try { $base.DeleteSubKeyTree("Software\Classes\$progId", $false) } catch { }
        } finally { $base.Dispose() }
    }
}

function Register-ExcelAddIn {
    param([string]$progId, [string]$friendlyName, [string]$dllPath)

    # Office COM add-ins use the application-scoped, unversioned path.
    # Microsoft documents HKCU\Software\Microsoft\Office\Excel\Addins\<ProgID>.
    $addinKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"

    if (-not (Test-Path $addinKey)) {
        New-Item -Path $addinKey -Force | Out-Null
    }

    Set-ItemProperty -Path $addinKey -Name "Description" -Value $friendlyName -Force
    Set-ItemProperty -Path $addinKey -Name "FriendlyName" -Value $friendlyName -Force
    Set-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -Force
    Set-ItemProperty -Path $addinKey -Name "CommandLineSafe" -Value 0 -Force
    Set-ItemProperty -Path $addinKey -Name "Location" -Value $dllPath -Force

    # DoNotDisableAddinList: prevents Excel from soft-disabling this add-in on load failure
    # Without this, Excel sets LoadBehavior=2 after any startup error, hiding the add-in
    $resiliencyKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency"
    if (-not (Test-Path $resiliencyKey)) {
        New-Item -Path $resiliencyKey -Force | Out-Null
    }
    $doNotDisableKey = "$resiliencyKey\DoNotDisableAddinList"
    if (-not (Test-Path $doNotDisableKey)) {
        New-Item -Path $doNotDisableKey -Force | Out-Null
    }
    Set-ItemProperty -Path $doNotDisableKey -Name $progId -Value 1 -Force
    Write-Host "  Set DoNotDisableAddinList for $progId" -ForegroundColor Gray

    # Clean CrashingAddinList (another Excel blocklist besides DisabledItems)
    $crashingKey = "$resiliencyKey\CrashingAddinList"
    if (Test-Path $crashingKey) {
        try {
            $crashVal = (Get-ItemProperty $crashingKey -Name $progId -ErrorAction SilentlyContinue).$progId
            if ($null -ne $crashVal) {
                Remove-ItemProperty -Path $crashingKey -Name $progId -Force -ErrorAction Stop
                Write-Host "  Cleaned CrashingAddinList entry for $progId" -ForegroundColor Gray
            }
        } catch { }
    }

    # Clean only DisabledItems values that mention DeepExcel. DisabledItems stores
    # opaque binary registry values, not child keys; deleting the whole key would
    # re-enable unrelated add-ins and is therefore unsafe.
    $disabledKey = "$resiliencyKey\DisabledItems"
    if (Test-Path $disabledKey) {
        try {
            $values = Get-ItemProperty -Path $disabledKey -ErrorAction Stop
            foreach ($property in $values.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }) {
                $bytes = $property.Value
                if ($bytes -isnot [byte[]]) { continue }
                $unicode = [Text.Encoding]::Unicode.GetString($bytes)
                $ascii = [Text.Encoding]::ASCII.GetString($bytes)
                if (($unicode + ' ' + $ascii) -match 'DeepExcel(\.AddIn)?') {
                    Remove-ItemProperty -Path $disabledKey -Name $property.Name -Force -ErrorAction Stop
                    Write-Host "  Cleaned DisabledItems value: $($property.Name)" -ForegroundColor Gray
                }
            }
        } catch { }
    }
}

function Unregister-ExcelAddIn {
    param([string]$progId)

    $addinKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"

    if (Test-Path $addinKey) {
        Remove-Item -Path $addinKey -Recurse -Force
    }
    $legacyAddinKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Addins\$progId"
    if (Test-Path $legacyAddinKey) {
        Remove-Item -Path $legacyAddinKey -Recurse -Force
    }
    $doNotDisableKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList"
    if (Test-Path $doNotDisableKey) {
        Remove-ItemProperty -Path $doNotDisableKey -Name $progId -Force -ErrorAction SilentlyContinue
    }
}

function Test-ComActivation {
    param([string]$progId)
    try {
        $type = [Type]::GetTypeFromProgID($progId, $true)
        $obj = [Activator]::CreateInstance($type)
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
    } catch {
        throw "DeepExcel COM activation failed in $([IntPtr]::Size * 8)-bit process: $($_.Exception.Message)"
    }
}

$clsid = "{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}"
$taskPaneClsid = "{B2C3D4E5-F6A7-404B-9A5F-9B3C1D2E3F4B}"
$taskPaneProgId = "DeepExcel.AddIn.TaskPaneControl"
$taskPaneClass = "DeepExcel.AddIn.TaskPaneControl"

if ($ActivationOnly) {
    Test-ComActivation -progId $progId
    Write-Host "DeepExcel COM activation verified in $([IntPtr]::Size * 8)-bit process." -ForegroundColor Green
    exit 0
}

if ($RepairOnly) {
    Register-ComClass -clsid $clsid -progId $progId -dllPath $dllPath -className $addInClass -assemblyValue $assemblyValue
    Register-ComClass -clsid $taskPaneClsid -progId $taskPaneProgId -dllPath $dllPath -className $taskPaneClass -assemblyValue $assemblyValue
    Register-ExcelAddIn -progId $progId -friendlyName "DeepExcel AI AddIn" -dllPath $dllPath
    Test-ComActivation -progId $progId
    Write-Host "Excel COM registration, activation, and resiliency state verified." -ForegroundColor Green
    exit 0
}

if ($Unregister) {
    Write-Host "[Unregister] DeepExcel.AddIn..." -ForegroundColor Yellow
    Unregister-ExcelAddIn -progId $progId
    Unregister-ComClass -clsid $clsid -progId $progId
    Unregister-ComClass -clsid $taskPaneClsid -progId $taskPaneProgId
    # Also clean HKLM residuals from old installer (needs admin, ignore failures)
    foreach ($p in @(
        "HKLM:\SOFTWARE\Classes\CLSID\$clsid",
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\$clsid",
        "HKLM:\SOFTWARE\Classes\CLSID\$taskPaneClsid",
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\$taskPaneClsid",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Excel\Addins\$progId"
    )) {
        if (Test-Path $p) {
            try { Remove-Item -Path $p -Recurse -Force -ErrorAction Stop } catch { }
        }
    }
    Write-Host "Unregistration complete!" -ForegroundColor Green
} else {
    Write-Host "[Register] DeepExcel.AddIn..." -ForegroundColor Yellow

    # Clean HKLM residuals from old RegAsm installer.
    # CRITICAL: HKLM has HIGHER priority than HKCU for COM CLSID resolution.
    # If HKLM residual exists and points to an old DLL path, Excel will load the old DLL
    # regardless of HKCU settings. This was the root cause of v0.3.4-v0.4.15 all failing
    # to load on user machines (see project_memory.md).
    $hklmResiduals = @(
        "HKLM:\SOFTWARE\Classes\CLSID\$clsid",
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\$clsid",
        "HKLM:\SOFTWARE\Classes\CLSID\$taskPaneClsid",
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\$taskPaneClsid",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Excel\Addins\$progId"
    )
    $blockedResiduals = @()
    foreach ($p in $hklmResiduals) {
        if (Test-Path $p) {
            try {
                Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                Write-Host "  Cleaned HKLM residual: $p" -ForegroundColor Gray
            } catch {
                Write-Host "  BLOCKED: Cannot clean HKLM (admin needed): $p" -ForegroundColor Red
                $blockedResiduals += $p
            }
        }
    }
    if ($blockedResiduals.Count -gt 0) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "CRITICAL: HKLM residuals cannot be cleaned!" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "HKLM has HIGHER priority than HKCU for COM CLSID." -ForegroundColor Yellow
        Write-Host "Excel will load the OLD DLL from HKLM, ignoring the new DLL in HKCU." -ForegroundColor Yellow
        Write-Host "This is the root cause of the add-in not loading." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "SOLUTION: Right-click PowerShell -> Run as Administrator -> rerun this script." -ForegroundColor Green
        Write-Host ""
        Write-Host "Blocked paths:" -ForegroundColor Yellow
        foreach ($p in $blockedResiduals) { Write-Host "  $p" -ForegroundColor Gray }
        Write-Host ""
        Write-Host "Aborting registration to avoid false success." -ForegroundColor Red
        exit 2
    }

    Register-ComClass -clsid $clsid -progId $progId -dllPath $dllPath -className $addInClass -assemblyValue $assemblyValue
    Register-ExcelAddIn -progId $progId -friendlyName "DeepExcel AI AddIn" -dllPath $dllPath

    # Register TaskPaneControl (required by CustomTaskPane)
    Register-ComClass -clsid $taskPaneClsid -progId $taskPaneProgId -dllPath $dllPath -className $taskPaneClass -assemblyValue $assemblyValue
    Write-Host "TaskPaneControl registered (ProgID: $taskPaneProgId)" -ForegroundColor Gray

    # Post-registration verification: confirm key registry entries were written.
    Write-Host ""
    Write-Host "=== Verification ===" -ForegroundColor Yellow
    $verifyOk = $true
    $addinKeyCheck = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"
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
        if ($release -ge 528040) {
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

    # COM instantiation test: try creating COM object via ProgID
    Write-Host ""
    Write-Host "=== COM Instantiation Test ===" -ForegroundColor Yellow
    try {
        $type = [Type]::GetTypeFromProgID($progId)
        if ($type) {
            $obj = [Activator]::CreateInstance($type)
            Write-Host "  [OK] COM instantiation succeeded (ProgID: $progId)" -ForegroundColor Green

            # QI diagnostic: check if the managed type implements IDTExtensibility2
            try {
                $asm = [System.Reflection.Assembly]::LoadFrom($dllPath)
                $managedType = $asm.GetType("DeepExcel.AddIn.ThisAddIn")
                if ($managedType) {
                    $idtType = $managedType.GetInterface("IDTExtensibility2")
                    if ($idtType) {
                        Write-Host "  [OK] ThisAddIn implements IDTExtensibility2" -ForegroundColor Green
                        Write-Host "       Interface GUID: $($idtType.GUID)" -ForegroundColor Gray
                        Write-Host "       Interface Assembly: $($idtType.Assembly.GetName().Name) v$($idtType.Assembly.GetName().Version)" -ForegroundColor Gray
                    } else {
                        Write-Host "  [FAIL] ThisAddIn does NOT implement IDTExtensibility2!" -ForegroundColor Red
                    }
                    # List all COM interfaces
                    $ifaces = $managedType.GetInterfaces()
                    foreach ($iface in $ifaces) {
                        $guidAttrs = $iface.GetCustomAttributes([System.Runtime.InteropServices.GuidAttribute], $false)
                        if ($guidAttrs.Length -gt 0) {
                            Write-Host "       COM iface: $($iface.FullName) GUID={$($guidAttrs[0].Value)}" -ForegroundColor Gray
                        }
                    }
                } else {
                    Write-Host "  [WARN] Type 'DeepExcel.AddIn.ThisAddIn' not found in assembly" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "  [WARN] QI diagnostic failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }

            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch { }
        } else {
            Write-Host "  [WARN] ProgID not found in registry (may require Excel restart to take effect)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  [WARN] COM instantiation failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "         This can be normal if Excel is currently running. Close Excel and retry." -ForegroundColor Gray
    }

    # Display recent log file contents for diagnostics
    Write-Host ""
    Write-Host "=== Log File Check ===" -ForegroundColor Yellow
    $logPath = "$env:APPDATA\DeepExcel\logs\DeepExcel_Load.log"
    $tempLogPath = "$env:TEMP\DeepExcel_Load.log"
    foreach ($lp in @($logPath, $tempLogPath)) {
        if (Test-Path $lp) {
            $logContent = Get-Content $lp -Tail 20 -ErrorAction SilentlyContinue
            if ($logContent) {
                Write-Host "  Log: $lp" -ForegroundColor Gray
                foreach ($line in $logContent) {
                    Write-Host "    $line" -ForegroundColor DarkGray
                }
            }
        }
    }

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
