# DeepExcel User-Scope Registration Script
# Registers the AddIn for the current user only (no admin required)

param(
    [switch]$Unregister = $false
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

# Path compliance check: Chinese/space/parenthesis paths break .NET CLR CodeBase loading
if ($scriptDir -match '[\u4e00-\u9fff]') {
    Write-Host "ERROR: Path contains Chinese characters. .NET CLR cannot load assemblies from such paths." -ForegroundColor Red
    Write-Host "  Current path: $scriptDir" -ForegroundColor Yellow
    Write-Host "  Please move this folder to an ASCII-only path (e.g. C:\DeepExcel) and rerun." -ForegroundColor Yellow
    exit 1
}
if ($scriptDir -match '[\(\)]') {
    Write-Host "WARN: Path contains parentheses. This may cause CodeBase parsing issues." -ForegroundColor Yellow
    Write-Host "  Current path: $scriptDir" -ForegroundColor Gray
}
if ($scriptDir -like '*Desktop*' -or $scriptDir -like '*Downloads*') {
    Write-Host "WARN: Running from Desktop/Downloads is not recommended. Consider C:\DeepExcel." -ForegroundColor Yellow
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

# Read Assembly version from DLL (replaces hardcoded 0.2.4.0)
try {
    $asmName = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath)
    $asmVersion = $asmName.Version.ToString()
    $assemblyValue = "DeepExcel.AddIn, Version=$asmVersion, Culture=neutral, PublicKeyToken=null"
    Write-Host "Assembly version: $asmVersion" -ForegroundColor Cyan
} catch {
    Write-Host "WARN: Cannot read assembly version from DLL, using fallback 0.2.4.0" -ForegroundColor Yellow
    $assemblyValue = "DeepExcel.AddIn, Version=0.2.4.0, Culture=neutral, PublicKeyToken=null"
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

    # .NET Component Category GUID - RegAsm writes this; without it some Excel
    # builds refuse to enumerate the CLSID as a valid .NET COM component.
    $dotNetCat = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"

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
        Set-ItemProperty -Path $inproc32 -Name "Assembly" -Value $assemblyValue -Force
        Set-ItemProperty -Path $inproc32 -Name "Class" -Value $className -Force
        Set-ItemProperty -Path $inproc32 -Name "CodeBase" -Value $dllPath -Force
        Set-ItemProperty -Path $inproc32 -Name "RuntimeVersion" -Value "v4.0.30319" -Force
        Set-ItemProperty -Path $inproc32 -Name "ThreadingModel" -Value "Both" -Force

        # Implemented Categories (.NET) - missing on fresh systems breaks COM visibility
        $catPath = Join-Path $hkcuClsid "Implemented Categories\$dotNetCat"
        if (-not (Test-Path $catPath)) {
            New-Item -Path $catPath -Force | Out-Null
        }

        # ProgId reverse mapping (CLSID -> ProgID) - RegAsm writes this, manual reg must add it
        $progIdSubPath = Join-Path $hkcuClsid "ProgId"
        if (-not (Test-Path $progIdSubPath)) {
            New-Item -Path $progIdSubPath -Force | Out-Null
        }
        Set-ItemProperty -Path $progIdSubPath -Name "(default)" -Value $progId -Force
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

    # Clean DisabledItems (Excel's main blocklist, binary entries, scan and remove DeepExcel refs)
    $disabledKey = "$resiliencyKey\DisabledItems"
    if (Test-Path $disabledKey) {
        try {
            $items = Get-ChildItem $disabledKey -ErrorAction Stop
            foreach ($item in $items) {
                try {
                    $data = (Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue)
                    $allVals = $data.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object { "$($_.Name)=$($_.Value)" }
                    $joined = $allVals -join " "
                    if ($joined -match "DeepExcel|DeepExcel\.AddIn") {
                        Remove-Item -Path $item.PSPath -Force -ErrorAction Stop
                        Write-Host "  Cleaned DisabledItems entry (matched DeepExcel)" -ForegroundColor Gray
                    }
                } catch { }
            }
        } catch { }
    }
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

    # Clean HKLM residuals from old RegAsm installer (needs admin, ignore failures)
    # Prevents HKCU/HKLM pointing to different DLLs causing .NET CLR load conflict
    foreach ($p in @(
        "HKLM:\SOFTWARE\Classes\CLSID\$clsid",
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\$clsid",
        "HKLM:\SOFTWARE\Classes\CLSID\$taskPaneClsid",
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\$taskPaneClsid",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Excel\Addins\$progId"
    )) {
        if (Test-Path $p) {
            try {
                Remove-Item -Path $p -Recurse -Force -ErrorAction Stop
                Write-Host "  Cleaned HKLM residual: $p" -ForegroundColor Gray
            } catch {
                Write-Host "  WARN: Cannot clean HKLM (admin needed): $p" -ForegroundColor DarkGray
            }
        }
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
