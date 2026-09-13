# Fresh-machine install acceptance test, run inside Windows Sandbox.
#
# Why this exists: v0.4.11 -> v0.4.17 were seven releases spent on installation
# and COM registration failures that no developer machine could reproduce,
# because a developer machine already has the registry entries, the runtimes,
# and no Mark-of-the-Web. Every release must be proven on a machine that has
# none of that.
#
# Two modes:
#   -Launch   (host)    builds the sandbox config and starts Windows Sandbox
#   default   (guest)   runs the actual checks inside the sandbox
#
# Usage on the host:
#   powershell -ExecutionPolicy Bypass -File scripts\verify-install-sandbox.ps1 -Launch
#
# The guest half is deliberately runnable on its own so it can also be pointed
# at a real clean VM, which is the only way to cover Excel-dependent steps that
# Windows Sandbox cannot (no Office inside the sandbox).

[CmdletBinding()]
param(
    [switch]$Launch,
    [string]$SetupPath,
    [string]$ResultPath = 'C:\DeepExcelVerify\result.txt'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Host side: build the .wsb and launch
# ---------------------------------------------------------------------------
function Start-SandboxRun {
    param([string]$Setup)

    $repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    if (-not $Setup) { $Setup = Join-Path $repoRoot 'dist\DeepExcel.Setup.exe' }
    if (-not (Test-Path $Setup)) {
        throw "Installer not found: $Setup. Run scripts\package_release.py first."
    }

    if (-not (Get-Command 'WindowsSandbox.exe' -ErrorAction SilentlyContinue)) {
        throw @'
Windows Sandbox is not available. Enable it with (admin PowerShell, then reboot):
  Enable-WindowsOptionalFeature -FeatureName "Containers-DisposableClientVM" -Online
Windows Sandbox requires Windows 10/11 Pro or Enterprise.
'@
    }

    # Everything the guest needs goes in one staged folder, mapped read-only.
    $stage = Join-Path ([IO.Path]::GetTempPath()) 'DeepExcelSandboxStage'
    if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    Copy-Item $Setup (Join-Path $stage 'DeepExcel.Setup.exe')
    Copy-Item $PSCommandPath (Join-Path $stage 'verify-install-sandbox.ps1')

    $sumFile = Join-Path (Split-Path -Parent $Setup) 'SHA256SUMS.txt'
    if (Test-Path $sumFile) { Copy-Item $sumFile (Join-Path $stage 'SHA256SUMS.txt') }

    $outDir = Join-Path $repoRoot 'dist\sandbox-verify'
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null

    $wsb = Join-Path $stage 'DeepExcelVerify.wsb'
    @"
<Configuration>
  <Networking>Enable</Networking>
  <MappedFolders>
    <MappedFolder>
      <HostFolder>$stage</HostFolder>
      <SandboxFolder>C:\DeepExcelStage</SandboxFolder>
      <ReadOnly>true</ReadOnly>
    </MappedFolder>
    <MappedFolder>
      <HostFolder>$outDir</HostFolder>
      <SandboxFolder>C:\DeepExcelVerify</SandboxFolder>
      <ReadOnly>false</ReadOnly>
    </MappedFolder>
  </MappedFolders>
  <LogonCommand>
    <Command>powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\DeepExcelStage\verify-install-sandbox.ps1 -SetupPath C:\DeepExcelStage\DeepExcel.Setup.exe</Command>
  </LogonCommand>
</Configuration>
"@ | Set-Content -LiteralPath $wsb -Encoding UTF8

    Write-Host "Launching Windows Sandbox..."
    Write-Host "  installer : $Setup"
    Write-Host "  results   : $outDir\result.txt"
    Write-Host ''
    Write-Host 'The sandbox window runs the checks and writes result.txt to the'
    Write-Host 'mapped folder above. Close the sandbox when it reports PASS/FAIL.'
    Start-Process 'WindowsSandbox.exe' -ArgumentList $wsb
}

# ---------------------------------------------------------------------------
# Guest side: the actual acceptance checks
# ---------------------------------------------------------------------------
$script:Results = New-Object System.Collections.Generic.List[object]

function Add-Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    $script:Results.Add([pscustomobject]@{ Name = $Name; Ok = $Ok; Detail = $Detail })
    $tag = if ($Ok) { '[PASS]' } else { '[FAIL]' }
    Write-Host "$tag $Name $Detail"
}

function Invoke-GuestVerification {
    param([string]$Setup, [string]$ResultFile)

    Write-Host '=== DeepExcel fresh-install verification ==='
    Write-Host "Setup: $Setup"
    Write-Host ''

    # 1. Integrity, exactly the way a user is told to check it.
    $sums = Join-Path (Split-Path -Parent $Setup) 'SHA256SUMS.txt'
    if (Test-Path $sums) {
        $actual = (Get-FileHash -Path $Setup -Algorithm SHA256).Hash.ToLowerInvariant()
        $expected = (Select-String -Path $sums -Pattern 'DeepExcel\.Setup\.exe' |
            Select-Object -First 1).Line -split '\s+' | Select-Object -First 1
        Add-Check 'SHA-256 matches SHA256SUMS.txt' ($actual -eq $expected.ToLowerInvariant()) "actual=$actual"
    } else {
        Add-Check 'SHA256SUMS.txt shipped alongside installer' $false 'file missing'
    }

    # 2. Silent install must succeed without admin. PrivilegesRequired=lowest,
    #    so any elevation prompt here is a packaging regression.
    Write-Host 'Installing (silent)...'
    $process = Start-Process -FilePath $Setup -ArgumentList '/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART' -Wait -PassThru
    Add-Check 'Installer exit code 0' ($process.ExitCode -eq 0) "exit=$($process.ExitCode)"

    $installDir = Join-Path $env:LOCALAPPDATA 'DeepExcel'
    Add-Check 'Install directory created' (Test-Path $installDir) $installDir

    foreach ($file in @('DeepExcel.AddIn.dll','DeepExcel.Repair.exe','DeepExcel.Probe32.exe','python\python.exe','sidecar\sidecar.py')) {
        Add-Check "Payload present: $file" (Test-Path (Join-Path $installDir $file))
    }

    # 3. No PowerShell scripts in the install directory. They were removed to
    #    cut an antivirus false-positive surface; a reappearance is a regression.
    $ps1 = @(Get-ChildItem -Path $installDir -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue)
    Add-Check 'No .ps1 in install directory' ($ps1.Count -eq 0) "found=$($ps1.Count)"

    # 4. Registration in BOTH registry views. The 32-bit view is the one that
    #    silently went missing for several releases.
    $clsid = '{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}'
    foreach ($view in @('Registry64','Registry32')) {
        $ok = $false
        $detail = ''
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
            $key = $base.OpenSubKey("Software\Classes\CLSID\$clsid\InprocServer32")
            if ($key) {
                $codeBase = $key.GetValue('CodeBase')
                $ok = $codeBase -and ($codeBase -like '*DeepExcel.AddIn.dll')
                $detail = "$codeBase"
                $key.Dispose()
            }
            $base.Dispose()
        } catch { $detail = $_.Exception.Message }
        Add-Check "COM registered in $view view" $ok $detail
    }

    $addinKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\DeepExcel.AddIn'
    $loadBehavior = if (Test-Path $addinKey) { (Get-ItemProperty $addinKey -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior } else { $null }
    Add-Check 'LoadBehavior = 3' ($loadBehavior -eq 3) "value=$loadBehavior"

    # 5. The repair tool's own verdict. This covers COM activation in both
    #    bitnesses, which is the only proof the whole chain actually works.
    $repair = Join-Path $installDir 'DeepExcel.Repair.exe'
    if (Test-Path $repair) {
        $verify = Start-Process -FilePath $repair -ArgumentList '--verify' -Wait -PassThru -NoNewWindow
        Add-Check 'DeepExcel.Repair.exe --verify reports healthy' ($verify.ExitCode -eq 0) "exit=$($verify.ExitCode)"
    } else {
        Add-Check 'DeepExcel.Repair.exe --verify reports healthy' $false 'tool missing'
    }

    # 6. Uninstall must leave no registration behind.
    $uninstaller = Get-ChildItem -Path $installDir -Filter 'unins*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($uninstaller) {
        Write-Host 'Uninstalling...'
        Start-Process -FilePath $uninstaller.FullName -ArgumentList '/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART' -Wait | Out-Null
        Start-Sleep -Seconds 3
        $residual = $false
        foreach ($view in @('Registry64','Registry32')) {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
            if ($base.OpenSubKey("Software\Classes\CLSID\$clsid")) { $residual = $true }
            $base.Dispose()
        }
        if (Test-Path $addinKey) { $residual = $true }
        Add-Check 'Uninstall removes COM registration' (-not $residual)
    } else {
        Add-Check 'Uninstall removes COM registration' $false 'uninstaller not found'
    }

    # ---- verdict -----------------------------------------------------------
    $failed = @($script:Results | Where-Object { -not $_.Ok })
    $verdict = if ($failed.Count -eq 0) { 'PASS' } else { 'FAIL' }

    $lines = @(
        "DeepExcel fresh-install verification: $verdict",
        "Timestamp: $(Get-Date -Format u)",
        "Checks: $($script:Results.Count), failed: $($failed.Count)",
        ''
    )
    foreach ($result in $script:Results) {
        $lines += ('{0} {1} {2}' -f $(if ($result.Ok) { '[PASS]' } else { '[FAIL]' }), $result.Name, $result.Detail)
    }

    $resultDir = Split-Path -Parent $ResultFile
    if ($resultDir -and -not (Test-Path $resultDir)) { New-Item -ItemType Directory -Path $resultDir -Force | Out-Null }
    $lines | Set-Content -LiteralPath $ResultFile -Encoding UTF8

    Write-Host ''
    Write-Host "=== $verdict ($($failed.Count) failed of $($script:Results.Count)) ==="
    Write-Host "Written to $ResultFile"

    # Windows Sandbox closes with the session; hold the window so a human can
    # read the outcome when running interactively.
    if ($env:COMPUTERNAME -and $Host.UI.RawUI) {
        Write-Host ''
        Write-Host 'Press Enter to close.'
        try { Read-Host | Out-Null } catch { }
    }

    if ($failed.Count -gt 0) { exit 1 }
}

# ---------------------------------------------------------------------------
if ($Launch) {
    Start-SandboxRun -Setup $SetupPath
} else {
    if (-not $SetupPath) { throw 'Specify -SetupPath, or use -Launch to run this inside Windows Sandbox.' }
    Invoke-GuestVerification -Setup $SetupPath -ResultFile $ResultPath
}
