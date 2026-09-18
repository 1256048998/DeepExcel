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
    [switch]$Trace,
    [switch]$WithWps,
    # Drive WPS with synthetic keystrokes. Off by default: it races the
    # add-in's initialisation and produced a ribbon with no DeepExcel tab.
    [switch]$AutoKeys,
    # Tests whether WPS can load the Excel COM add-in directly, the way FFCell
    # ships: one VSTO/COM add-in serving both hosts, no JS add-in at all.
    # Disables the JS add-in first so any ribbon tab that appears can only have
    # come from COM.
    [switch]$ComWps,
    [string]$SetupPath,
    # Path to a WPS installer on the host. Not downloaded automatically: the
    # sandbox is disposable, so an automatic download would re-fetch several
    # hundred MB on every run, and the download URL is not a stable contract.
    [string]$WpsSetupPath,
    [string]$ResultPath = 'C:\DeepExcelVerify\result.txt',
    # Skip the "Press Enter to close" hold at the end. Without this the sandbox
    # window sits there waiting for a keystroke long after the result file has
    # been written, so an unattended run -- a release pipeline, or anyone who
    # started it and walked away -- never actually finishes.
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Host side: build the .wsb and launch
# ---------------------------------------------------------------------------
function Start-SandboxRun {
    param([string]$Setup, [bool]$WithTrace, [string]$WpsSetup, [bool]$ComWpsTest,
          [bool]$AutoKeysTest, [bool]$NoPauseRun)

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

    $traceArg = ''
    if ($WithTrace) { $traceArg = ' -Trace' }
    $pauseArg = ''
    if ($NoPauseRun) { $pauseArg = ' -NoPause' }

    $keysArg = ''
    if ($AutoKeysTest) { $keysArg = ' -AutoKeys' }

    $comArg = ''
    if ($ComWpsTest) { $comArg = ' -ComWps' }

    $wpsArg = ''
    if ($WpsSetup) {
        if (-not (Test-Path $WpsSetup)) { throw "WPS installer not found: $WpsSetup" }
        Copy-Item $WpsSetup (Join-Path $stage 'WpsSetup.exe')
        $wpsArg = ' -WithWps'
        Write-Host "  WPS setup : $WpsSetup"
    }

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
    <Command>powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\DeepExcelStage\verify-install-sandbox.ps1 -SetupPath C:\DeepExcelStage\DeepExcel.Setup.exe$traceArg$wpsArg$comArg$keysArg$pauseArg</Command>
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

# Process Monitor tells us which file open actually returns NAME NOT FOUND.
# COM's 0x80070002 names no file, and Fusion logged no bind attempt at all for
# the add-in, so the failure is below the CLR -- file-level tracing is the only
# thing left that can name it.
#
# The trace lives on the VM's own disk, never in the mapped folder: a raw PML
# runs to hundreds of MB and the mapped folder is the host's disk. Only the
# filtered summary is exported. Everything else dies with the VM.
function Invoke-ActivationTrace {
    param([string]$Repair, [string]$OutDir)

    $tools = 'C:\pmtools'
    $traceDir = 'C:\pmtrace'
    New-Item -ItemType Directory -Path $tools, $traceDir -Force | Out-Null

    # Every failure here goes to the mapped folder. The previous run produced no
    # trace and no explanation, because these messages only went to a console
    # that dies with the VM.
    $log = Join-Path $OutDir 'trace-setup.log'
    function Write-TraceLog([string]$m) {
        Add-Content -LiteralPath $log -Value ('{0}  {1}' -f (Get-Date -Format 'HH:mm:ss'), $m)
        Write-Host "  (trace) $m"
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $zip = Join-Path $tools 'ProcessMonitor.zip'
        Write-TraceLog 'downloading ProcessMonitor.zip from download.sysinternals.com'
        Invoke-WebRequest -Uri 'https://download.sysinternals.com/files/ProcessMonitor.zip' `
                          -OutFile $zip -UseBasicParsing
        Write-TraceLog ('downloaded {0:N0} bytes' -f (Get-Item $zip).Length)
        Expand-Archive -LiteralPath $zip -DestinationPath $tools -Force
        Write-TraceLog ('extracted: ' + ((Get-ChildItem $tools -Filter '*.exe' | ForEach-Object Name) -join ', '))
    } catch {
        Write-TraceLog ("download/extract FAILED: " + $_.Exception.Message)
        return
    }

    $procmon = Join-Path $tools 'Procmon64.exe'
    if (-not (Test-Path $procmon)) { $procmon = Join-Path $tools 'Procmon.exe' }
    if (-not (Test-Path $procmon)) { Write-TraceLog 'Procmon executable not found in archive'; return }
    Write-TraceLog "using $procmon"

    $pml = Join-Path $traceDir 'trace.pml'
    Start-Process -FilePath $procmon -ArgumentList '/AcceptEula','/Quiet','/Minimized','/BackingFile',$pml | Out-Null
    Start-Sleep -Seconds 8

    Start-Process -FilePath $Repair -ArgumentList '--verify' -Wait -NoNewWindow `
        -RedirectStandardOutput (Join-Path $OutDir 'repair-verify-traced.log') `
        -RedirectStandardError  (Join-Path $OutDir 'repair-verify-traced.log.err') | Out-Null
    Start-Sleep -Seconds 3

    Start-Process -FilePath $procmon -ArgumentList '/Terminate' -Wait | Out-Null
    Start-Sleep -Seconds 5

    $csv = Join-Path $traceDir 'trace.csv'
    Start-Process -FilePath $procmon -ArgumentList '/OpenLog',$pml,'/SaveAs',$csv -Wait | Out-Null

    if (-not (Test-Path $csv)) { Write-TraceLog 'CSV export produced nothing'; return }
    Write-TraceLog ('CSV exported: {0:N0} bytes' -f (Get-Item $csv).Length)

    # Only the interesting rows leave the VM.
    $rows = Import-Csv -LiteralPath $csv
    # Widened deliberately: COM activation fails before the CLR is involved, so
    # the interesting misses may be mscoree/registry lookups, not files under
    # the install directory. Filtering only on the add-in's own name could hide
    # the very event being hunted.
    $miss = $rows | Where-Object {
        $_.Result -match 'NAME NOT FOUND|PATH NOT FOUND|ACCESS DENIED' -and
        ($_.'Process Name' -match 'DeepExcel|Probe32|Repair' -or $_.Path -match 'DeepExcel|mscoree')
    }
    $summary = Join-Path $OutDir 'activation-trace.log'
    $out = @("Total traced events: $($rows.Count)",
             "DeepExcel NAME/PATH NOT FOUND events: $($miss.Count)",
             '')
    $out += ($miss | Select-Object -First 400 |
             ForEach-Object { '{0}  {1}  {2}  -> {3}' -f $_.'Process Name', $_.Operation, $_.Path, $_.Result })
    $out | Set-Content -LiteralPath $summary -Encoding UTF8
    Write-Host "  (trace) $($miss.Count) miss events written to $summary"
}

# Installs WPS inside the sandbox so the WPS half of the acceptance test runs
# against a real host application instead of just asserting on files.
#
# This is the only way to answer the open question: the installer writes
# enable="enable_dev" into publish.xml, and whether a normal (non-developer)
# WPS actually loads an enable_dev entry has never been verified. The ribbon
# either shows the DeepExcel tab or it does not -- that needs a human looking
# at the sandbox window, so this function ends by launching WPS.
function Install-WpsInGuest {
    # Everything here is logged to the mapped folder. The previous run skipped
    # WPS and left no trace of why -- the sandbox console dies with the VM, so a
    # step with no artifact is a step that cannot be diagnosed.
    $log = 'C:\DeepExcelVerify\wps-install.log'
    $dir = Split-Path -Parent $log
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    function Write-WpsLog([string]$Message) {
        $line = '{0}  {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message
        Add-Content -LiteralPath $log -Value $line
        Write-Host "  (wps) $Message"
    }

    $setup = 'C:\DeepExcelStage\WpsSetup.exe'
    Write-WpsLog "staged installer present: $(Test-Path $setup)"
    if (-not (Test-Path $setup)) { return $false }
    Write-WpsLog ('installer size: {0:N0} bytes' -f (Get-Item $setup).Length)

    # Try the common silent switches in order. WPS is not a plain NSIS package,
    # so /S alone is not guaranteed; each attempt gets a short window to show
    # progress before moving on.
    $searchPaths = @('C:\Program Files (x86)\Kingsoft', 'C:\Program Files\Kingsoft',
                     "$env:LOCALAPPDATA\Kingsoft", "$env:ProgramData\Kingsoft")

    function Find-Et {
        param([string[]]$Paths)
        foreach ($p in $Paths) {
            if (Test-Path $p) {
                $hit = Get-ChildItem -Path $p -Recurse -Filter 'et.exe' -ErrorAction SilentlyContinue |
                       Select-Object -First 1
                if ($hit) { return $hit }
            }
        }
        return $null
    }

    foreach ($switches in @('/S', '/silent', '/quiet /norestart', '')) {
        $label = if ($switches) { $switches } else { '(no switches, interactive)' }
        Write-WpsLog "launching installer with: $label"
        try {
            if ($switches) {
                Start-Process -FilePath $setup -ArgumentList $switches -PassThru | Out-Null
            } else {
                Start-Process -FilePath $setup -PassThru | Out-Null
            }
        } catch {
            Write-WpsLog "launch failed: $($_.Exception.Message)"
            continue
        }

        # Interactive attempt gets much longer: a human has to click through it.
        $minutes = if ($switches) { 4 } else { 15 }
        $deadline = (Get-Date).AddMinutes($minutes)
        while ((Get-Date) -lt $deadline) {
            $found = Find-Et -Paths $searchPaths
            if ($found) {
                Write-WpsLog "installed: $($found.FullName)"
                Set-Content -LiteralPath 'C:\wps-et-path.txt' -Value $found.FullName -Encoding ASCII
                return $true
            }
            Start-Sleep -Seconds 10
        }
        Write-WpsLog "et.exe not found after $minutes min with: $label"
        $procs = @(Get-Process -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -match 'wps|et|setup|ksolaunch' } |
                   ForEach-Object { $_.Name })
        Write-WpsLog ("running related processes: " + ($procs -join ', '))
    }

    Write-WpsLog 'giving up on WPS install; WPS checks will be file-level only'
    return $false
}

# Screenshots taken INSIDE the VM, written to the mapped folder.
#
# Whether a ribbon tab appears is a visual fact, and nothing outside the VM can
# see the sandbox window. Capturing from within the guest removes the human from
# the loop entirely and never touches the host's mouse or keyboard.
function Save-GuestScreenshot {
    param([string]$OutDir, [string]$Name)
    try {
        # Captured in a CHILD process on purpose.
        #
        # This script is started by the sandbox LogonCommand, before WPS is
        # installed. Installing WPS churns the registry underneath it, and the
        # long-lived process is then unable to load any new assembly at all --
        # both System.Windows.Forms and System.Drawing failed with 0x800703FA
        # ("registry key marked for deletion"), and retrying inside the same
        # process never recovers. A freshly spawned powershell.exe gets a clean
        # assembly-load context and succeeds.
        $path = Join-Path $OutDir ($Name + '.png')
        $capture = @"
Add-Type -AssemblyName System.Drawing
`$w = 1920; `$h = 1080
try {
  `$vc = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue |
        Where-Object { `$_.CurrentHorizontalResolution } | Select-Object -First 1
  if (`$vc) { `$w = [int]`$vc.CurrentHorizontalResolution; `$h = [int]`$vc.CurrentVerticalResolution }
} catch { }
`$bmp = New-Object System.Drawing.Bitmap `$w, `$h
`$gfx = [System.Drawing.Graphics]::FromImage(`$bmp)
`$gfx.CopyFromScreen(0, 0, 0, 0, `$bmp.Size)
`$bmp.Save('$path', [System.Drawing.Imaging.ImageFormat]::Png)
`$gfx.Dispose(); `$bmp.Dispose()
Write-Output ('{0}x{1}' -f `$w, `$h)
"@
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($capture))
        $stdout = Join-Path $OutDir ('.' + $Name + '.out')
        Start-Process -FilePath 'powershell.exe' `
            -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encoded `
            -Wait -NoNewWindow -RedirectStandardOutput $stdout -RedirectStandardError "$stdout.err" | Out-Null

        if (Test-Path $path) {
            $size = (Get-Content -LiteralPath $stdout -ErrorAction SilentlyContinue | Select-Object -First 1)
            Write-Host "  (shot) $path"
            Add-Content -LiteralPath (Join-Path $OutDir 'screenshot.log') -Value ("OK   {0}  {1}" -f $Name, $size)
        } else {
            $err = (Get-Content -LiteralPath "$stdout.err" -Raw -ErrorAction SilentlyContinue)
            Add-Content -LiteralPath (Join-Path $OutDir 'screenshot.log') -Value ("FAIL {0}  {1}" -f $Name, $err)
        }
        Remove-Item $stdout, "$stdout.err" -ErrorAction SilentlyContinue
    } catch {
        # Console output dies with the VM, so the reason has to reach the
        # mapped folder or it is lost.
        Write-Host "  (shot) failed: $($_.Exception.Message)"
        try {
            Add-Content -LiteralPath (Join-Path $OutDir 'screenshot.log') `
                        -Value ("FAIL {0}  {1}" -f $Name, $_.Exception.ToString())
        } catch { }
    }
}

# Keystrokes into the guest's foreground window, via a child process for the
# same assembly-load reason as the screenshots. WPS opens on a login dialog and
# a "what's new" tab, so the ribbon does not exist until those are dismissed --
# that is why the first automated attempt found no ribbon elements at all.
function Send-GuestKeys {
    param([string]$Keys, [int]$SleepMs = 1500)
    $script = @"
`$w = New-Object -ComObject WScript.Shell
Start-Sleep -Milliseconds 300
`$w.SendKeys('$Keys')
Start-Sleep -Milliseconds $SleepMs
"@
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($script))
    Start-Process -FilePath 'powershell.exe' `
        -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encoded `
        -Wait -NoNewWindow | Out-Null
}

# Dumps the WPS window's automation tree so element names can be read from the
# host instead of guessed. Guessing "the button is called X" already cost one
# run.
function Save-GuestUiTree {
    param([string]$OutDir, [string]$Name)
    try {
        $ok = $false
        for ($try = 1; $try -le 5 -and -not $ok; $try++) {
            try { Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes -ErrorAction Stop; $ok = $true }
            catch { Start-Sleep -Seconds 5 }
        }
        if (-not $ok) { throw 'UIAutomation assemblies could not be loaded' }
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                             [System.Windows.Automation.Condition]::TrueCondition)
        $lines = @("elements: $($all.Count)")
        foreach ($e in $all) {
            try {
                $n = $e.Current.Name
                $c = $e.Current.ControlType.ProgrammaticName
                if ($n) { $lines += ('{0}  [{1}]' -f $n, $c) }
            } catch { }
        }
        $lines | Set-Content -LiteralPath (Join-Path $OutDir ($Name + '.txt')) -Encoding UTF8
        Write-Host "  (ui) tree dumped: $Name ($($all.Count) elements)"
    } catch {
        try {
            Add-Content -LiteralPath (Join-Path $OutDir 'screenshot.log') `
                        -Value ("FAIL uitree  " + $_.Exception.ToString())
        } catch { }
    }
}

# Clicks the ribbon button from inside the guest via UI Automation, so the
# "does the panel open" question does not need a human either. Host input is
# never touched -- this drives the VM's own UI tree.
function Invoke-GuestRibbonClick {
    param([string]$AutomationName)
    try {
        $ok = $false
        for ($try = 1; $try -le 5 -and -not $ok; $try++) {
            try { Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes -ErrorAction Stop; $ok = $true }
            catch { Start-Sleep -Seconds 5 }
        }
        if (-not $ok) { throw 'UIAutomation assemblies could not be loaded' }
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $cond = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, $AutomationName)
        $deadline = (Get-Date).AddSeconds(30)
        $el = $null
        while (-not $el -and (Get-Date) -lt $deadline) {
            $el = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
            if (-not $el) { Start-Sleep -Seconds 2 }
        }
        if (-not $el) { Write-Host "  (ui) element not found: $AutomationName"; return $false }
        $pattern = $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        $pattern.Invoke()
        Write-Host "  (ui) invoked: $AutomationName"
        return $true
    } catch {
        Write-Host "  (ui) failed on '$AutomationName': $($_.Exception.Message)"
        return $false
    }
}

function Invoke-GuestVerification {
    param([string]$Setup, [string]$ResultFile, [bool]$WithTrace, [bool]$WithWps,
          [bool]$ComWps, [bool]$AutoKeys, [bool]$NoPause)

    Write-Host '=== DeepExcel fresh-install verification ==='
    Write-Host "Setup: $Setup"
    Write-Host ''

    # WPS goes first so DeepExcel installs into a machine that already has WPS,
    # which is the real user's order. Installing it afterwards would let the
    # DeepExcel installer create jsaddons itself and hide any ordering problem.
    $wpsReady = $false
    if ($WithWps) { $wpsReady = Install-WpsInGuest }

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
    #
    # /LOG writes into the mapped (writable) folder, not the guest's %TEMP%.
    # Without it a FAIL is undiagnosable: the sandbox is disposable, so the
    # Inno log dies with the VM and all that survives is "exit=0, no keys".
    # The log is the only artifact that says which [Code] step actually ran.
    $outDir = Split-Path -Parent $ResultFile
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $installLog = Join-Path $outDir 'install.log'

    Write-Host 'Installing (silent)...'
    $process = Start-Process -FilePath $Setup -ArgumentList '/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART',"/LOG=$installLog" -Wait -PassThru
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
        # Capture the tool's own diagnosis, not just its exit code -- "exit=1"
        # alone does not say which part of the chain it found broken.
        $repairOut = Join-Path $outDir 'repair-verify.log'
        $verify = Start-Process -FilePath $repair -ArgumentList '--verify' -Wait -PassThru -NoNewWindow `
            -RedirectStandardOutput $repairOut -RedirectStandardError "$repairOut.err"
        Add-Check 'DeepExcel.Repair.exe --verify reports healthy' ($verify.ExitCode -eq 0) "exit=$($verify.ExitCode)"

        # When verification fails, re-run the exact step the installer runs
        # (--repair) with output captured. The installer invokes it with
        # --quiet, so its reasoning is lost precisely when it is needed. This
        # is diagnosis only -- it does not affect any check's verdict.
        if ($verify.ExitCode -ne 0) {
            $repairRun = Join-Path $outDir 'repair-run.log'
            Start-Process -FilePath $repair -ArgumentList '--repair' -Wait -NoNewWindow `
                -RedirectStandardOutput $repairRun -RedirectStandardError "$repairRun.err" | Out-Null
            Write-Host "  (diagnostic) --repair output written to $repairRun"

            # Separate "registration is wrong" from "a dependency is missing":
            # COM activation returning 0x80070002 says only "some file was not
            # found", which is equally consistent with a bad CodeBase path and
            # with an assembly the add-in needs at load time being absent.
            $dump = Join-Path $outDir 'registration-dump.log'
            $lines = @()
            foreach ($view in @('Registry64','Registry32')) {
                $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
                $key = $base.OpenSubKey("Software\Classes\CLSID\$clsid\InprocServer32")
                $lines += "--- $view ---"
                if ($key) {
                    foreach ($n in $key.GetValueNames()) {
                        $val = $key.GetValue($n)
                        $nm = if ($n) { $n } else { '(default)' }
                        $lines += "  $nm = $val"
                        if ($n -eq 'CodeBase') {
                            $p = ($val -replace '^file:///','') -replace '/','\'
                            $lines += "  CodeBase file exists: $(Test-Path $p)"
                        }
                    }
                    $key.Dispose()
                } else { $lines += '  (no InprocServer32 key)' }
                $base.Dispose()
            }
            $lines += ''
            $lines += '--- assemblies present in install dir ---'
            $lines += (Get-ChildItem -Path $installDir -Filter '*.dll' -ErrorAction SilentlyContinue |
                        ForEach-Object { '  ' + $_.Name })
            $lines += ''
            $lines += "--- .NET Framework release ---"
            $ndp = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
            $ndpRelease = 'not present'
            if (Test-Path $ndp) {
                $ndpRelease = (Get-ItemProperty $ndp -Name Release -ErrorAction SilentlyContinue).Release
            }
            $lines += "  Release = $ndpRelease"
            $lines | Set-Content -LiteralPath $dump -Encoding UTF8
            Write-Host "  (diagnostic) registration dump written to $dump"

            # 0x80070002 from COM activation says only "a file was not found".
            # It does not say WHICH file, and that distinction decides whether
            # this is a real packaging bug or an artifact of Windows Sandbox
            # having no Office installed. Fusion's assembly binding log names
            # the assembly that failed to load, which settles it.
            #
            # Fusion only logs for processes started after these values are set,
            # so the probe is re-run below. The sandbox account is an admin, so
            # writing to HKLM here is fine; nothing outside the VM is touched.
            $fusionDir = Join-Path $outDir 'fusion'
            New-Item -ItemType Directory -Path $fusionDir -Force | Out-Null
            $fusionKey = 'HKLM:\SOFTWARE\Microsoft\Fusion'
            try {
                New-Item -Path $fusionKey -Force | Out-Null
                Set-ItemProperty -Path $fusionKey -Name 'EnableLog'    -Value 1 -Type DWord
                Set-ItemProperty -Path $fusionKey -Name 'ForceLog'     -Value 1 -Type DWord
                Set-ItemProperty -Path $fusionKey -Name 'LogFailures'  -Value 1 -Type DWord
                Set-ItemProperty -Path $fusionKey -Name 'LogResourceBinds' -Value 1 -Type DWord
                # Trailing separator is required; Fusion ignores the path without it.
                Set-ItemProperty -Path $fusionKey -Name 'LogPath' -Value ($fusionDir + '\') -Type String

                Start-Process -FilePath $repair -ArgumentList '--verify' -Wait -NoNewWindow `
                    -RedirectStandardOutput (Join-Path $outDir 'repair-verify-fusion.log') `
                    -RedirectStandardError  (Join-Path $outDir 'repair-verify-fusion.log.err') | Out-Null

                $logCount = @(Get-ChildItem -Path $fusionDir -Recurse -File -ErrorAction SilentlyContinue).Count
                Write-Host "  (diagnostic) fusion binding logs captured: $logCount file(s) under $fusionDir"
            } catch {
                Write-Host "  (diagnostic) fusion logging unavailable: $($_.Exception.Message)"
            } finally {
                # Leave the VM's registry as found; it is disposable anyway, but
                # ForceLog on a shared image would be a nasty surprise.
                foreach ($n in 'EnableLog','ForceLog','LogFailures','LogResourceBinds','LogPath') {
                    Remove-ItemProperty -Path $fusionKey -Name $n -ErrorAction SilentlyContinue
                }
            }

            # Opt-in: downloads Process Monitor into the VM. Off by default so a
            # normal acceptance run stays offline and fast.
            if ($WithTrace) { Invoke-ActivationTrace -Repair $repair -OutDir $outDir }
        }
    } else {
        Add-Check 'DeepExcel.Repair.exe --verify reports healthy' $false 'tool missing'
    }

    # 5b. WPS. Its add-in is plain JS -- no COM, no .NET, no bitness -- so it
    #     has no reason to depend on anything Excel above. But the installer
    #     activates it LAST inside the same try block, so an Excel-side failure
    #     skips it entirely: the files land in jsaddons (the [Files] section is
    #     unconditional) while publish.xml never learns about them, and WPS
    #     silently loads nothing. Checking the folder alone would miss this,
    #     which is the whole point of asserting on the manifest.
    $jsaddons = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
    $addinDir = Get-ChildItem -Path $jsaddons -Directory -Filter 'DeepExcel_*' -ErrorAction SilentlyContinue |
                Select-Object -First 1
    Add-Check 'WPS add-in payload copied to jsaddons' ($null -ne $addinDir) "$($addinDir.Name)"

    $publishXml = Join-Path $jsaddons 'publish.xml'
    $registered = $false
    $detail = 'publish.xml not found'
    if (Test-Path $publishXml) {
        try {
            [xml]$manifest = Get-Content -LiteralPath $publishXml -Raw
            $entry = $manifest.SelectNodes('//*[@url]') |
                     Where-Object { $_.url -like 'DeepExcel_*' } |
                     Select-Object -First 1
            $registered = $null -ne $entry
            $detail = if ($registered) { "url=$($entry.url)" } else { 'no DeepExcel entry in publish.xml' }
        } catch {
            $detail = "publish.xml unreadable: $($_.Exception.Message)"
        }
    }
    Add-Check 'WPS add-in registered in publish.xml' $registered $detail

    # 5c. The manifest check above proves the entry was written. It cannot prove
    #     the add-in actually works, so the window is handed to a human here,
    #     before the uninstall step tears everything down.
    #
    #     What is NOT in question any more: enable="enable_dev". That was the
    #     prime suspect for months and it is now ruled out twice over --
    #     Kingsoft's own wpsjs toolchain emits exactly that value from the code
    #     path that builds end-user distribution packages (build.js,
    #     CreatePublishXml), and an earlier sandbox run saw the ribbon tab
    #     appear, which could not happen if WPS refused to load the entry.
    #
    #     The open question is now one step further in: the tab renders from
    #     ribbon.xml with no JS at all, so a tab alone proves nothing about
    #     whether the callbacks are wired. jsplugins.xml -- the manifest that
    #     binds them -- was missing from the release payload until 85ea244, and
    #     that is what produced the field report of "WPS installs but does not
    #     work". So what needs looking at is whether the buttons DO something.
    # 5b-2. Decide whether main.js runs at all.
    #
    # The ribbon tab and its buttons come from ribbon.xml, which WPS renders
    # without any JS. Both JS callbacks (GetImage for the icons, OnAction for
    # the button) do nothing, which is equally consistent with "main.js never
    # loaded" and "main.js loaded but its require() calls threw" -- those are
    # wrapped in try blocks that swallow the error.
    #
    # main.js is written against Node-style require(), so require('fs') should
    # exist in that runtime. A marker written at the top of the file settles it
    # without needing anyone to click anything.
    $jsDiag = Join-Path $outDir 'wps-js-diag.log'
    if ($addinDir) {
        $mainJs = Join-Path $addinDir.FullName 'main.js'
        if (Test-Path $mainJs) {
            $probe = @"
/* DIAG */ try { require('fs').appendFileSync('$($jsDiag -replace '\\','\\')', 'main.js top-level executed ' + new Date().toISOString() + '\n') } catch (e) { }
"@
            $tail = @"

/* DIAG */ try { require('fs').appendFileSync('$($jsDiag -replace '\\','\\')', 'main.js reached end of file\n') } catch (e) { }
"@
            $body = Get-Content -LiteralPath $mainJs -Raw
            Set-Content -LiteralPath $mainJs -Value ($probe + "`n" + $body + $tail) -Encoding UTF8
            Write-Host '  (wps) instrumented main.js with load probes'
        } else {
            Write-Host "  (wps) main.js not found under $($addinDir.FullName)"
        }
    }

    # 5b-3. COM route experiment.
    if ($ComWps -and $addinDir) {
        Write-Host ''
        Write-Host '=== COM route experiment ==='

        # The installer rolled the COM registration back when the Excel-side
        # post-install step failed, so put it back first -- --repair rebuilds
        # both registry views and the Excel add-in key.
        if (Test-Path $repair) {
            Start-Process -FilePath $repair -ArgumentList '--repair' -Wait -NoNewWindow `
                -RedirectStandardOutput (Join-Path $outDir 'com-wps-repair.log') `
                -RedirectStandardError  (Join-Path $outDir 'com-wps-repair.log.err') | Out-Null
            Write-Host '  restored COM registration via --repair'
        }

        # Disable the JS add-in, otherwise its tab would be indistinguishable
        # from a tab loaded through COM.
        $pub = Join-Path $jsaddons 'publish.xml'
        if (Test-Path $pub) {
            Copy-Item $pub (Join-Path $outDir 'publish.xml.before-com-test') -Force
            [xml]$m = Get-Content -LiteralPath $pub -Raw
            $gone = 0
            foreach ($n in @($m.SelectNodes('//*[@url]'))) {
                if ($n.url -like 'DeepExcel_*') { $n.ParentNode.RemoveChild($n) | Out-Null; $gone++ }
            }
            $m.Save($pub)
            Write-Host "  disabled JS add-in (removed $gone publish.xml entry/entries)"
        }

        # WPS reads its own add-in list, plus a trust whitelist on the personal
        # edition. Register under both spellings of the vendor key -- the casing
        # differs between WPS builds.
        foreach ($vendor in 'Kingsoft', 'kingsoft') {
            $base = "HKCU:\Software\$vendor\Office\ET"
            try {
                New-Item -Path "$base\AddIns\DeepExcel.AddIn" -Force | Out-Null
                Set-ItemProperty -Path "$base\AddIns\DeepExcel.AddIn" -Name 'Description'  -Value 'DeepExcel AI AddIn'
                Set-ItemProperty -Path "$base\AddIns\DeepExcel.AddIn" -Name 'FriendlyName' -Value 'DeepExcel AI AddIn'
                Set-ItemProperty -Path "$base\AddIns\DeepExcel.AddIn" -Name 'LoadBehavior' -Value 3 -Type DWord
                New-Item -Path "$base\AddinsWL\DeepExcel.AddIn" -Force | Out-Null
                Set-ItemProperty -Path "$base\AddinsWL\DeepExcel.AddIn" -Name '(default)' -Value 1
            } catch {
                Write-Host "  registering under $vendor failed: $($_.Exception.Message)"
            }
        }
        Write-Host '  registered DeepExcel.AddIn under WPS ET AddIns + AddinsWL'
        Write-Host '  -> any DeepExcel tab in WPS now can only come from COM'
    }

    if ($wpsReady) {
        $etPath = Get-Content -LiteralPath 'C:\wps-et-path.txt' -ErrorAction SilentlyContinue
        if ($etPath) {
            # Keep this file ASCII-only. It is UTF-8 without a BOM, and
            # Windows PowerShell 5.1 decodes such .ps1 files with the system
            # ANSI codepage -- on a CJK codepage the multi-byte sequences eat
            # the closing quote and the whole script stops parsing.
            Write-Host ''
            Write-Host '--------------------------------------------------------------'
            Write-Host ' Launching WPS Spreadsheets. Two things to check,'
            Write-Host ' in this order:'
            Write-Host ''
            Write-Host '  1. Is there a "DeepExcel" tab, WITH its icons?'
            Write-Host '     Icons come from a JS callback; the tab itself does not.'
            Write-Host '     tab + icons     -> jsplugins.xml is being honoured'
            Write-Host '     tab, no icons   -> the manifest is not loading'
            Write-Host '     no tab at all   -> publish.xml never took effect'
            Write-Host ''
            Write-Host '  2. Click "Open panel". Does the task pane appear?'
            Write-Host '     This is the check that matters: it is the exact'
            Write-Host '     symptom a user reported (tab present, clicks dead)'
            Write-Host '     and the one 85ea244 was meant to fix.'
            Write-Host ''
            Write-Host ' enable="enable_dev" is NOT under suspicion -- see 5c.'
            Write-Host ' Also read wps-js-diag.log in the results folder: it says'
            Write-Host ' whether main.js executed at all, which no amount of'
            Write-Host ' looking at the ribbon can tell you.'
            Write-Host ''
            Write-Host ' Close WPS, then press Enter here to continue (uninstall check).'
            Write-Host '--------------------------------------------------------------'
            Start-Process -FilePath $etPath | Out-Null

            # Do NOT use Read-Host to pause here. Under the sandbox LogonCommand
            # it does not reliably block: the previous run sailed straight past
            # it and went on to uninstall, destroying the very state that was
            # supposed to be inspected.
            #
            # Wait on a sentinel file instead. The results folder is mapped
            # read-write, so the host side can drop the file once the human has
            # actually looked at the ribbon.
            # WPS opens on a login dialog plus a "what's new" tab; the ribbon
            # does not exist until both are out of the way. Escape closes the
            # dialog, Ctrl+N gets an actual spreadsheet.
            Start-Sleep -Seconds 25
            Save-GuestScreenshot -OutDir $outDir -Name '01-wps-opened'

            # Synthetic keystrokes are OFF by default. Driving Escape/Ctrl+N
            # produced a WPS window with no DeepExcel tab at all, unlike every
            # manual run -- the automation appears to race the add-in's own
            # initialisation, and a run that perturbs what it measures is worse
            # than no run. WPS also wants a QR-code sign-in that cannot be
            # automated. So by default this pauses and lets a human drive.
            if ($AutoKeys) {
                Send-GuestKeys -Keys '{ESC}' -SleepMs 2500
                Send-GuestKeys -Keys '^n'   -SleepMs 6000
                Save-GuestScreenshot -OutDir $outDir -Name '02-workbook-open'
            }

            # UI Automation also has to run in a child process -- the parent
            # cannot load assemblies any more (see Save-GuestScreenshot).
            $tabName   = [string]([char]0x6253 + [char]0x5F00 + [char]0x9762 + [char]0x677F)
            $treeFile  = Join-Path $outDir 'ui-tree.txt'
            $actionLog = Join-Path $outDir 'ui-actions.log'
            $uiScript = @"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
`$root = [System.Windows.Automation.AutomationElement]::RootElement
`$all = `$root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                       [System.Windows.Automation.Condition]::TrueCondition)
`$names = @("elements: " + `$all.Count)
foreach (`$e in `$all) {
  try { if (`$e.Current.Name) { `$names += (`$e.Current.Name + '  [' + `$e.Current.ControlType.ProgrammaticName + ']') } } catch { }
}
`$names | Set-Content -LiteralPath '$treeFile' -Encoding UTF8

function Invoke-ByName([string]`$n) {
  `$c = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, `$n)
  `$el = `$root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, `$c)
  if (-not `$el) { return "not found: `$n" }
  try {
    `$p = `$el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    `$p.Invoke(); return "invoked: `$n"
  } catch {
    try {
      `$sp = `$el.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
      `$sp.Select(); return "selected: `$n"
    } catch { return ("pattern failed on `$n : " + `$_.Exception.Message) }
  }
}
Invoke-ByName 'DeepExcel' | Out-File -LiteralPath '$actionLog' -Encoding UTF8 -Append
Start-Sleep -Seconds 3
Invoke-ByName '$tabName' | Out-File -LiteralPath '$actionLog' -Encoding UTF8 -Append
"@
            $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($uiScript))
            Start-Process -FilePath 'powershell.exe' `
                -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $enc `
                -Wait -NoNewWindow | Out-Null

            Start-Sleep -Seconds 12
            Save-GuestScreenshot -OutDir $outDir -Name '03-after-panel-click'

            # The add-in's own log is the only place that says how far OnAction
            # got. It lives in the profile, not the mapped folder.
            $panelLog = Join-Path $env:LOCALAPPDATA 'DeepExcel\logs\wps-panel.log'
            if (Test-Path $panelLog) {
                Copy-Item $panelLog (Join-Path $outDir 'wps-panel.log') -Force
                Write-Host '  (wps) copied add-in panel log'
            } else {
                Add-Content -LiteralPath $actionLog -Value "add-in log not created at $panelLog"
            }

            $go = Join-Path (Split-Path -Parent $ResultFile) 'continue.txt'
            Remove-Item $go -ErrorAction SilentlyContinue
            Write-Host " Waiting for $go (create it from the host to continue)."
            # Watch for the add-in's log DURING the wait, not after it.
            #
            # The previous version copied it before the human had touched
            # anything, so it always captured nothing. The log only appears once
            # the panel button is actually clicked, which is precisely what this
            # wait is for. Screenshots are also taken periodically, so an error
            # dialog that the human closes is still recorded.
            $panelLogSrc = Join-Path $env:LOCALAPPDATA 'DeepExcel\logs\wps-panel.log'
            $panelLogDst = Join-Path $outDir 'wps-panel.log'
            $waitUntil = (Get-Date).AddMinutes(45)
            $tick = 0
            while (-not (Test-Path $go) -and (Get-Date) -lt $waitUntil) {
                Start-Sleep -Seconds 5
                $tick++
                if (Test-Path $panelLogSrc) {
                    try { Copy-Item $panelLogSrc $panelLogDst -Force } catch { }
                }
                # every ~30s
                if ($tick % 6 -eq 0) {
                    Save-GuestScreenshot -OutDir $outDir -Name 'live-latest'
                }
            }
            if (Test-Path $panelLogSrc) {
                try { Copy-Item $panelLogSrc $panelLogDst -Force } catch { }
            }
            Save-GuestScreenshot -OutDir $outDir -Name '04-final'
            Write-Host ' Continuing.'

            # Whatever WPS itself recorded about the add-in. Paths are guesses,
            # so this lists what exists rather than assuming any single one.
            $wpsLogSummary = Join-Path $outDir 'wps-logs.txt'
            $roots = @("$env:APPDATA\kingsoft", "$env:LOCALAPPDATA\kingsoft",
                       "$env:APPDATA\Kingsoft", "$env:LOCALAPPDATA\Kingsoft", $env:TEMP)
            $report = @()
            foreach ($root in $roots) {
                if (-not (Test-Path $root)) { continue }
                $hits = Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Extension -in '.log','.txt' -and $_.Length -gt 0 -and
                                       $_.LastWriteTime -gt (Get-Date).AddHours(-2) } |
                        Select-Object -First 40
                foreach ($h in $hits) { $report += ('{0}  ({1} bytes)' -f $h.FullName, $h.Length) }
            }
            $report | Set-Content -LiteralPath $wpsLogSummary -Encoding UTF8
            Write-Host "  (wps) $($report.Count) recent log files listed in $wpsLogSummary"

            # Listing filenames is not enough -- WPS records add-in loading in
            # its own log, and that text is the only place likely to say why the
            # JS callbacks are dead. Pull the relevant lines out; the full log is
            # ~300 KB and most of it is unrelated.
            $interesting = 'DeepExcel|jsaddon|jsplugin|addon|plugin'
            $extract = @()
            foreach ($root in @("$env:APPDATA\kingsoft\office6\log", "$env:APPDATA\Kingsoft\office6\log")) {
                if (-not (Test-Path $root)) { continue }
                $logs = Get-ChildItem -Path $root -Recurse -File -Filter '*.log' -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -gt (Get-Date).AddHours(-2) }
                foreach ($l in $logs) {
                    $hits = Select-String -LiteralPath $l.FullName -Pattern $interesting -ErrorAction SilentlyContinue
                    if ($hits) {
                        $extract += ''
                        $extract += ('===== {0} ({1} matching lines) =====' -f $l.FullName, $hits.Count)
                        $extract += ($hits | Select-Object -First 200 |
                                     ForEach-Object { '{0}: {1}' -f $_.LineNumber, $_.Line.Trim() })
                    }
                }
            }
            $extract | Set-Content -LiteralPath (Join-Path $outDir 'wps-log-extract.txt') -Encoding UTF8
            Write-Host "  (wps) $($extract.Count) log lines extracted"

            # The installed add-in folder as WPS actually sees it.
            if ($addinDir) {
                Get-ChildItem -Path $addinDir.FullName -Recurse -File -ErrorAction SilentlyContinue |
                    ForEach-Object { $_.FullName.Substring($addinDir.FullName.Length + 1) } |
                    Set-Content -LiteralPath (Join-Path $outDir 'wps-addin-tree.txt') -Encoding UTF8
            }
        }
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
    # read the outcome when running interactively. -NoPause turns that off for
    # unattended runs, where nobody is there to press the key and the window
    # would otherwise outlive the run indefinitely.
    if ((-not $NoPause) -and $env:COMPUTERNAME -and $Host.UI.RawUI) {
        Write-Host ''
        Write-Host 'Press Enter to close.'
        try { Read-Host | Out-Null } catch { }
    }

    if ($failed.Count -gt 0) { exit 1 }
}

# ---------------------------------------------------------------------------
    # Note: -Param:$switch.IsPresent does not parse (the member access is taken
    # as a separate token). These targets are [bool], so cast explicitly.
if ($Launch) {
    Start-SandboxRun -Setup $SetupPath -WithTrace ([bool]$Trace) -WpsSetup $WpsSetupPath -ComWpsTest ([bool]$ComWps) -AutoKeysTest ([bool]$AutoKeys) -NoPauseRun ([bool]$NoPause)
} else {
    if (-not $SetupPath) { throw 'Specify -SetupPath, or use -Launch to run this inside Windows Sandbox.' }
    Invoke-GuestVerification -Setup $SetupPath -ResultFile $ResultPath -WithTrace ([bool]$Trace) -WithWps ([bool]$WithWps) -ComWps ([bool]$ComWps) -AutoKeys ([bool]$AutoKeys) -NoPause ([bool]$NoPause)
}
