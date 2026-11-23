# 在 Windows Sandbox 里试用 WPS 版 DeepExcel —— 关掉沙箱窗口，WPS 和 DeepExcel 全部消失，本机不留任何东西。
#
# This is the hands-on trial tool, not the release acceptance test
# (that is verify-install-sandbox.ps1). It gets a person to "WPS is open with
# the DeepExcel tab" with as few clicks as possible.
#
# Usage on the host:
#   powershell -ExecutionPolicy Bypass -File scripts\try-wps-sandbox.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\try-wps-sandbox.ps1 -SetupPath dist\DeepExcel.Setup.exe
#
# What it does:
#   host  : fetch the latest public DeepExcel installer (SHA-256 checked) or use
#           -SetupPath; start a sandbox via `wsb`; turn off Smart App Control
#           inside it (the installer is unsigned); start the guest half.
#   guest : start the WPS installer (or open wps.cn if none is cached) and wait
#           for it; install DeepExcel silently; launch WPS Spreadsheets.
#
# Two 2026-10 sandbox facts this works around (see docs/DEPLOYMENT.md):
#   - .wsb LogonCommand no longer runs, so everything goes through `wsb exec`.
#   - The sandbox image ships with Smart App Control on, which blocks unsigned
#     installers with no "run anyway".
#
# Everything persistent lives in dist\wps-sandbox (gitignored):
#   stage\  mapped read-only  -> C:\DeepExcelTry     (installers + this script)
#   out\    mapped writable   -> C:\DeepExcelTryOut  (logs; WPS installer cache)
# The WPS installer downloaded inside the sandbox is copied to out\ and reused
# on the next run, so WPS is only fetched once.

[CmdletBinding()]
param(
    # Guest half. Started by the host via `wsb exec`; not for direct use.
    [switch]$Guest,
    # Install a specific DeepExcel build instead of the latest public release.
    [string]$SetupPath,
    # A WPS installer on the host. Optional: one is cached after the first run.
    [string]$WpsSetupPath,
    # Sandbox RAM cap. Bounded on purpose: the point is not to slow the host.
    [int]$MemoryMB = 4096,
    # Testing this script itself: install DeepExcel without waiting for WPS.
    [switch]$SkipWps
)

$ErrorActionPreference = 'Stop'

$Repo         = '1256048998/DeepExcel'
$GuestStage   = 'C:\DeepExcelTry'
$GuestOut     = 'C:\DeepExcelTryOut'

# ---------------------------------------------------------------------------
# Host
# ---------------------------------------------------------------------------
function Get-LatestReleaseInstaller {
    param([string]$Dest)

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $base = "https://github.com/$Repo/releases/latest/download"

    $sums = (Invoke-WebRequest -Uri "$base/SHA256SUMS.txt" -UseBasicParsing).Content
    if ($sums -is [byte[]]) { $sums = [Text.Encoding]::UTF8.GetString($sums) }
    $line = ($sums -split "`n") | Where-Object { $_ -match '\sDeepExcel\.Setup\.exe\s*$' } | Select-Object -First 1
    if (-not $line) { throw '最新发布的 SHA256SUMS.txt 里没有 DeepExcel.Setup.exe 这一行。' }
    $expected = ($line.Trim() -split '\s+')[0].ToLowerInvariant()

    if ((Test-Path $Dest) -and ((Get-FileHash $Dest -Algorithm SHA256).Hash.ToLowerInvariant() -eq $expected)) {
        Write-Host '  DeepExcel 安装包：已是最新发布版（使用缓存）'
        return
    }

    Write-Host '  DeepExcel 安装包：正在从 GitHub 下载最新发布版...'
    $tmp = "$Dest.download"
    Invoke-WebRequest -Uri "$base/DeepExcel.Setup.exe" -OutFile $tmp -UseBasicParsing
    $actual = (Get-FileHash $tmp -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne $expected) {
        Remove-Item $tmp -Force
        throw "下载的安装包 SHA-256 不匹配（期望 $expected，实际 $actual），已丢弃。"
    }
    Move-Item $tmp $Dest -Force
    Write-Host '  DeepExcel 安装包：下载完成，SHA-256 校验通过'
}

function Find-HostWpsInstaller {
    param([string]$Cached)
    if ($WpsSetupPath) {
        if (-not (Test-Path $WpsSetupPath)) { throw "找不到 WPS 安装包：$WpsSetupPath" }
        return (Resolve-Path $WpsSetupPath).Path
    }
    if (Test-Path $Cached) { return $Cached }
    $dl = Join-Path $env:USERPROFILE 'Downloads'
    if (Test-Path $dl) {
        $hit = Get-ChildItem $dl -File -Filter '*.exe' -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -match '(?i)wps' } |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    return $null
}

function Invoke-Wsb {
    param([string[]]$WsbArgs)
    # Windows PowerShell turns native stderr into a terminating error under
    # 'Stop'; wsb failures must come back as a value (the readiness poll
    # expects several).
    $ErrorActionPreference = 'Continue'
    $output = & wsb.exe @WsbArgs 2>&1 | Out-String
    return [pscustomobject]@{ Ok = ($LASTEXITCODE -eq 0); Output = $output.Trim() }
}

function Start-Trial {
    if (-not (Get-Command 'wsb.exe' -ErrorAction SilentlyContinue)) {
        throw @'
找不到 Windows Sandbox（wsb.exe）。需要 Windows 10/11 专业版或企业版，并在管理员 PowerShell 里启用后重启：
  Enable-WindowsOptionalFeature -FeatureName "Containers-DisposableClientVM" -Online
'@
    }

    $running = Invoke-Wsb @('list')
    if ($running.Output -match '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-') {
        throw '已经有一个 Windows Sandbox 在运行（同一时间只能开一个）。先关掉那个沙箱窗口再运行本脚本。'
    }

    # Not under %LOCALAPPDATA%: when this runs from a packaged (MSIX) app, writes
    # there are virtualized, and the sandbox service then fails the mapping with
    # a bare 0x80070002. dist\ is gitignored and always a real path.
    $root  = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\dist\wps-sandbox'))
    $stage = Join-Path $root 'stage'
    $out   = Join-Path $root 'out'
    New-Item -ItemType Directory -Path $stage, $out -Force | Out-Null

    Write-Host '准备文件...'
    $setupDest = Join-Path $stage 'DeepExcel.Setup.exe'
    if ($SetupPath) {
        if (-not (Test-Path $SetupPath)) { throw "找不到安装包：$SetupPath" }
        Copy-Item $SetupPath $setupDest -Force
        Write-Host "  DeepExcel 安装包：$SetupPath"
    } else {
        Get-LatestReleaseInstaller -Dest $setupDest
    }
    $ver = (Get-Item $setupDest).VersionInfo.ProductVersion
    Write-Host "  DeepExcel 版本：$ver"

    # A WPS installer the previous sandbox downloaded is promoted into stage\.
    $wpsStaged = Join-Path $stage 'WpsSetup.exe'
    $wpsFromLastRun = Join-Path $out 'WpsSetup.exe'
    if (Test-Path $wpsFromLastRun) { Move-Item $wpsFromLastRun $wpsStaged -Force }
    $wps = Find-HostWpsInstaller -Cached $wpsStaged
    if ($wps -and ($wps -ne $wpsStaged)) { Copy-Item $wps $wpsStaged -Force }
    if ($wps) { Write-Host "  WPS 安装包：$wps" }
    else      { Write-Host '  WPS 安装包：本机没有，首次需要在沙箱里从 wps.cn 下载（之后会自动缓存）' }

    Copy-Item $PSCommandPath (Join-Path $stage 'try-wps-sandbox.ps1') -Force
    Get-ChildItem $out -File -Filter '*.log' -ErrorAction SilentlyContinue | Remove-Item -Force

    # Single line, no double quotes: it travels as one native argument.
    $config = "<Configuration><Networking>Enable</Networking><MemoryInMB>$MemoryMB</MemoryInMB>" +
              "<MappedFolders>" +
              "<MappedFolder><HostFolder>$stage</HostFolder><SandboxFolder>$GuestStage</SandboxFolder><ReadOnly>true</ReadOnly></MappedFolder>" +
              "<MappedFolder><HostFolder>$out</HostFolder><SandboxFolder>$GuestOut</SandboxFolder><ReadOnly>false</ReadOnly></MappedFolder>" +
              "</MappedFolders></Configuration>"

    Write-Host "启动沙箱（内存上限 $MemoryMB MB）..."
    $start = Invoke-Wsb @('start', '--config', $config)
    if (-not $start.Ok) { throw "沙箱启动失败：$($start.Output)" }
    $id = [regex]::Match($start.Output, '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}').Value
    if (-not $id) { throw "没能从 wsb 输出里拿到沙箱 ID：$($start.Output)" }

    # `wsb start` is headless: no window and no user logon until something
    # connects, and every ExistingLogin exec fails with 0x80070520 until then.
    Start-Process -FilePath 'wsb.exe' -ArgumentList 'connect', '--id', $id -WindowStyle Hidden

    Write-Host '等待沙箱桌面就绪（通常 30-90 秒）...'
    $deadline = (Get-Date).AddMinutes(5)
    while ($true) {
        if ((Invoke-Wsb @('exec', '--id', $id, '-r', 'ExistingLogin', '-c', 'cmd /c exit 0')).Ok) { break }
        if ((Get-Date) -gt $deadline) { throw "沙箱 5 分钟内没有进入桌面（ID $id）。" }
        Start-Sleep -Seconds 5
    }

    Write-Host '在沙箱里关闭「智能应用控制」（否则未签名的安装包会被直接拦下）...'
    $sac = Invoke-Wsb @('exec', '--id', $id, '-r', 'System', '-c',
        'cmd /c reg add HKLM\SYSTEM\CurrentControlSet\Control\CI\Policy /v VerifiedAndReputablePolicyState /t REG_DWORD /d 0 /f & CiTool.exe --refresh')
    if (-not $sac.Ok) { Write-Warning "关闭智能应用控制失败，安装可能被拦：$($sac.Output)" }

    $guestArgs = "-Guest"
    if ($SkipWps) { $guestArgs += ' -SkipWps' }
    $launch = Invoke-Wsb @('exec', '--id', $id, '-r', 'ExistingLogin', '-c',
        "cmd /c start powershell.exe -NoProfile -ExecutionPolicy Bypass -File $GuestStage\try-wps-sandbox.ps1 $guestArgs")
    if (-not $launch.Ok) { throw "沙箱里的安装脚本没能启动：$($launch.Output)" }

    Write-Host ''
    Write-Host '好了，接下来看沙箱窗口：'
    if ($wps -or $SkipWps) {
        Write-Host '  1. WPS 安装界面弹出后，点「立即安装」'
    } else {
        Write-Host '  1. 沙箱里的浏览器会打开 wps.cn，下载 WPS 并运行安装包，点「立即安装」'
    }
    Write-Host '  2. WPS 装好后，脚本会自动装 DeepExcel 并打开 WPS 表格'
    Write-Host '  3. 新建空白表格 -> 顶部 DeepExcel 选项卡 -> 打开面板'
    Write-Host ''
    Write-Host '试完直接关掉沙箱窗口即可，里面的一切都会被清除。'
    Write-Host "日志：$out"
}

# ---------------------------------------------------------------------------
# Guest
# ---------------------------------------------------------------------------
function Find-Et {
    foreach ($p in @("$env:LOCALAPPDATA\Kingsoft", 'C:\Program Files (x86)\Kingsoft', 'C:\Program Files\Kingsoft')) {
        if (Test-Path $p) {
            $hit = Get-ChildItem -Path $p -Recurse -Filter 'et.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($hit) { return $hit.FullName }
        }
    }
    return $null
}

function Wait-WpsInstalled {
    $staged = Join-Path $GuestStage 'WpsSetup.exe'
    if (Test-Path $staged) {
        Write-Host '启动 WPS 安装程序，请在弹出的窗口里点「立即安装」...'
        Start-Process -FilePath $staged
    } else {
        Write-Host '沙箱里没有 WPS 安装包，正在打开 wps.cn。'
        Write-Host '请下载 WPS（Windows 版）并运行安装包，点「立即安装」。'
        Start-Process 'https://www.wps.cn/'
    }

    Write-Host '等待 WPS 安装完成（最长 30 分钟）...'
    $deadline = (Get-Date).AddMinutes(30)
    $et = $null
    while (-not ($et = Find-Et)) {
        if ((Get-Date) -gt $deadline) { throw '30 分钟内没检测到 WPS 表格（et.exe），放弃。' }
        Start-Sleep -Seconds 5
    }
    Write-Host "检测到 WPS 表格：$et"

    # et.exe appears well before the installer finishes; wait for it to go.
    $quiet = 0
    $deadline = (Get-Date).AddMinutes(10)
    while ($quiet -lt 3 -and (Get-Date) -lt $deadline) {
        $busy = Get-Process -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '(?i)setup|install' -and $_.Name -notmatch '(?i)^(TrustedInstaller|SetupHost)$' }
        if ($busy) { $quiet = 0 } else { $quiet++ }
        Start-Sleep -Seconds 5
    }

    # Cache whatever the user downloaded, so the next sandbox skips the website.
    if (-not (Test-Path $staged)) {
        $dl = Join-Path $env:USERPROFILE 'Downloads'
        $got = Get-ChildItem $dl -File -Filter '*.exe' -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -match '(?i)wps' } |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($got) {
            Copy-Item $got.FullName (Join-Path $GuestOut 'WpsSetup.exe') -Force
            Write-Host "已缓存 WPS 安装包，下次试用不用再下载：$($got.Name)"
        }
    }
    return $et
}

function Start-GuestTrial {
    Start-Transcript -LiteralPath (Join-Path $GuestOut 'guest.log') -Force | Out-Null
    try {
        $et = $null
        if (-not $SkipWps) {
            $et = Find-Et
            if (-not $et) { $et = Wait-WpsInstalled }
            # WPS starts itself after installing; close it so it reloads add-ins.
            Get-Process -Name 'et', 'wps', 'wpp', 'wpspdf' -ErrorAction SilentlyContinue | Stop-Process -Force
            Start-Sleep -Seconds 2
        }

        Write-Host '安装 DeepExcel...'
        $log = Join-Path $GuestOut 'deepexcel-install.log'
        $p = Start-Process -FilePath (Join-Path $GuestStage 'DeepExcel.Setup.exe') `
             -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', "/LOG=$log" -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw "DeepExcel 安装失败，退出码 $($p.ExitCode)。日志：$log" }

        $publish = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons\publish.xml'
        if (-not ((Test-Path $publish) -and (Select-String -LiteralPath $publish -Pattern 'DeepExcel' -Quiet))) {
            throw "DeepExcel 装完了，但 WPS 加载项清单里没有它：$publish"
        }
        Write-Host 'DeepExcel 安装完成，已登记到 WPS 加载项。'

        if ($et) {
            Write-Host '打开 WPS 表格...'
            Start-Process -FilePath $et
            Write-Host ''
            Write-Host '新建空白表格 -> 顶部 DeepExcel 选项卡 -> 打开面板。'
            Write-Host '第一次打开 WPS 可能要你同意协议或扫码登录，按提示操作即可。'
        }
        Write-Host ''
        Write-Host 'RESULT: READY'
    } catch {
        Write-Host ''
        Write-Host "RESULT: FAILED  $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        Stop-Transcript | Out-Null
    }
    Write-Host ''
    Read-Host '这个窗口可以关掉了（按回车关闭）'
}

if ($Guest) { Start-GuestTrial } else { Start-Trial }
