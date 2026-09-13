# WPS COM 加载项验证脚本
#
# 目的：验证 WPS 能否直接加载 DeepExcel 的 Excel COM 加载项——也就是
# 方方格子(FFCell)的做法：一份 VSTO/COM 加载项同时服务 Excel 和 WPS，
# 完全不需要单独的 WPS JS 加载项。
#
# 背景：现有的 WPS JS 加载项在干净环境实测是坏的——选项卡出现，但所有
# JS 回调都不触发。根因已确认：加载项目录缺 wpsjs 官方要求的入口文件
# index.html，WPS 因此从未加载过 main.js；而且我们的 WPS 代码按 Node
# 环境写（fs / child_process），WPS 实际是浏览器环境，两者不兼容。
#
# 本脚本只做两件事，都在 HKCU 和一个 XML 里，不安装任何软件：
#   1. 把 DeepExcel.AddIn 注册给 WPS 表格(ET)
#   2. 临时禁用 JS 加载项，这样 Ribbon 上再出现 DeepExcel 就只可能来自 COM
#
# 用法（普通权限即可，不需要管理员）：
#   powershell -ExecutionPolicy Bypass -File scripts\wps-com-test.ps1 -Enable
#   ... 打开 WPS 表格查看 ...
#   powershell -ExecutionPolicy Bypass -File scripts\wps-com-test.ps1 -Rollback

[CmdletBinding()]
param(
    [switch]$Enable,
    [switch]$Rollback
)

$ErrorActionPreference = 'Stop'

$ProgId   = 'DeepExcel.AddIn'
$Jsaddons = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$Publish  = Join-Path $Jsaddons 'publish.xml'
$Backup   = Join-Path $Jsaddons 'publish.xml.comtest-backup'

function Write-Step($m) { Write-Host "  $m" -ForegroundColor Gray }

function Assert-DeepExcelRegistered {
    # COM 注册在不在。不在的话这个验证没有意义——Ribbon 不出现只能说明
    # 加载项没装，不能说明 WPS 不支持 COM。
    $clsid = '{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}'
    $ok = $false
    foreach ($view in @('Registry64','Registry32')) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
        $key  = $base.OpenSubKey("Software\Classes\CLSID\$clsid\InprocServer32")
        if ($key) {
            $cb = $key.GetValue('CodeBase')
            if ($cb) { Write-Step "COM 注册存在($view): $cb"; $ok = $true }
            $key.Dispose()
        }
        $base.Dispose()
    }
    if (-not $ok) {
        Write-Host ''
        Write-Host '中止：没有找到 DeepExcel 的 COM 注册。' -ForegroundColor Red
        Write-Host '这台机器上 DeepExcel 似乎没装，或注册被清掉了。' -ForegroundColor Red
        Write-Host '请先确认 Excel 里能正常看到 DeepExcel 选项卡，再跑本脚本。' -ForegroundColor Red
        exit 1
    }
}

function Enable-ComRoute {
    Write-Host ''
    Write-Host '=== 启用 COM 路线验证 ===' -ForegroundColor Cyan
    Assert-DeepExcelRegistered

    # 1) 禁用 JS 加载项
    if (Test-Path $Publish) {
        if (-not (Test-Path $Backup)) {
            Copy-Item $Publish $Backup -Force
            Write-Step "已备份 publish.xml -> $Backup"
        } else {
            Write-Step '备份已存在，保留原备份（不覆盖）'
        }
        [xml]$m = Get-Content -LiteralPath $Publish -Raw
        $removed = 0
        foreach ($n in @($m.SelectNodes('//*[@url]'))) {
            if ($n.url -like 'DeepExcel_*') { $n.ParentNode.RemoveChild($n) | Out-Null; $removed++ }
        }
        $m.Save($Publish)
        Write-Step "已从 publish.xml 移除 $removed 个 DeepExcel 节点（其他插件不受影响）"
    } else {
        Write-Step 'publish.xml 不存在，说明本机没有 JS 加载项，跳过这步'
    }

    # 2) 注册给 WPS 表格。不同 WPS 版本这个键的大小写不一致，两种都写。
    foreach ($vendor in 'Kingsoft', 'kingsoft') {
        $addins = "HKCU:\Software\$vendor\Office\ET\AddIns\$ProgId"
        $wl     = "HKCU:\Software\$vendor\Office\ET\AddinsWL\$ProgId"
        New-Item -Path $addins -Force | Out-Null
        Set-ItemProperty -Path $addins -Name 'Description'  -Value 'DeepExcel AI AddIn'
        Set-ItemProperty -Path $addins -Name 'FriendlyName' -Value 'DeepExcel AI AddIn'
        Set-ItemProperty -Path $addins -Name 'LoadBehavior' -Value 3 -Type DWord
        New-Item -Path $wl -Force | Out-Null
    }
    Write-Step '已把 DeepExcel.AddIn 注册到 WPS 表格(ET) 的加载项列表和信任白名单'

    Write-Host ''
    Write-Host '下一步：完全退出 WPS（任务栏托盘也要退），再重新打开 WPS 表格。' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '看两件事：' -ForegroundColor Yellow
    Write-Host '  1. Ribbon 里有没有 DeepExcel 选项卡？'
    Write-Host '  2. 如果有，点里面的按钮，面板能不能打开？'
    Write-Host ''
    Write-Host '  有选项卡且面板能开 -> COM 路线可行，可以砍掉整套 WPS JS 代码'
    Write-Host '  有选项卡但面板打不开 -> COM 能加载，但任务窗格不被支持'
    Write-Host '  没有选项卡         -> WPS 不加载 COM 加载项，B 方案不成立'
    Write-Host ''
    Write-Host '看完记得跑回滚：' -ForegroundColor Yellow
    Write-Host '  powershell -ExecutionPolicy Bypass -File scripts\wps-com-test.ps1 -Rollback'
    Write-Host ''
}

function Rollback-ComRoute {
    Write-Host ''
    Write-Host '=== 回滚 ===' -ForegroundColor Cyan

    if (Test-Path $Backup) {
        Copy-Item $Backup $Publish -Force
        Remove-Item $Backup -Force
        Write-Step '已还原 publish.xml，JS 加载项恢复原状'
    } else {
        Write-Step '没有找到备份，publish.xml 未改动'
    }

    foreach ($vendor in 'Kingsoft', 'kingsoft') {
        foreach ($sub in 'AddIns', 'AddinsWL') {
            $p = "HKCU:\Software\$vendor\Office\ET\$sub\$ProgId"
            if (Test-Path $p) { Remove-Item $p -Recurse -Force; Write-Step "已删除 $p" }
        }
    }

    Write-Host ''
    Write-Host '回滚完成。重启 WPS 即可恢复到验证前的状态。' -ForegroundColor Green
    Write-Host ''
}

if ($Enable -and $Rollback) { throw '-Enable 和 -Rollback 不能同时用' }
if ($Enable)        { Enable-ComRoute }
elseif ($Rollback)  { Rollback-ComRoute }
else {
    Write-Host ''
    Write-Host '用法：' -ForegroundColor Cyan
    Write-Host '  -Enable    注册 COM 给 WPS 并临时禁用 JS 加载项'
    Write-Host '  -Rollback  原样还原'
    Write-Host ''
}
