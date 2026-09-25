param([string]$OutputDir)

$ErrorActionPreference = 'Stop'
$baseDir = Resolve-Path "$PSScriptRoot\.."
$addinDir = Join-Path $baseDir 'src\DeepExcel.AddIn'
$outDir = if ($OutputDir) { [IO.Path]::GetFullPath($OutputDir) } else { Join-Path $addinDir 'bin\Release' }
$packages = Join-Path $baseDir 'packages'
$csc = Join-Path $packages 'Microsoft.Net.Compilers.3.8.0\tools\csc.exe'

if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

# Fail before csc runs, not after.
#
# csc deletes the existing output file before creating the new one. When Excel
# holds DeepExcel.AddIn.dll the delete succeeds (the CLR loader opens with
# FILE_SHARE_DELETE) but the create fails -- so a "failed" build leaves
# bin\Release with no add-in DLL at all. bin/ is gitignored, so there is nothing
# to restore from and the next thing the developer sees is "加载项主文件缺失".
$lockingProcesses = @(Get-Process -Name 'EXCEL', 'et', 'wps' -ErrorAction SilentlyContinue)
if ($lockingProcesses.Count -gt 0 -and (Test-Path (Join-Path $outDir 'DeepExcel.AddIn.dll'))) {
    $names = ($lockingProcesses | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
    throw @"
Close Excel/WPS before compiling. Running: $names

They hold a lock on $outDir\DeepExcel.AddIn.dll. Compiling anyway would DELETE
the existing DLL and then fail to write the new one, leaving no add-in at all.

  Get-Process EXCEL,et,wps -ErrorAction SilentlyContinue | Stop-Process -Force

To compile without touching the registered build, pass a different output path:

  scripts\_compile_only.ps1 -OutputDir `$env:TEMP\DeepExcelBuild
"@
}

# Keep bin\Release reproducible.  WebView2's managed AnyCPU loader resolves the
# native DLL from runtimes\win-<arch>\native.  A root-level x86 loader shadows
# that mechanism and makes the pane fail in 64-bit Excel.
$webViewPackage = Join-Path $packages 'Microsoft.Web.WebView2.1.0.2420.47'
$dependencyCopies = @{
    (Join-Path $webViewPackage 'lib\net45\Microsoft.Web.WebView2.WinForms.dll') = (Join-Path $outDir 'Microsoft.Web.WebView2.WinForms.dll')
    (Join-Path $webViewPackage 'lib\net45\Microsoft.Web.WebView2.Core.dll') = (Join-Path $outDir 'Microsoft.Web.WebView2.Core.dll')
    (Join-Path $packages 'System.Text.Json.8.0.0\lib\net462\System.Text.Json.dll') = (Join-Path $outDir 'System.Text.Json.dll')
    (Join-Path $packages 'System.Text.Encodings.Web.8.0.0\lib\net462\System.Text.Encodings.Web.dll') = (Join-Path $outDir 'System.Text.Encodings.Web.dll')
    (Join-Path $packages 'Microsoft.Bcl.AsyncInterfaces.8.0.0\lib\net462\Microsoft.Bcl.AsyncInterfaces.dll') = (Join-Path $outDir 'Microsoft.Bcl.AsyncInterfaces.dll')
    (Join-Path $packages 'System.Buffers.4.5.1\lib\net461\System.Buffers.dll') = (Join-Path $outDir 'System.Buffers.dll')
    (Join-Path $packages 'System.Memory.4.5.5\lib\net461\System.Memory.dll') = (Join-Path $outDir 'System.Memory.dll')
    (Join-Path $packages 'System.Numerics.Vectors.4.5.0\lib\net46\System.Numerics.Vectors.dll') = (Join-Path $outDir 'System.Numerics.Vectors.dll')
    (Join-Path $packages 'System.Runtime.CompilerServices.Unsafe.6.0.0\lib\net461\System.Runtime.CompilerServices.Unsafe.dll') = (Join-Path $outDir 'System.Runtime.CompilerServices.Unsafe.dll')
    (Join-Path $packages 'System.Threading.Tasks.Extensions.4.5.4\lib\portable-net45+win8+wp8+wpa81\System.Threading.Tasks.Extensions.dll') = (Join-Path $outDir 'System.Threading.Tasks.Extensions.dll')
    (Join-Path $packages 'System.ValueTuple.4.5.0\lib\net47\System.ValueTuple.dll') = (Join-Path $outDir 'System.ValueTuple.dll')
    'C:\Program Files (x86)\Common Files\Microsoft Shared\MSEnv\PublicAssemblies\Extensibility.dll' = (Join-Path $outDir 'Extensibility.dll')
    'C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\Microsoft.Office.Interop.Excel.dll' = (Join-Path $outDir 'Microsoft.Office.Interop.Excel.dll')
    'C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\OFFICE.dll' = (Join-Path $outDir 'OFFICE.dll')
    'C:\Windows\assembly\GAC_MSIL\Microsoft.Vbe.Interop\15.0.0.0__71e9bce111e9429c\Microsoft.Vbe.Interop.dll' = (Join-Path $outDir 'Microsoft.Vbe.Interop.dll')
}
foreach ($source in $dependencyCopies.Keys) {
    if (-not (Test-Path $source)) { throw "Required build dependency not found: $source" }
    Copy-Item -LiteralPath $source -Destination $dependencyCopies[$source] -Force
}

foreach ($arch in @('x86', 'x64', 'arm64')) {
    $nativeSource = Join-Path $webViewPackage "runtimes\win-$arch\native\WebView2Loader.dll"
    $nativeDest = Join-Path $outDir "runtimes\win-$arch\native"
    if (-not (Test-Path $nativeSource)) { throw "Missing WebView2 $arch loader: $nativeSource" }
    New-Item -ItemType Directory -Path $nativeDest -Force | Out-Null
    Copy-Item -LiteralPath $nativeSource -Destination (Join-Path $nativeDest 'WebView2Loader.dll') -Force
}
$legacyRootLoader = Join-Path $outDir 'WebView2Loader.dll'
if (Test-Path $legacyRootLoader) { Remove-Item -LiteralPath $legacyRootLoader -Force }

# 收集 .cs 文件，排除：
# - bin/obj 目录
# - DeepExcelRibbon.cs（VSTO 依赖）
# - Interop 目录（自定义 interop 会与引用的 Extensibility.dll/OFFICE.dll 冲突，导致 CS0436 和 QueryInterface 失败）
$csFiles = Get-ChildItem -Path $addinDir -Recurse -Filter '*.cs' -ErrorAction SilentlyContinue |
    Where-Object {
        $_.FullName -notmatch '\\(obj|bin)\\' -and
        $_.Name -ne 'DeepExcelRibbon.cs' -and
        $_.FullName -notmatch '\\Interop\\'
    } |
    ForEach-Object { $_.FullName }

Write-Host "Compiling $($csFiles.Count) files (excluded Interop/, DeepExcelRibbon.cs)..."

# 构建参数
$args = @(
    '/target:library',
    "/out:`"$outDir\DeepExcel.AddIn.dll`"",
    '/langversion:9.0',
    '/optimize',
    '/debug:pdbonly',
    '/define:TRACE',
    '/platform:anycpu',
    '/unsafe-',
    '/nowin32manifest',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\netstandard.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Runtime.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Core.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Drawing.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Net.Http.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Xml.dll"',
    # JavaScriptSerializer, used only by Updates/ to parse the signed update
    # manifest. In the GAC, so the same source also compiles into the standalone
    # DeepExcel.Updater.exe with no files beside it.
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Web.Extensions.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Xml.Linq.dll"',
    '/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Microsoft.CSharp.dll"',
    '/reference:"C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\Microsoft.Office.Interop.Excel.dll"',
    '/reference:"C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\OFFICE.dll"',
    '/reference:"C:\Program Files (x86)\Common Files\Microsoft Shared\MSEnv\PublicAssemblies\Extensibility.dll"',
    '/reference:"C:\Windows\assembly\GAC_MSIL\Microsoft.Vbe.Interop\15.0.0.0__71e9bce111e9429c\Microsoft.Vbe.Interop.dll"',
    "/reference:`"$outDir\Microsoft.Web.WebView2.WinForms.dll`"",
    "/reference:`"$outDir\Microsoft.Web.WebView2.Core.dll`"",
    "/reference:`"$outDir\System.Text.Json.dll`"",
    "/reference:`"$outDir\System.Text.Encodings.Web.dll`"",
    "/reference:`"$outDir\Microsoft.Bcl.AsyncInterfaces.dll`"",
    "/reference:`"$outDir\System.Buffers.dll`"",
    "/reference:`"$outDir\System.Memory.dll`"",
    "/reference:`"$outDir\System.Numerics.Vectors.dll`"",
    "/reference:`"$outDir\System.Runtime.CompilerServices.Unsafe.dll`"",
    "/reference:`"$outDir\System.Threading.Tasks.Extensions.dll`"",
    "/reference:`"$outDir\System.ValueTuple.dll`""
)

foreach ($f in $csFiles) { $args += "`"$f`"" }

# /resource 参数需要特殊处理
$args += '/resource:"' + $addinDir + '\Resources\DeepExcelRibbon.xml",DeepExcel.AddIn.Resources.DeepExcelRibbon.xml'

# 调用 csc
& $csc $args 2>&1 | ForEach-Object {
    $line = $_.ToString()
    Write-Host $line
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "BUILD FAILED (exit $LASTEXITCODE)"
    exit 1
}

Write-Host ""
Write-Host "=== BUILD SUCCESS ==="
Get-Item "$outDir\DeepExcel.AddIn.dll" | Select-Object Name, Length, LastWriteTime

# 复制 sidecar 文件
$srcSidecar = Join-Path $baseDir 'src\DeepExcel.Sidecar'
$sidecarOutput = [IO.Path]::GetFullPath((Join-Path $outDir 'sidecar'))
$expectedOutputParent = [IO.Path]::GetFullPath($outDir).TrimEnd('\')
if ((Split-Path -Parent $sidecarOutput).TrimEnd('\') -ne $expectedOutputParent) {
    throw "Refusing to recreate unexpected sidecar output: $sidecarOutput"
}
if (Test-Path -LiteralPath $sidecarOutput) {
    if ((Get-Item -LiteralPath $sidecarOutput -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
        throw "Refusing to recreate reparse-point sidecar output: $sidecarOutput"
    }
    Remove-Item -LiteralPath $sidecarOutput -Recurse -Force
}
New-Item -ItemType Directory -Path $sidecarOutput -Force | Out-Null
# 整个侧车源码树都复制（顶层 .py + 子包 / 数据目录），只排除测试与缓存。
# 以前逐个列出 4 个文件名：新增模块就会在用户机器上 ImportError，而开发机
# 直接跑源码目录，永远发现不了。
Get-ChildItem -LiteralPath $srcSidecar -Force | Where-Object {
    $_.Name -notin @('tests', '__pycache__', '.pytest_cache') -and
    ($_.PSIsContainer -or $_.Extension -in @('.py', '.md', '.json'))
} | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $sidecarOutput -Recurse -Force
}
Get-ChildItem -LiteralPath $sidecarOutput -Recurse -Directory -Force |
    Where-Object { $_.Name -in @('__pycache__', 'tests', '.pytest_cache') } |
    Remove-Item -Recurse -Force
Write-Host "Sidecar files copied"

# ★ 复制 WebViewAssets 前端构建产物到 bin\Release
# csc.exe 直接编译不执行 msbuild AfterBuild target，必须手动复制
$srcAssets = Join-Path $addinDir 'WebViewAssets'
$dstAssets = [IO.Path]::GetFullPath((Join-Path $outDir 'WebViewAssets'))
if (Test-Path $srcAssets) {
    if ((Split-Path -Parent $dstAssets).TrimEnd('\') -ne $expectedOutputParent) {
        throw "Refusing to recreate unexpected WebViewAssets output: $dstAssets"
    }
    if (Test-Path -LiteralPath $dstAssets) {
        if ((Get-Item -LiteralPath $dstAssets -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Refusing to recreate reparse-point WebViewAssets output: $dstAssets"
        }
        Remove-Item -LiteralPath $dstAssets -Recurse -Force
    }
    New-Item -ItemType Directory -Path $dstAssets -Force | Out-Null
    Copy-Item (Join-Path $srcAssets '*') $dstAssets -Recurse -Force
    Write-Host "WebViewAssets copied to bin\Release"
} else {
    throw "WebViewAssets source not found at $srcAssets"
}

Copy-Item (Join-Path $addinDir 'App.config') (Join-Path $outDir 'DeepExcel.AddIn.dll.config') -Force
