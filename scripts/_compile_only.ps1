$ErrorActionPreference = 'Stop'
$baseDir = Resolve-Path "$PSScriptRoot\.."
$addinDir = Join-Path $baseDir 'src\DeepExcel.AddIn'
$outDir = Join-Path $addinDir 'bin\Release'
$packages = Join-Path $baseDir 'packages'
$csc = Join-Path $packages 'Microsoft.Net.Compilers.3.8.0\tools\csc.exe'

# Collect .cs files, excluding:
# - bin/obj directories
# - DeepExcelRibbon.cs (VSTO dependency)
# - OfficeInterop.cs (conflicts with referenced OFFICE.dll PIA)
# - Extensibility.cs (use GAC PIA instead, see reference below)
# Reason for NOT inlining Extensibility.cs: user machines with Office/VS installed have
# Extensibility.dll PIA in GAC (v7.0.3300.0__b03f5f7f11d50a3a). Inlined version has different
# assembly identity (no version, no PublicKeyToken), so CLR treats them as different types
# even with same GUID -> QueryInterface(IDTExtensibility2) fails -> OnConnection never called.
# Solution: reference the PIA. It's in GAC on all machines with Office Developer Tools or VS.
# For machines without PIA, we ship Extensibility.dll alongside and OnAssemblyResolve loads it.
$csFiles = Get-ChildItem -Path $addinDir -Recurse -Filter '*.cs' -ErrorAction SilentlyContinue |
    Where-Object {
        $_.FullName -notmatch '\\(obj|bin)\\' -and
        $_.Name -ne 'DeepExcelRibbon.cs' -and
        $_.Name -ne 'OfficeInterop.cs' -and
        $_.Name -ne 'Extensibility.cs'
    } |
    ForEach-Object { $_.FullName }

Write-Host "Compiling $($csFiles.Count) files (excluded OfficeInterop.cs, Extensibility.cs, DeepExcelRibbon.cs)..."

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
Copy-Item (Join-Path $srcSidecar 'sidecar.py') (Join-Path $outDir 'sidecar') -Force
Copy-Item (Join-Path $srcSidecar 'ipc.py') (Join-Path $outDir 'sidecar') -Force
Copy-Item (Join-Path $srcSidecar 'excel_tools.py') (Join-Path $outDir 'sidecar') -Force
Copy-Item (Join-Path $srcSidecar 'system_prompt.py') (Join-Path $outDir 'sidecar') -Force
Write-Host "Sidecar files copied"

# ★ 复制 WebViewAssets 前端构建产物到 bin\Release
# csc.exe 直接编译不执行 msbuild AfterBuild target，必须手动复制
$srcAssets = Join-Path $addinDir 'WebViewAssets'
$dstAssets = Join-Path $outDir 'WebViewAssets'
if (Test-Path $srcAssets) {
    if (-not (Test-Path $dstAssets)) { New-Item -ItemType Directory -Path $dstAssets -Force | Out-Null }
    Copy-Item (Join-Path $srcAssets '*') $dstAssets -Recurse -Force
    Write-Host "WebViewAssets copied to bin\Release"
} else {
    Write-Host "WARNING: WebViewAssets source not found at $srcAssets"
}
