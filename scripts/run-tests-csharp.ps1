# Build and run the C# unit tests without dotnet CLI or Visual Studio.
#
# This project has no dotnet CLI available, so the test project cannot be built
# with `dotnet test`. Roslyn (packages/Microsoft.Net.Compilers) and
# xunit.runner.console are used directly instead.
#
# The non-obvious part is the facade references: xunit.assert ships as
# netstandard1.1 and type-forwards through System.Runtime, System.Collections
# and friends. Without them the compile fails with a wall of CS0012 errors that
# look like missing usings. The facades come from the
# Microsoft.NETFramework.ReferenceAssemblies.net48 package rather than the
# machine, so this works on a box with no Visual Studio installed.
#
#   powershell -ExecutionPolicy Bypass -File scripts\run-tests-csharp.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\run-tests-csharp.ps1 -Filter SessionManagerTests

param(
    [string]$Filter,
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
$baseDir = Resolve-Path "$PSScriptRoot\.."
$packages = Join-Path $baseDir 'packages'
$testDir = Join-Path $baseDir 'src\DeepExcel.Tests'
$outDir = if ($OutputDir) { [IO.Path]::GetFullPath($OutputDir) } else { Join-Path $env:TEMP 'DeepExcelTestRun' }

$csc = Join-Path $packages 'Microsoft.Net.Compilers.3.8.0\tools\csc.exe'
$runner = Join-Path $packages 'xunit.runner.console.2.6.6\tools\net472\xunit.console.exe'
$facadeDir = Join-Path $packages 'Microsoft.NETFramework.ReferenceAssemblies.net48.1.0.3\build\.NETFramework\v4.8\Facades'

foreach ($required in @($csc, $runner)) {
    if (-not (Test-Path $required)) { throw "Missing build dependency: $required" }
}
if (-not (Test-Path $facadeDir)) {
    throw @"
Reference assembly facades are missing. Restore them with:
  .\nuget.exe install Microsoft.NETFramework.ReferenceAssemblies.net48 -Version 1.0.3 -OutputDirectory packages
"@
}

if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# Build the add-in into the test output folder. A separate directory keeps this
# away from the registered bin\Release build, so running tests never disturbs a
# running Excel (see the guard in _compile_only.ps1).
Write-Host '==> Building DeepExcel.AddIn'
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot '_compile_only.ps1') -OutputDir $outDir |
    Where-Object { $_ -match 'error|BUILD' }
if (-not (Test-Path (Join-Path $outDir 'DeepExcel.AddIn.dll'))) { throw 'Add-in build failed' }

Write-Host '==> Compiling tests'
$xunitLibs = @{
    'xunit.assert.dll'             = Join-Path $packages 'xunit.assert.2.6.6\lib\netstandard1.1\xunit.assert.dll'
    'xunit.core.dll'               = Join-Path $packages 'xunit.extensibility.core.2.6.6\lib\net452\xunit.core.dll'
    'xunit.execution.desktop.dll'  = Join-Path $packages 'xunit.extensibility.execution.2.6.6\lib\net452\xunit.execution.desktop.dll'
    'xunit.abstractions.dll'       = Join-Path $packages 'xunit.abstractions.2.0.3\lib\net35\xunit.abstractions.dll'
}
foreach ($name in $xunitLibs.Keys) {
    if (-not (Test-Path $xunitLibs[$name])) { throw "Missing xunit assembly: $($xunitLibs[$name])" }
    Copy-Item $xunitLibs[$name] (Join-Path $outDir $name) -Force
}

$fx = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319'
$references = @(
    "/reference:`"$fx\mscorlib.dll`"",
    "/reference:`"$fx\System.dll`"",
    "/reference:`"$fx\System.Core.dll`"",
    "/reference:`"$fx\System.Net.Http.dll`"",
    "/reference:`"$fx\netstandard.dll`"",
    "/reference:`"$outDir\DeepExcel.AddIn.dll`"",
    "/reference:`"$outDir\System.Text.Json.dll`"",
    "/reference:`"$outDir\System.Text.Encodings.Web.dll`"",
    "/reference:`"$outDir\System.Memory.dll`"",
    "/reference:`"$outDir\System.Buffers.dll`"",
    "/reference:`"$outDir\System.Runtime.CompilerServices.Unsafe.dll`"",
    "/reference:`"$outDir\System.Threading.Tasks.Extensions.dll`"",
    "/reference:`"$outDir\System.ValueTuple.dll`"",
    "/reference:`"$outDir\xunit.assert.dll`"",
    "/reference:`"$outDir\xunit.core.dll`"",
    "/reference:`"$outDir\xunit.abstractions.dll`"",
    # Several tests use Excel interop types in their signatures. _compile_only.ps1
    # already staged these next to the add-in, so referencing them is cheaper and
    # less brittle than maintaining an exclusion list that grows with every new
    # test file.
    "/reference:`"$outDir\Microsoft.Office.Interop.Excel.dll`"",
    "/reference:`"$outDir\OFFICE.dll`"",
    "/reference:`"$outDir\Microsoft.Vbe.Interop.dll`"",
    "/reference:`"$outDir\Extensibility.dll`""
)
Get-ChildItem (Join-Path $facadeDir '*.dll') | ForEach-Object {
    $references += "/reference:`"$($_.FullName)`""
}

$sources = Get-ChildItem -Path $testDir -Filter '*.cs' -File |
    ForEach-Object { "`"$($_.FullName)`"" }
if (-not $sources) { throw "No test sources found in $testDir" }
Write-Host "    $($sources.Count) test file(s)"

$compileArgs = @(
    '/target:library',
    "/out:`"$outDir\DeepExcel.Tests.dll`"",
    '/langversion:9.0',
    '/platform:anycpu',
    '/debug:pdbonly',
    '/nowarn:CS1701'
) + $references + $sources

& $csc $compileArgs 2>&1 | Where-Object { $_ -match 'error|warning CS' } | Select-Object -First 15
if ($LASTEXITCODE -ne 0) { throw "Test compilation failed (exit $LASTEXITCODE)" }

# Test data lives next to the test assembly, not in the repository tree, so the
# tests do not have to guess where the repository root is at runtime.
$fixtureSource = Join-Path $testDir 'fixtures'
if (Test-Path $fixtureSource) {
    Copy-Item $fixtureSource (Join-Path $outDir 'fixtures') -Recurse -Force
}

# Mirror the deployed layout: the add-in's binding redirects sit next to it.
$appConfig = Join-Path $baseDir 'src\DeepExcel.AddIn\App.config'
if (-not (Test-Path $appConfig)) { throw "Missing $appConfig" }
Copy-Item $appConfig (Join-Path $outDir 'DeepExcel.AddIn.dll.config') -Force
Copy-Item $appConfig (Join-Path $outDir 'DeepExcel.Tests.dll.config') -Force

Write-Host '==> Running tests'
# -appdomains denied is required, not a preference.
#
# System.Text.Json 8.0 references System.Runtime.CompilerServices.Unsafe 4.0.4.1
# while the shipped file is 6.0.0.0. In a separate AppDomain that reference fails
# to bind and JsonSerializer's static constructor throws, so any test touching
# JSON fails. Running in the runner's own process loads the assembly through the
# LoadFrom context, which resolves same-directory dependencies with relaxed
# versioning -- the same way mscoree loads the add-in from its CodeBase inside
# EXCEL.EXE, so this is also the closer match to production.
#
# Without this flag the result depended on assembly load order: the full suite
# passed while the same tests failed in isolation. A green run that means
# nothing is worse than a red one.
$runnerArgs = @("$outDir\DeepExcel.Tests.dll", '-noshadow', '-nologo', '-appdomains', 'denied')
if ($Filter) { $runnerArgs += @('-class', "DeepExcel.Tests.$Filter") }

& $runner $runnerArgs
$exit = $LASTEXITCODE
if ($exit -ne 0) {
    Write-Host "TESTS FAILED (exit $exit)"
    exit $exit
}
Write-Host ''
Write-Host '=== TESTS PASSED ==='
