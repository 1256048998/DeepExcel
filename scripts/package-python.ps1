#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$PythonVersion = '3.11.9',
    [string]$OutputDir,
    [string]$TempDir
)

$ErrorActionPreference = 'Stop'
if (-not $OutputDir) { $OutputDir = Join-Path $PSScriptRoot '..\src\DeepExcel.AddIn\bin\Release\python' }
if (-not $TempDir) { $TempDir = Join-Path $env:TEMP 'deepexcel-python-packaging' }
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
$TempDir = [IO.Path]::GetFullPath($TempDir)

if (Test-Path -LiteralPath $TempDir) {
    Remove-Item -LiteralPath $TempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

if (Test-Path -LiteralPath $OutputDir) {
    Remove-Item -LiteralPath $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

$archiveName = "python-$PythonVersion-embed-amd64.zip"
$archivePath = Join-Path $TempDir $archiveName
$expectedPythonHashes = @{
    '3.11.9' = '009D6BF7E3B2DDCA3D784FA09F90FE54336D5B60F0E0F305C37F400BF83CFD3B'
}
if (-not $expectedPythonHashes.ContainsKey($PythonVersion)) {
    throw "No audited Python archive hash is configured for version $PythonVersion."
}
$pythonUrls = @(
    "https://registry.npmmirror.com/-/binary/python/$PythonVersion/$archiveName",
    "https://www.python.org/ftp/python/$PythonVersion/$archiveName"
)

$downloaded = $false
foreach ($url in $pythonUrls) {
    try {
        Write-Host "Downloading embedded Python from $url"
        Invoke-WebRequest -Uri $url -OutFile $archivePath -UseBasicParsing -TimeoutSec 180
        $downloaded = $true
        break
    } catch {
        Write-Warning "Download failed: $($_.Exception.Message)"
    }
}
if (-not $downloaded) { throw 'Unable to download embedded Python.' }

$pythonArchiveHash = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash
if ($pythonArchiveHash -ne $expectedPythonHashes[$PythonVersion]) {
    throw "Embedded Python archive SHA-256 mismatch: $pythonArchiveHash"
}

Expand-Archive -LiteralPath $archivePath -DestinationPath $OutputDir -Force
$pythonExe = Join-Path $OutputDir 'python.exe'
if (-not (Test-Path -LiteralPath $pythonExe)) { throw "python.exe missing from $OutputDir" }
$pythonSignature = Get-AuthenticodeSignature -LiteralPath $pythonExe
if ($pythonSignature.Status -ne 'Valid' -or
    $pythonSignature.SignerCertificate.Subject -notmatch 'Python Software Foundation') {
    throw "Embedded python.exe has an invalid publisher signature: $($pythonSignature.Status)"
}

$pthFile = Get-ChildItem -LiteralPath $OutputDir -Filter 'python*._pth' | Select-Object -First 1
if (-not $pthFile) { throw 'Embedded Python _pth file was not found.' }
$pthContent = Get-Content -LiteralPath $pthFile.FullName
$pthContent = $pthContent | ForEach-Object {
    if ($_ -match '^#import site') { 'import site' } else { $_ }
}
# Embedded Python's _pth mode ignores the script directory. DeepExcel starts
# <install>\sidecar\sidecar.py from <install>\python\python.exe, so add the
# sibling directory explicitly or imports such as excel_tools fail at startup.
if ($pthContent -notcontains '..\sidecar') {
    $pthContent = @($pthContent[0], '..\sidecar') + @($pthContent[1..($pthContent.Count - 1)])
}
Set-Content -LiteralPath $pthFile.FullName -Value $pthContent -Encoding ASCII

$pipVersion = '26.2.1'
$pipWheelName = "pip-$pipVersion-py3-none-any.whl"
$pipWheelPath = Join-Path $TempDir $pipWheelName
$pipWheelUrl = 'https://files.pythonhosted.org/packages/f3/6e/1736e5b4ae2b778ef2f81c47d797de9f891d4d8acb047a24ca37a60294dd/pip-26.2.1-py3-none-any.whl'
$pipWheelHash = '71138ADF1F4CA900CDB7D289C21B7494329F2332B6D85F0E1C42108C0384ED3E'
Invoke-WebRequest -Uri $pipWheelUrl -OutFile $pipWheelPath -UseBasicParsing -TimeoutSec 180
if ((Get-FileHash -LiteralPath $pipWheelPath -Algorithm SHA256).Hash -ne $pipWheelHash) {
    throw 'Pinned pip wheel SHA-256 mismatch.'
}
$pipBootstrap = 'import sys; wheel=sys.argv[1]; sys.argv=sys.argv[1:]; sys.path.insert(0,wheel); from pip._internal.cli.main import main; sys.exit(main())'
& $pythonExe -c $pipBootstrap $pipWheelPath install --no-index --no-warn-script-location $pipWheelPath
if ($LASTEXITCODE -ne 0) { throw 'Pinned pip wheel bootstrap failed.' }

$requirementsLock = Join-Path $PSScriptRoot 'python-requirements.lock.txt'
if (-not (Test-Path -LiteralPath $requirementsLock -PathType Leaf)) {
    throw "Python dependency lock is missing: $requirementsLock"
}
& $pythonExe -m pip install --require-hashes --only-binary=:all: --no-warn-script-location `
    --requirement $requirementsLock
if ($LASTEXITCODE -ne 0) { throw 'Python dependencies failed to install.' }

$sidecarSource = Join-Path $PSScriptRoot '..\src\DeepExcel.Sidecar'
$sidecarDestination = Join-Path $OutputDir 'sidecar'
New-Item -ItemType Directory -Path $sidecarDestination -Force | Out-Null
# 与 _compile_only.ps1 相同：整棵侧车源码树（顶层 .py + perception 等子包），只排除测试与缓存。
# 只复制顶层 *.py 会漏掉子包。
Get-ChildItem -LiteralPath $sidecarSource -Force | Where-Object {
    $_.Name -notin @('tests', '__pycache__', '.pytest_cache') -and
    ($_.PSIsContainer -or $_.Extension -in @('.py', '.md', '.json'))
} | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $sidecarDestination -Recurse -Force
}

Get-ChildItem -LiteralPath $OutputDir -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force
Get-ChildItem -LiteralPath $OutputDir -Recurse -Directory -Filter 'pip-*.dist-info' -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force

& $pythonExe -c 'import claude_agent_sdk'
if ($LASTEXITCODE -ne 0) { throw 'claude-agent-sdk import validation failed.' }

# Validate the natural runtime import path. Do not insert sys.path here: doing
# so previously hid a package that always crashed for end users.
& $pythonExe -c 'import sidecar'
if ($LASTEXITCODE -ne 0) { throw 'Sidecar import validation failed.' }

Remove-Item -LiteralPath $TempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Embedded Python package ready: $OutputDir"
