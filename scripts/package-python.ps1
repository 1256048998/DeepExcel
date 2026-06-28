# scripts/package-python.ps1
# 下载 Python embeddable 并安装 claude-agent-sdk + 依赖
# 用法: powershell -ExecutionPolicy Bypass -File scripts\package-python.ps1

param(
    [string]$PythonVersion = "3.11.9",
    [string]$OutputDir = "$PSScriptRoot\..\src\DeepExcel.AddIn\bin\Release\python",
    [string]$TempDir = "$env:TEMP\deepexcel-python-packaging"
)

$ErrorActionPreference = "Stop"

# --- 1. 准备目录 ---
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

if (Test-Path $OutputDir) { Remove-Item $OutputDir -Recurse -Force }
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# --- 2. 下载 Python embeddable zip（使用淘宝 npmmirror 镜像加速） ---
$arch = "amd64"
$zipName = "python-$PythonVersion-embed-$arch.zip"
$mirrors = @(
    "https://registry.npmmirror.com/-/binary/python/$PythonVersion/$zipName",
    "https://www.python.org/ftp/python/$PythonVersion/$zipName"
)
$zipPath = Join-Path $TempDir $zipName

$downloaded = $false
foreach ($url in $mirrors) {
    Write-Host "尝试下载 Python embeddable: $url"
    try {
        Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing -TimeoutSec 120
        $downloaded = $true
        Write-Host "下载成功（来源: $url）"
        break
    } catch {
        Write-Host "下载失败，尝试下一个镜像: $_"
    }
}
if (-not $downloaded) { throw "所有镜像均下载失败" }

# --- 3. 解压到输出目录 ---
Write-Host "解压到 $OutputDir"
Expand-Archive -Path $zipPath -DestinationPath $OutputDir -Force

# --- 4. 启用 pip (取消 _pth 中的 site-packages 注释, 下载 get-pip.py) ---
$pythonExe = Join-Path $OutputDir "python.exe"
$pthFile = Get-ChildItem -Path $OutputDir -Filter "python*._pth" | Select-Object -First 1

if ($pthFile) {
    Write-Host "启用 site-packages: $($pthFile.FullName)"
    $content = Get-Content $pthFile.FullName
    $content = $content | ForEach-Object {
        if ($_ -match "^#import site") { "import site" } else { $_ }
    }
    Set-Content -Path $pthFile.FullName -Value $content
}

# get-pip.py
$getPipMirrors = @(
    "https://bootstrap.pypa.io/get-pip.py",
    "https://pypi.tuna.tsinghua.edu.cn/packages/source/g/get-pip/get-pip-24.0.tar.gz"
)
$getPipPath = Join-Path $TempDir "get-pip.py"
$gpDownloaded = $false
foreach ($url in $getPipMirrors) {
    Write-Host "尝试下载 get-pip: $url"
    try {
        Invoke-WebRequest -Uri $url -OutFile $getPipPath -UseBasicParsing -TimeoutSec 60
        $gpDownloaded = $true
        break
    } catch {
        Write-Host "下载失败: $_"
    }
}
if (-not $gpDownloaded) { throw "get-pip.py 下载失败" }

Write-Host "安装 pip（使用清华镜像）"
& $pythonExe $getPipPath --no-warn-script-location --index-url https://pypi.tuna.tsinghua.edu.cn/simple
if ($LASTEXITCODE -ne 0) { throw "pip 安装失败" }

# --- 5. 安装 claude-agent-sdk 及依赖（使用清华镜像） ---
$tsinghuaIndex = "https://pypi.tuna.tsinghua.edu.cn/simple"
Write-Host "安装 claude-agent-sdk==0.2.109（清华镜像）"
& $pythonExe -m pip install --no-warn-script-location -i $tsinghuaIndex `
    "claude-agent-sdk==0.2.109" `
    "anyio>=4.0" `
    "httpx>=0.27" `
    "pydantic>=2.0"
if ($LASTEXITCODE -ne 0) { throw "claude-agent-sdk 安装失败" }

# --- 6. 复制 sidecar 业务代码 ---
$sidecarSrc = "$PSScriptRoot\..\src\DeepExcel.Sidecar"
$sidecarDest = Join-Path $OutputDir "sidecar"
Write-Host "复制 sidecar 代码到 $sidecarDest"

New-Item -ItemType Directory -Path $sidecarDest -Force | Out-Null
Copy-Item -Path "$sidecarSrc\*.py" -Destination $sidecarDest -Force
if (Test-Path "$sidecarSrc\tests") {
    Copy-Item -Path "$sidecarSrc\tests" -Destination $sidecarDest -Recurse -Force
}

# --- 7. 清理 pip 缓存与 __pycache__ ---
Write-Host "清理缓存"
Get-ChildItem -Path $OutputDir -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force
Get-ChildItem -Path $OutputDir -Recurse -Directory -Filter "*.dist-info" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match "^pip-" } |
    Remove-Item -Recurse -Force

# --- 8. 验证 ---
Write-Host "验证安装"
$verifyCmd = 'import claude_agent_sdk; print("claude-agent-sdk:", claude_agent_sdk.__version__)'
& $pythonExe -c $verifyCmd
if ($LASTEXITCODE -ne 0) { throw "claude-agent-sdk 导入验证失败" }

$sidecarVerifyPath = $sidecarDest.Replace("'", "''")
$verifySidecarCmd = "import sys; sys.path.insert(0, r'$sidecarVerifyPath'); import sidecar; print('sidecar OK')"
& $pythonExe -c $verifySidecarCmd
if ($LASTEXITCODE -ne 0) { throw "sidecar 导入验证失败" }

# --- 9. 清理临时目录 ---
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=== Python 打包完成 ==="
Write-Host "输出目录: $OutputDir"
Write-Host "Python exe: $pythonExe"
Write-Host "Sidecar: $sidecarDest"
Write-Host ""
Write-Host "DeepExcel 加载项将通过 python.exe sidecar\sidecar.py 启动"
