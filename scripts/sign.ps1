#requires -Version 5.1
<#
.SYNOPSIS
    Authenticode-sign one or more binaries with SHA256 + RFC3161 timestamp.

.DESCRIPTION
    Single entry point for code signing, used by:
      * deploy/DeepExcel.Setup.iss  ->  SignTool=deepsign=powershell.exe ... -File "..\scripts\sign.ps1" $f
      * scripts/package_release.py  ->  signs the built DLL and the final Setup.exe

    It locates signtool.exe automatically, signs each file with the certificate
    configured via environment variables, and stamps an RFC3161 timestamp so the
    signature survives past the code-signing cert's expiry.

    Configure exactly one certificate source:
      1. Local PFX : $env:DEEPEXCEL_PFX  (+ $env:DEEPEXCEL_PFX_PASS)
      2. Current-user certificate store : $env:DEEPEXCEL_CERT_THUMBPRINT
         (used only by the explicitly labelled internal-test build)
      3. Azure Trusted Signing : $env:USE_AZURE_TRUSTED_SIGNING=1
         (requires AZTS_TENANT_ID / AZTS_CLIENT_ID / AZTS_CLIENT_SECRET / AZTS_ENDPOINT)

    If NO certificate is configured, it prints a clear warning and returns 0 so
    DEV builds still succeed. PRODUCTION CI MUST set a cert, otherwise the build
    ships UNSIGNED (SmartScreen/antivirus will block or scare users).

.EXAMPLE
    .\sign.ps1 dist\DeepExcel.Setup.exe
    .\sign.ps1 build\DeepExcel.AddIn.dll dist\DeepExcel.Setup.exe
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)]
    [string[]]$Files
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 1. Locate signtool.exe
# ---------------------------------------------------------------------------
function Find-SignTool {
    if ($env:SIGNTOOL -and (Test-Path $env:SIGNTOOL)) { return $env:SIGNTOOL }

    $roots = @($env:ProgramFiles, ${env:ProgramFiles(x86)})
    foreach ($root in $roots) {
        if (-not $root) { continue }
        $hit = Get-ChildItem -Path $root -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
               Where-Object { $_.FullName -match 'x64|x86' } | Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    # Fall back to vswhere (Visual Studio / Windows SDK)
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $sdkDir = & $vswhere -latest -products * -requires Microsoft.Component.Windows.SDK `
                    -property installationPath 2>$null
        if ($sdkDir) {
            $hit = Get-ChildItem -Path $sdkDir -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
                   Select-Object -First 1
            if ($hit) { return $hit.FullName }
        }
    }
    # Last resort: rely on PATH
    $onPath = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }
    return $null
}

# ---------------------------------------------------------------------------
# 2. Determine timestamp server (RFC3161)
# ---------------------------------------------------------------------------
$tsRfc3161 = $env:TIMESTAMP_RFC3161
if (-not $tsRfc3161) { $tsRfc3161 = 'http://timestamp.digicert.com' }

# ---------------------------------------------------------------------------
# 3. Decide certificate source
# ---------------------------------------------------------------------------
$useAzure = ($env:USE_AZURE_TRUSTED_SIGNING -eq '1') -or ($env:USE_AZURE_TRUSTED_SIGNING -eq 'true')
$pfx = $env:DEEPEXCEL_PFX
$pfxPass = $env:DEEPEXCEL_PFX_PASS
$certThumbprint = ($env:DEEPEXCEL_CERT_THUMBPRINT -replace '\s', '').ToUpperInvariant()
$configuredSources = @(
    [bool]$pfx,
    [bool]$certThumbprint,
    [bool]$useAzure
) | Where-Object { $_ }

if ($configuredSources.Count -gt 1) {
    Write-Error '[sign] Configure exactly one signing source; PFX, certificate thumbprint, and Azure cannot be mixed.'
    exit 1
}

if (-not $useAzure -and -not $pfx -and -not $certThumbprint) {
    Write-Warning "[sign] No code-signing certificate configured (set DEEPEXCEL_PFX, DEEPEXCEL_CERT_THUMBPRINT, or USE_AZURE_TRUSTED_SIGNING=1)."
    Write-Warning "[sign] Shipping UNSIGNED build. SmartScreen/antivirus may block end users. Returning success for dev builds."
    exit 0
}

# ---------------------------------------------------------------------------
# 4. Sign each file
# ---------------------------------------------------------------------------
if (-not $pfx -and $certThumbprint) {
    if ($certThumbprint -notmatch '^[0-9A-F]{40}$') {
        Write-Error '[sign] DEEPEXCEL_CERT_THUMBPRINT must be a 40-character SHA-1 certificate thumbprint.'
        exit 1
    }
    # Use the .NET store API. Nested build processes may start PowerShell
    # without the Cert: provider drive, even though the Windows certificate
    # store and private key are available to the same user.
    $store = New-Object Security.Cryptography.X509Certificates.X509Store(
        'My', [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    try {
        $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $certificate = $store.Certificates |
            Where-Object { ($_.Thumbprint -replace '\s', '').ToUpperInvariant() -eq $certThumbprint } |
            Select-Object -First 1
    } finally {
        $store.Close()
    }
    if (-not $certificate -or -not $certificate.HasPrivateKey) {
        Write-Error "[sign] Current-user signing certificate/private key not found: $certThumbprint"
        exit 1
    }
    $hasCodeSigningEku = $certificate.EnhancedKeyUsageList |
        Where-Object { [string]$_.ObjectId -eq '1.3.6.1.5.5.7.3.3' }
    if (-not $hasCodeSigningEku) {
        Write-Error "[sign] Certificate is not valid for code signing: $certThumbprint"
        exit 1
    }
    # CODEX and other build hosts may supply a PowerShell 7-only PSModulePath
    # to Windows PowerShell. Import the inbox security module by absolute path
    # so Set-AuthenticodeSignature remains available in nested builds.
    $securityModule = Join-Path $env:WINDIR `
        'System32\WindowsPowerShell\v1.0\Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1'
    Import-Module -Name $securityModule -Force -ErrorAction Stop
    foreach ($f in $Files) {
        if (-not (Test-Path -LiteralPath $f -PathType Leaf)) {
            Write-Error "[sign] File not found: $f"
            exit 1
        }
        Write-Host "[sign] Internal certificate store signing: $f"
        $result = Set-AuthenticodeSignature -LiteralPath $f -Certificate $certificate `
            -HashAlgorithm SHA256 -TimestampServer $tsRfc3161 -IncludeChain All
        if ($result.Status -ne 'Valid') {
            Write-Error "[sign] FAILED: $f ($($result.Status): $($result.StatusMessage))"
            exit 1
        }
        Write-Host "[sign]   OK ($certThumbprint)"
    }
    exit 0
}

if ($useAzure) {
    # Azure Trusted Signing via the dotnet tool 'Azure.CodeSigning.DotNetTool'
    $tool = Get-Command azuresigntool -ErrorAction SilentlyContinue
    if (-not $tool) {
        Write-Warning "[sign] Azure Trusted Signing requested but 'azuresigntool' not found. Install: dotnet tool install --global Azure.CodeSigning.DotNetTool"
        exit 1
    }
    $endpoint = $env:AZTS_ENDPOINT
    $tenant   = $env:AZTS_TENANT_ID
    $client   = $env:AZTS_CLIENT_ID
    $secret   = $env:AZTS_CLIENT_SECRET
    if (-not ($endpoint -and $tenant -and $client -and $secret)) {
        Write-Warning "[sign] Azure Trusted Signing requires AZTS_ENDPOINT / AZTS_TENANT_ID / AZTS_CLIENT_ID / AZTS_CLIENT_SECRET."
        exit 1
    }
    foreach ($f in $Files) {
        Write-Host "[sign] Azure Trusted Signing: $f"
        & azuresigntool sign -d "DeepExcel" -t $tenant -u $client -c $secret `
            --endpoint $endpoint -tr $tsRfc3161 --files "$f"
        if ($LASTEXITCODE -ne 0) { Write-Error "[sign] FAILED: $f"; exit $LASTEXITCODE }
    }
    exit 0
}

# --- Local PFX path ---
$signtool = Find-SignTool
if (-not $signtool) {
    Write-Error "[sign] signtool.exe not found. Install the Windows SDK / Visual Studio build tools, or set `$env:SIGNTOOL."
    exit 1
}
Write-Host "[sign] Using signtool: $signtool"

$tmpPass = $null
if (-not $pfxPass) {
    # Prompt only when running interactively (CI should pass the secret via env)
    $secure = Read-Host -AsSecureString "Enter PFX password"
    $tmpPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
    $pfxPass = $tmpPass
}

foreach ($f in $Files) {
    if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { Write-Error "[sign] File not found: $f"; exit 1 }
    Write-Host "[sign] Signing (SHA256 + RFC3161): $f"
    & "$signtool" sign /fd SHA256 /f "$pfx" /p "$pfxPass" `
        /tr "$tsRfc3161" /td SHA256 /d "DeepExcel" "$f"
    if ($LASTEXITCODE -ne 0) { Write-Error "[sign] FAILED: $f"; exit $LASTEXITCODE }
    Write-Host "[sign]   OK"
}
exit 0
