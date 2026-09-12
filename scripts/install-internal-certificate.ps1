#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$CertificatePath,
    [string]$InstallerPath
)

$ErrorActionPreference = 'Stop'
# Do not rely on PSModulePath. Some launchers inherit a PowerShell 7-only
# module path into Windows PowerShell 5.1, preventing Authenticode cmdlets and
# the Cert: provider from loading during a double-click installation.
$securityModule = Join-Path $env:WINDIR `
    'System32\WindowsPowerShell\v1.0\Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1'
Import-Module -Name $securityModule -Force -ErrorAction Stop

$scriptDirectory = [IO.Path]::GetFullPath($PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($CertificatePath)) {
    $CertificatePath = Join-Path $scriptDirectory 'DeepExcel.Internal.cer'
}
if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
    $InstallerPath = Join-Path $scriptDirectory 'DeepExcel.Setup.INTERNAL.exe'
}
$CertificatePath = [IO.Path]::GetFullPath($CertificatePath)
if (-not (Test-Path -LiteralPath $CertificatePath -PathType Leaf)) {
    throw "DeepExcel internal certificate not found: $CertificatePath"
}
$InstallerPath = [IO.Path]::GetFullPath($InstallerPath)
if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) {
    throw "DeepExcel internal installer not found: $InstallerPath"
}

$certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
$codeSigningOid = '1.3.6.1.5.5.7.3.3'
$hasCodeSigningEku = $certificate.Extensions |
    Where-Object { $_.Oid.Value -eq '2.5.29.37' } |
    ForEach-Object { $_.EnhancedKeyUsages } |
    Where-Object { $_.Value -eq $codeSigningOid }
$keyUsage = $certificate.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.15' } | Select-Object -First 1
$basicConstraints = $certificate.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.19' } | Select-Object -First 1
$publicRsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($certificate)

if ($certificate.Subject -ne 'CN=DeepExcel Internal Testing' -or
    $certificate.Issuer -ne $certificate.Subject -or
    $certificate.HasPrivateKey -or
    -not $hasCodeSigningEku -or
    -not $keyUsage -or
    -not ($keyUsage.KeyUsages -band [Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature) -or
    -not $basicConstraints -or $basicConstraints.CertificateAuthority -or
    -not $publicRsa -or $publicRsa.KeySize -lt 3072 -or
    $certificate.SignatureAlgorithm.Value -ne '1.2.840.113549.1.1.11' -or
    $certificate.NotBefore -gt (Get-Date) -or
    $certificate.NotAfter -le (Get-Date)) {
    if ($publicRsa) { $publicRsa.Dispose() }
    throw 'Refusing to trust an unexpected, expired, or non-code-signing certificate.'
}
if ($publicRsa) { $publicRsa.Dispose() }

# Validate that the installer bundled beside this script was signed by the
# exact certificate about to be trusted. Before trust installation Windows may
# report the chain as untrusted, but the embedded signer certificate is still
# available for this exact-thumbprint comparison.
$signatureBeforeTrust = Get-AuthenticodeSignature -LiteralPath $InstallerPath
if (-not $signatureBeforeTrust.SignerCertificate -or
    $signatureBeforeTrust.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) {
    throw 'Installer signature does not match the bundled DeepExcel certificate.'
}
if ($signatureBeforeTrust.Status -in @('NotSigned', 'HashMismatch', 'NotSupported')) {
    throw "Installer has an invalid Authenticode signature: $($signatureBeforeTrust.Status)"
}

$rootPath = "Cert:\CurrentUser\Root\$($certificate.Thumbprint)"
$publisherPath = "Cert:\CurrentUser\TrustedPublisher\$($certificate.Thumbprint)"
$rootAlreadyTrusted = Test-Path -LiteralPath $rootPath
$publisherAlreadyTrusted = Test-Path -LiteralPath $publisherPath
$addedRoot = $false
$addedPublisher = $false
try {
    if (-not $rootAlreadyTrusted) {
        & certutil.exe -user -f -addstore Root $CertificatePath | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Failed to trust certificate in CurrentUser Root: $LASTEXITCODE" }
        $addedRoot = $true
    }
    if (-not $publisherAlreadyTrusted) {
        & certutil.exe -user -f -addstore TrustedPublisher $CertificatePath | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Failed to trust certificate in CurrentUser TrustedPublisher: $LASTEXITCODE" }
        $addedPublisher = $true
    }

    $signatureAfterTrust = Get-AuthenticodeSignature -LiteralPath $InstallerPath
    if ($signatureAfterTrust.Status -ne 'Valid' -or
        $signatureAfterTrust.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) {
        throw "Installer signature verification failed after trust installation: $($signatureAfterTrust.Status)"
    }
} catch {
    $originalError = $_
    $rollbackErrors = @()
    if ($addedPublisher) {
        & certutil.exe -user -delstore TrustedPublisher $certificate.Thumbprint | Out-Null
        if ($LASTEXITCODE -ne 0 -or (Test-Path -LiteralPath $publisherPath)) {
            $rollbackErrors += "TrustedPublisher cleanup failed (exit $LASTEXITCODE)"
        }
    }
    if ($addedRoot) {
        & certutil.exe -user -delstore Root $certificate.Thumbprint | Out-Null
        if ($LASTEXITCODE -ne 0 -or (Test-Path -LiteralPath $rootPath)) {
            $rollbackErrors += "Root cleanup failed (exit $LASTEXITCODE)"
        }
    }
    if ($rollbackErrors.Count -gt 0) {
        throw "Certificate trust failed: $($originalError.Exception.Message). ROLLBACK FAILED: $($rollbackErrors -join '; ')"
    }
    throw $originalError
}

Write-Host 'DeepExcel internal publisher certificate installed for the current user.' -ForegroundColor Green
Write-Host "Thumbprint: $($certificate.Thumbprint)"
Write-Host 'You can now run DeepExcel.Setup.INTERNAL.exe.'
