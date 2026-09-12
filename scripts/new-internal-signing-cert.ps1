#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$OutputDirectory,
    [switch]$ForceNew
)

$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $projectRoot '.local-signing' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$expectedDefault = [IO.Path]::GetFullPath((Join-Path $projectRoot '.local-signing'))
if ($OutputDirectory -ne $expectedDefault) {
    throw "Refusing to write internal signing material outside $expectedDefault"
}

$subject = 'CN=DeepExcel Internal Testing'
$friendlyName = 'DeepExcel Internal Test Code Signing'
$codeSigningOid = '1.3.6.1.5.5.7.3.3'
$minimumValidity = (Get-Date).AddDays(30)

function Test-InternalSigningCertificate {
    param([Security.Cryptography.X509Certificates.X509Certificate2]$Candidate)

    if (-not $Candidate -or $Candidate.Subject -ne $subject -or
        $Candidate.Issuer -ne $Candidate.Subject -or
        $Candidate.FriendlyName -ne $friendlyName -or
        -not $Candidate.HasPrivateKey -or
        $Candidate.NotBefore -gt (Get-Date) -or $Candidate.NotAfter -le $minimumValidity -or
        $Candidate.PublicKey.Oid.Value -ne '1.2.840.113549.1.1.1' -or
        $Candidate.SignatureAlgorithm.Value -ne '1.2.840.113549.1.1.11') {
        return $false
    }

    $eku = $Candidate.Extensions |
        Where-Object { $_.Oid.Value -eq '2.5.29.37' } |
        ForEach-Object { $_.EnhancedKeyUsages } |
        Where-Object { $_.Value -eq $codeSigningOid }
    $keyUsage = $Candidate.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.15' } | Select-Object -First 1
    $basicConstraints = $Candidate.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.19' } | Select-Object -First 1
    if (-not $eku -or -not $keyUsage -or
        -not ($keyUsage.KeyUsages -band [Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature) -or
        -not $basicConstraints -or $basicConstraints.CertificateAuthority) {
        return $false
    }

    $publicRsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($Candidate)
    $privateRsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Candidate)
    try {
        if (-not $publicRsa -or $publicRsa.KeySize -lt 3072 -or -not $privateRsa) {
            return $false
        }
        if ($privateRsa -is [Security.Cryptography.RSACng]) {
            return $privateRsa.Key.ExportPolicy -eq [Security.Cryptography.CngExportPolicies]::None
        }
        if ($privateRsa -is [Security.Cryptography.RSACryptoServiceProvider]) {
            return -not $privateRsa.CspKeyContainerInfo.Exportable
        }
        return $false
    } finally {
        if ($publicRsa) { $publicRsa.Dispose() }
        if ($privateRsa) { $privateRsa.Dispose() }
    }
}

$certificate = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { Test-InternalSigningCertificate $_ } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if ($ForceNew -or -not $certificate) {
    $certificate = New-SelfSignedCertificate `
        -Type Custom `
        -Subject $subject `
        -FriendlyName $friendlyName `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeyAlgorithm RSA `
        -KeyLength 3072 `
        -HashAlgorithm SHA256 `
        -KeyUsage DigitalSignature `
        -KeyExportPolicy NonExportable `
        -NotAfter (Get-Date).AddYears(3) `
        -TextExtension @(
            "2.5.29.37={text}$codeSigningOid",
            '2.5.29.19={critical}{text}CA=false'
        )
    Write-Host "Created internal code-signing certificate: $($certificate.Thumbprint)"
} else {
    Write-Host "Reusing internal code-signing certificate: $($certificate.Thumbprint)"
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$cerPath = Join-Path $OutputDirectory 'DeepExcel.Internal.cer'
$thumbprintPath = Join-Path $OutputDirectory 'certificate-thumbprint.txt'
Export-Certificate -Cert $certificate -FilePath $cerPath -Force | Out-Null
Set-Content -LiteralPath $thumbprintPath -Value $certificate.Thumbprint -Encoding ASCII

# Trust the public certificate for this Windows user so local verification of
# the self-signed Authenticode chain succeeds. The non-exportable private key
# remains only in CurrentUser\My on this development computer.
& certutil.exe -user -f -addstore Root $cerPath | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Failed to trust internal certificate in CurrentUser Root: $LASTEXITCODE" }
& certutil.exe -user -f -addstore TrustedPublisher $cerPath | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Failed to trust internal certificate in CurrentUser TrustedPublisher: $LASTEXITCODE" }

Write-Host "Public certificate: $cerPath"
Write-Host "Thumbprint file:   $thumbprintPath"
Write-Host "Private key:       CurrentUser certificate store only (non-exportable)"
