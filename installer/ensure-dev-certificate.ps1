# Ensures the development code-signing certificate is in CurrentUser\My.
# The private key stays in the Windows certificate store. This script does not write a .pfx.

$ErrorActionPreference = "Stop"
$subject = "CN=Adithya Padmanabhan"
$csproj = Join-Path (Split-Path $PSScriptRoot -Parent) "src\CAT\CAT.csproj"
$text = Get-Content -LiteralPath $csproj -Raw
if ($text -notmatch 'ManifestCertificateThumbprint>([A-Fa-f0-9]{40})<') {
    throw "ManifestCertificateThumbprint was not found in CAT.csproj."
}
$expected = $Matches[1].ToUpperInvariant()

$cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $subject -and $_.HasPrivateKey -and $_.Thumbprint.ToUpperInvariant() -eq $expected } |
    Select-Object -First 1

if ($cert) {
    Write-Host "Development certificate is present: $subject"
    Write-Host "Thumbprint: $($cert.Thumbprint)"
    return
}

$other = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $subject -and $_.HasPrivateKey }
if ($other) {
    throw "A certificate for $subject exists, but its thumbprint does not match CAT.csproj ($expected). Do not create a second certificate. Export is not performed by this script."
}

throw @"
Certificate $subject ($expected) is not in Cert:\CurrentUser\My on this PC.
Create it in PowerShell (do not save a .pfx into the repo):

  New-SelfSignedCertificate -Type CodeSigningCert -Subject '$subject' -FriendlyName 'Adithya Padmanabhan VSTO development' -CertStoreLocation 'Cert:\CurrentUser\My' -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -KeySpec Signature -NotAfter (Get-Date).AddYears(5)

If the new thumbprint differs, put that thumbprint in ManifestCertificateThumbprint in src\CAT\CAT.csproj.
"@
