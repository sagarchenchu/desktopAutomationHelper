[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PackageDirectory,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{40}$')][string]$CertificateThumbprint,
    [ValidateSet('CurrentUser', 'LocalMachine')][string]$CertificateStore = 'CurrentUser',
    [Parameter(Mandatory)][string]$SignToolPath,
    [Parameter(Mandatory)][string]$Inf2CatPath,
    [uri]$TimestampUrl
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$package = (Resolve-Path $PackageDirectory).Path
$signTool = (Resolve-Path $SignToolPath).Path
$inf2cat = (Resolve-Path $Inf2CatPath).Path
$certificate = Get-Item "Cert:\$CertificateStore\My\$CertificateThumbprint"
if (-not $certificate.HasPrivateKey) { throw 'The selected signing certificate has no accessible private key.' }
if ($certificate.NotAfter -le (Get-Date) -or $certificate.NotBefore -gt (Get-Date)) { throw 'The selected certificate is not currently valid.' }
$arguments = @('sign', '/fd', 'SHA256', '/sha1', $CertificateThumbprint)
if ($CertificateStore -eq 'LocalMachine') { $arguments += '/sm' }
if ($TimestampUrl) {
    if ($TimestampUrl.Scheme -ne 'https') { throw 'Use an HTTPS timestamp endpoint.' }
    $arguments += @('/tr', $TimestampUrl.AbsoluteUri, '/td', 'SHA256')
}
# Sign binaries BEFORE catalog generation. Signing a DLL changes its catalog hash.
foreach ($name in @('DesktopVirtualDisplay.dll', 'DesktopVddCtl.exe')) {
    $file = Join-Path $package $name
    if (-not (Test-Path $file)) { throw "Missing file: $file" }
    & $signTool @arguments $file
    if ($LASTEXITCODE -ne 0) { throw "Signing failed: $file" }
}
& $inf2cat "/driver:$package" /os:ServerRS5_X64,10_RS5_X64,10_CO_X64 /uselocaltime
if ($LASTEXITCODE -ne 0) { throw 'Catalog regeneration failed.' }
$catalog = Join-Path $package 'DesktopVirtualDisplay.cat'
& $signTool @arguments $catalog
if ($LASTEXITCODE -ne 0) { throw 'Catalog signing failed.' }
foreach ($name in @('DesktopVirtualDisplay.dll', 'DesktopVirtualDisplay.inf')) {
    & $signTool verify /pa /v /c $catalog (Join-Path $package $name)
    if ($LASTEXITCODE -ne 0) { throw "Catalog membership/trust verification failed: $name" }
}
@{ schemaVersion = 1; subject = $certificate.Subject; thumbprint = $CertificateThumbprint;
    signedAtUtc = [DateTime]::UtcNow.ToString('O'); targetMachineAcceptanceVerified = $false
} | ConvertTo-Json | Set-Content (Join-Path $package 'signing-evidence.json') -Encoding utf8
Get-ChildItem $package -File -Recurse | Where-Object Name -NE 'SHA256.json' | ForEach-Object {
    [pscustomobject]@{ file = [IO.Path]::GetRelativePath($package, $_.FullName); sha256 = (Get-FileHash $_.FullName -Algorithm SHA256).Hash }
} | ConvertTo-Json | Set-Content (Join-Path $package 'SHA256.json') -Encoding utf8
Write-Output 'Signed package prepared. Verify acceptance under the target machine policy before installation.'
