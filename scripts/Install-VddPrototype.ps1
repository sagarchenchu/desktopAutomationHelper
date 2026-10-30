#Requires -RunAsAdministrator
[CmdletBinding(SupportsShouldProcess)]
param([Parameter(Mandatory)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$package = (Resolve-Path $PackageDirectory).Path
$inf = Join-Path $package 'DesktopVirtualDisplay.inf'
$dll = Join-Path $package 'DesktopVirtualDisplay.dll'
$catalog = Join-Path $package 'DesktopVirtualDisplay.cat'
foreach ($file in @($inf, $dll, $catalog)) {
    if (-not (Test-Path $file -PathType Leaf)) { throw "Missing package file: $file" }
}
$signature = Get-AuthenticodeSignature $catalog
if ($signature.Status -ne 'Valid') {
    throw 'A trusted signed catalog is required. This installer does not change certificate trust, Secure Boot, or test-signing settings.'
}
Write-Output "Catalog publisher: $($signature.SignerCertificate.Subject)"
Write-Output "Catalog signer thumbprint: $($signature.SignerCertificate.Thumbprint)"
if ($PSCmdlet.ShouldProcess($inf, 'Stage the signed prototype driver package in the Windows Driver Store')) {
    & "$env:SystemRoot/System32/pnputil.exe" /add-driver $inf
    $result = $LASTEXITCODE
    if ($result -notin @(0, 3010)) { throw "PnPUtil failed with exit code $result. Inspect its output and SetupAPI.dev.log." }
    if ($result -eq 3010) { Write-Warning 'Windows reports that a restart is required.' }
    Write-Output 'Package staged. No permanent display device was created. Use DesktopVddCtl hold to acquire a temporary display.'
}
