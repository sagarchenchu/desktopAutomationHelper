[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repoRoot = Split-Path $PSScriptRoot -Parent
$nativeRoot = Join-Path $repoRoot 'src/DesktopVirtualDisplay'
$outputRoot = Join-Path $repoRoot 'artifacts/vdd'
$package = Join-Path $outputRoot 'package'
$nuget = (Get-Command nuget.exe -ErrorAction Stop).Source
$msbuildCommand = Get-Command msbuild.exe -ErrorAction SilentlyContinue
if ($msbuildCommand) { $msbuild = $msbuildCommand.Source }
else {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $msbuild = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
    if (-not $msbuild) { throw 'Visual Studio 2022 C++ Build Tools/MSBuild are required.' }
}
& $nuget restore (Join-Path $nativeRoot 'packages.config') -PackagesDirectory (Join-Path $nativeRoot 'packages') -NonInteractive
if ($LASTEXITCODE -ne 0) { throw 'Native SDK/WDK package restore failed.' }
# Clean only this script's generated output, so stale binaries cannot enter a package.
if (Test-Path $outputRoot) { Remove-Item $outputRoot -Recurse -Force }
New-Item $package -ItemType Directory -Force | Out-Null
foreach ($project in @('Driver/DesktopVirtualDisplay.vcxproj', 'Control/DesktopVddCtl.vcxproj', 'Tests/VddContracts.vcxproj')) {
    & $msbuild (Join-Path $nativeRoot $project) /m /t:Build /p:Configuration=Release /p:Platform=x64 /p:SignMode=Off /p:Inf2CatSkip=true /v:minimal
    if ($LASTEXITCODE -ne 0) { throw "Native build failed: $project" }
}
& (Join-Path $outputRoot 'tools/VddContracts.exe')
if ($LASTEXITCODE -ne 0) { throw 'Native contract tests failed.' }
& (Join-Path $outputRoot 'tools/DesktopVddCtl.exe') --help
if ($LASTEXITCODE -ne 0) { throw 'Published utility help failed.' }
$statusJson = & (Join-Path $outputRoot 'tools/DesktopVddCtl.exe') status
$statusExit = $LASTEXITCODE
if ($statusExit -notin @(0, 2)) { throw "Unexpected read-only status exit: $statusExit" }
$status = $statusJson | ConvertFrom-Json
if ($status.schemaVersion -ne 1) { throw 'Unexpected status protocol.' }
if ($statusExit -eq 0 -and -not $status.pipelineReady) { throw 'Status exit/readiness mismatch.' }
Write-Output $statusJson
foreach ($name in @('DesktopVirtualDisplay.dll', 'DesktopVirtualDisplay.inf')) {
    $source = Join-Path $outputRoot "driver/$name"
    if (-not (Test-Path $source)) { throw "Expected driver output missing: $source" }
    Copy-Item $source $package
}
Copy-Item (Join-Path $outputRoot 'tools/DesktopVddCtl.exe') $package
$wdk = Join-Path $nativeRoot 'packages/Microsoft.Windows.WDK.x64.10.0.26100.6584'
$inf2cat = Get-ChildItem $wdk -Filter Inf2Cat.exe -Recurse | Where-Object FullName -Match '[\\/]x86[\\/]' | Select-Object -First 1
if (-not $inf2cat) { $inf2cat = Get-ChildItem $wdk -Filter Inf2Cat.exe -Recurse | Select-Object -First 1 }
if (-not $inf2cat) { throw 'Inf2Cat was not found in the pinned WDK package.' }
& $inf2cat.FullName "/driver:$package" /os:ServerRS5_X64,10_RS5_X64,10_CO_X64 /uselocaltime
if ($LASTEXITCODE -ne 0) { throw 'Driver catalog validation/generation failed.' }
foreach ($framework in @('WinForms', 'Wpf')) {
    & dotnet publish (Join-Path $repoRoot "tests/DesktopRuntime.Fixtures.$framework/DesktopRuntime.Fixtures.$framework.csproj") `
        -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:PublishTrimmed=false -p:DebugType=None -o (Join-Path $package "fixtures/$framework")
    if ($LASTEXITCODE -ne 0) { throw "Fixture build failed: $framework" }
}
Copy-Item (Join-Path $nativeRoot 'LICENSE-MS-PL.txt') $package
Copy-Item (Join-Path $nativeRoot 'NOTICE.md') $package
Copy-Item (Join-Path $repoRoot 'docs/desktop-runtime-vdd-prototype.md') (Join-Path $package 'README.md')
Copy-Item (Join-Path $PSScriptRoot 'Install-VddPrototype.ps1') $package
Copy-Item (Join-Path $PSScriptRoot 'Sign-VddPrototype.ps1') $package
Copy-Item (Join-Path $repoRoot 'tests/DesktopRuntime.Acceptance') (Join-Path $package 'acceptance') -Recurse
@{
    schemaVersion = 1
    maturity = 'experimental-unsigned-prototype'
    driverVersion = '0.1.0.0'
    kitVersion = '10.0.26100.6584'
    minimumOsBuildIntent = 17763
    iddCxMinimum = '1.3'
    nativeBuildPassed = $true
    contractTestsPassed = $true
    deviceInstallationVerified = $false
    viewerIndependentExecutionVerified = $false
    productionSigned = $false
} | ConvertTo-Json | Set-Content (Join-Path $package 'build-evidence.json') -Encoding utf8
Get-ChildItem $package -File -Recurse | ForEach-Object {
    [pscustomobject]@{ file = [IO.Path]::GetRelativePath($package, $_.FullName); sha256 = (Get-FileHash $_.FullName -Algorithm SHA256).Hash }
} | ConvertTo-Json | Set-Content (Join-Path $package 'SHA256.json') -Encoding utf8
Write-Output "Unsigned prototype package: $package"
