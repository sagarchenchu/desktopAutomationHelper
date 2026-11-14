<#
.SYNOPSIS
  Phase 8 Windows acceptance helper — assert Debug Mode HTTP start/inspect/verify/stop succeed.

.DESCRIPTION
  Cloud + Windows workers must treat a scenario as FAIL unless these flags are true:
    startHttpOk, inspectHttpOk, verifyHttpOk (when verify is exercised), stopHttpOk

  Historical S1–S9 on tip 28b8712 claimed PASS while inspectHttpOk/stopHttpOk were false
  (IntPtr STJ 500). That claim is retracted. Re-run on the post-review tip (DNU-23) and only
  mark overallPass=true when HTTP flags are OK.

.PARAMETER BaseUrl
  Driver base URL (e.g. http://127.0.0.1:91xx).

.PARAMETER Token
  Bearer token from driver startup banner / probe.

.PARAMETER X / Y
  Screen point for inspect/hover.

.PARAMETER SkipVerify
  When set, verifyHttpOk is reported as $null (not required for scenarios that only inspect).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$BaseUrl,
    [Parameter(Mandatory = $true)][string]$Token,
    [Parameter(Mandatory = $true)][int]$X,
    [Parameter(Mandatory = $true)][int]$Y,
    [string]$OutputDirectory,
    [switch]$SkipVerify
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-Driver {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body = $null
    )
    $headers = @{ Authorization = "Bearer $Token" }
    $uri = ($BaseUrl.TrimEnd('/') + $Path)
    try {
        if ($null -eq $Body) {
            $resp = Invoke-WebRequest -Uri $uri -Method $Method -Headers $headers -UseBasicParsing
        }
        else {
            $json = $Body | ConvertTo-Json -Depth 8 -Compress
            $resp = Invoke-WebRequest -Uri $uri -Method $Method -Headers $headers -ContentType 'application/json' -Body $json -UseBasicParsing
        }
        return @{ Ok = ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 300); StatusCode = [int]$resp.StatusCode; Content = $resp.Content; Error = $null }
    }
    catch {
        $code = $null
        if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
        return @{ Ok = $false; StatusCode = $code; Content = $null; Error = "$_" }
    }
}

$outDir = if ($OutputDirectory) { $OutputDirectory } else { $null }
$start = Invoke-Driver -Method 'POST' -Path ("/debug/start" + $(if ($outDir) { "?outputDirectory=$([uri]::EscapeDataString($outDir))" } else { '' }))
$inspect = Invoke-Driver -Method 'POST' -Path '/debug/inspect' -Body @{ x = $X; y = $Y; writeArtifact = $true }
$verify = $null
if (-not $SkipVerify) {
    $verify = Invoke-Driver -Method 'POST' -Path '/debug/verify'
}
$stop = Invoke-Driver -Method 'POST' -Path '/debug/stop'

$result = [ordered]@{
    startHttpOk   = [bool]$start.Ok
    inspectHttpOk = [bool]$inspect.Ok
    verifyHttpOk  = if ($SkipVerify) { $null } else { [bool]$verify.Ok }
    stopHttpOk    = [bool]$stop.Ok
    startStatus   = $start.StatusCode
    inspectStatus = $inspect.StatusCode
    verifyStatus  = if ($SkipVerify) { $null } else { $verify.StatusCode }
    stopStatus    = $stop.StatusCode
    inspectError  = $inspect.Error
    verifyError   = if ($SkipVerify) { $null } else { $verify.Error }
    httpAllOk     = $false
}

$required = @($result.startHttpOk, $result.inspectHttpOk, $result.stopHttpOk)
if (-not $SkipVerify) { $required += $result.verifyHttpOk }
$result.httpAllOk = -not ($required -contains $false)

# Fail closed: do not allow PASS when HTTP flags are false.
if (-not $result.httpAllOk) {
    Write-Error ("Phase 8 HTTP acceptance FAILED: " + ($result | ConvertTo-Json -Compress))
    exit 1
}

$result | ConvertTo-Json -Depth 6
exit 0
