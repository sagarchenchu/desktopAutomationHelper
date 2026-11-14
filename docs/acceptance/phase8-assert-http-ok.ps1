<#
.SYNOPSIS
  Phase 8 Windows acceptance helper — assert Debug Mode HTTP start/inspect/verify/stop succeed.

.DESCRIPTION
  Cloud + Windows workers must treat a scenario as FAIL unless these flags are true:
    startHttpOk, inspectHttpOk, verifyHttpOk (when verify is exercised), stopHttpOk

  Historical PASS claims with inspectHttpOk=false or inconsistent SHAs are retracted.
  Re-run is DNU-25: one committed tip SHA; every result file must record that unchanged SHA.
  S5 must candidates->operator approve->authorize Allowed=true.
  S6 authorize denied when vision off (executionToken path).
  S7 destructive denial with vision enabled (not Vision:Enabled=false).

.PARAMETER BaseUrl
  Driver base URL (e.g. http://127.0.0.1:91xx).

.PARAMETER Token
  Bearer token from driver startup banner / probe.

.PARAMETER X / Y
  Screen point for inspect/hover.

.PARAMETER SkipVerify
  When set, verifyHttpOk is reported as $null (not required for scenarios that only inspect).

.PARAMETER AssertVisionAuthorizeDenied
  Scenario 6: also POST /vision/authorize and require Allowed=false while Vision:Enabled=false.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$BaseUrl,
    [Parameter(Mandatory = $true)][string]$Token,
    [Parameter(Mandatory = $true)][int]$X,
    [Parameter(Mandatory = $true)][int]$Y,
    [string]$OutputDirectory,
    [switch]$SkipVerify,
    [switch]$AssertVisionAuthorizeDenied
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
        $content = $null
        if ($_.Exception.Response) {
            try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = $null }
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                if ($stream) {
                    $reader = New-Object System.IO.StreamReader($stream)
                    $content = $reader.ReadToEnd()
                    $reader.Dispose()
                }
            } catch { $content = $null }
        }
        return @{ Ok = $false; StatusCode = $code; Content = $content; Error = "$_" }
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

if ($AssertVisionAuthorizeDenied) {
    # Round-3 / DNU-25: authorize consumes executionToken (not candidateId).
    $auth = Invoke-Driver -Method 'POST' -Path '/vision/authorize' -Body @{
        executionToken = 'vtx-acceptance-probe-unissued'
        action         = 'click'
        isDestructive  = $false
    }
    $allowed = $false
    if ($auth.Content) {
        try {
            $parsed = $auth.Content | ConvertFrom-Json
            if ($null -ne $parsed.allowed) { $allowed = [bool]$parsed.allowed }
            elseif ($null -ne $parsed.Allowed) { $allowed = [bool]$parsed.Allowed }
        }
        catch { $allowed = $true }
    }
    elseif ($auth.Ok) { $allowed = $true }
    $result.authorizeHttpOk = [bool]($auth.Ok -or ($null -ne $auth.Content))
    $result.authorizeAllowed = $allowed
    $result.authorizeHttpDenied = (-not $allowed)
    $result.authorizeStatus = $auth.StatusCode
    $result.authorizeError = $auth.Error
}

$required = @($result.startHttpOk, $result.inspectHttpOk, $result.stopHttpOk)
if (-not $SkipVerify) { $required += $result.verifyHttpOk }
if ($AssertVisionAuthorizeDenied) { $required += $result.authorizeHttpDenied }
$result.httpAllOk = -not ($required -contains $false)

# Fail closed: do not allow PASS when HTTP flags are false.
if (-not $result.httpAllOk) {
    Write-Error ("Phase 8 HTTP acceptance FAILED: " + ($result | ConvertTo-Json -Compress))
    exit 1
}

$result | ConvertTo-Json -Depth 6
exit 0
