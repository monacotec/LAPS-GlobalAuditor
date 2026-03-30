#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, ImportExcel

<#
.SYNOPSIS
    Finds hybrid Entra-joined devices that do not have a LAPS password stored in Entra ID.

.DESCRIPTION
    Connects to Microsoft Graph, retrieves all hybrid Azure AD joined devices,
    checks each for a Local Administrator Password Solution (LAPS) credential,
    and reports devices that are missing one.

.PARAMETER IncludeStaleDevices
    Include devices that haven't signed in within the last 90 days.
    By default, stale devices are excluded.

.PARAMETER StaleDaysThreshold
    Number of days since last sign-in to consider a device stale. Default: 90.

.EXAMPLE
    .\Get-DevicesWithoutLAPS.ps1

.EXAMPLE
    .\Get-DevicesWithoutLAPS.ps1 -IncludeStaleDevices -StaleDaysThreshold 180
#>

[CmdletBinding()]
param(
    [switch]$IncludeStaleDevices,

    [int]$StaleDaysThreshold = 90
)

$ErrorActionPreference = 'Stop'

# --- Logging setup ---
$logDir = 'C:\GI'
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path $logDir "LAPSAudit_$timestamp.log"
$transcriptFile = Join-Path $logDir "LAPSAudit_$timestamp_transcript.log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $logFile -Value $entry
    switch ($Level) {
        'WARN'  { Write-Host $Message -ForegroundColor Yellow }
        'ERROR' { Write-Host $Message -ForegroundColor Red }
        default { Write-Host $Message -ForegroundColor Cyan }
    }
}

Start-Transcript -Path $transcriptFile -Append | Out-Null
Write-Log "Script started. Log file: $logFile"

# --- Connect to Graph ---
$requiredScopes = @(
    'Device.Read.All',
    'DeviceLocalCredential.Read.All'
)

Write-Log "Connecting to Microsoft Graph..."
try {
    $context = Get-MgContext
    if (-not $context) {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    }
    else {
        $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }
        if ($missingScopes) {
            Write-Log "Re-authenticating to acquire missing scopes: $($missingScopes -join ', ')" -Level WARN
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        }
    }
    $ctx = Get-MgContext
    Write-Log "Connected as: $($ctx.Account) | Tenant: $($ctx.TenantId)"
}
catch {
    Write-Log "Failed to connect to Microsoft Graph: $_" -Level ERROR
    Stop-Transcript | Out-Null
    return
}

# --- Get all hybrid-joined devices ---
Write-Log "Retrieving hybrid Entra-joined devices..."

$filter = "trustType eq 'ServerAd'"
$select = "id,displayName,deviceId,operatingSystem,operatingSystemVersion,approximateLastSignInDateTime,accountEnabled,trustType"

try {
    $devices = Get-MgDevice -Filter $filter -Property $select -All -CountVariable deviceCount `
        -ConsistencyLevel eventual
    Write-Log "Found $($devices.Count) hybrid-joined device(s)."
}
catch {
    Write-Log "Failed to retrieve devices: $_" -Level ERROR
    Stop-Transcript | Out-Null
    return
}

if ($devices.Count -eq 0) {
    Write-Log "No hybrid-joined devices found in this tenant." -Level WARN
    Stop-Transcript | Out-Null
    return
}

# --- Filter stale devices ---
$staleCount = 0
if (-not $IncludeStaleDevices) {
    $cutoffDate = (Get-Date).AddDays(-$StaleDaysThreshold)
    $activeDevices = $devices | Where-Object {
        $_.ApproximateLastSignInDateTime -and
        $_.ApproximateLastSignInDateTime -gt $cutoffDate
    }
    $staleCount = $devices.Count - $activeDevices.Count
    if ($staleCount -gt 0) {
        Write-Log "Excluded $staleCount stale device(s) (no sign-in in $StaleDaysThreshold days). Use -IncludeStaleDevices to include them." -Level WARN
    }
    $devices = $activeDevices
}

# --- Check each device for LAPS credential ---
Write-Log "Checking LAPS status for $($devices.Count) device(s)..."

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter = 0

foreach ($device in $devices) {
    $counter++
    $pct = [math]::Round(($counter / $devices.Count) * 100)
    Write-Progress -Activity "Checking LAPS credentials" -Status "$counter / $($devices.Count) ($pct%)" `
        -PercentComplete $pct -CurrentOperation $device.DisplayName

    $hasLaps = $false
    try {
        # Query the LAPS credential for this device by its deviceId
        $lapsCredential = Get-MgDeviceLocalCredential -Filter "deviceId eq '$($device.DeviceId)'" -ErrorAction Stop
        if ($lapsCredential) {
            $hasLaps = $true
        }
    }
    catch {
        # 404 or empty result means no LAPS credential
        $hasLaps = $false
    }

    if (-not $hasLaps) {
        $results.Add([PSCustomObject]@{
            DeviceName       = $device.DisplayName
            DeviceId         = $device.DeviceId
            ObjectId         = $device.Id
            OS               = $device.OperatingSystem
            OSVersion        = $device.OperatingSystemVersion
            Enabled          = $device.AccountEnabled
            LastSignIn       = $device.ApproximateLastSignInDateTime
            TrustType        = $device.TrustType
        })
    }
}

Write-Progress -Activity "Checking LAPS credentials" -Completed

# --- Output results ---
Write-Log "========== Results =========="
Write-Log "Total hybrid devices checked : $($devices.Count)"
Write-Log "Devices WITHOUT LAPS         : $($results.Count)"
Write-Log "Devices with LAPS            : $($devices.Count - $results.Count)"

if ($results.Count -gt 0) {
    Write-Log "Devices missing LAPS password in Entra:" -Level WARN
    $results | Format-Table -Property DeviceName, OS, OSVersion, Enabled, LastSignIn -AutoSize

    # Export to xlsx
    $xlsxFile = Join-Path $logDir "LAPSAudit_$timestamp.xlsx"
    try {
        $excelParams = @{
            Path          = $xlsxFile
            WorksheetName = 'Devices Without LAPS'
            AutoSize      = $true
            AutoFilter    = $true
            FreezeTopRow  = $true
            BoldTopRow    = $true
            TableStyle    = 'Medium6'
            Title         = "LAPS Audit - Devices Without Entra LAPS"
            TitleBold     = $true
        }
        $results | Export-Excel @excelParams

        # Add a summary sheet
        $summary = [PSCustomObject]@{
            'Report Date'              = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            'Tenant ID'                = (Get-MgContext).TenantId
            'Total Hybrid Devices'     = $devices.Count
            'Devices Without LAPS'     = $results.Count
            'Devices With LAPS'        = $devices.Count - $results.Count
            'Stale Devices Excluded'   = if ($IncludeStaleDevices) { 'N/A' } else { $staleCount }
            'Stale Threshold (Days)'   = $StaleDaysThreshold
        }
        $summary | Export-Excel -Path $xlsxFile -WorksheetName 'Summary' -AutoSize -BoldTopRow

        Write-Log "Results exported to: $xlsxFile"
    }
    catch {
        Write-Log "Failed to export xlsx: $_" -Level ERROR
    }
}
else {
    Write-Log "All hybrid-joined devices have LAPS credentials in Entra."
}

Stop-Transcript | Out-Null
Write-Log "Script complete. Transcript: $transcriptFile"

# Return objects to the pipeline for further processing
$results
