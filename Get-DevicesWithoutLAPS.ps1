#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, ImportExcel

<#
.SYNOPSIS
    Finds hybrid Entra-joined devices that do not have a LAPS password stored in Entra ID.

.DESCRIPTION
    Connects to Microsoft Graph, retrieves all hybrid Azure AD joined devices,
    checks each for a Local Administrator Password Solution (LAPS) credential,
    and reports devices that are missing one. Active and stale devices are always
    included but reported in separate sections and Excel sheets.

.PARAMETER StaleDaysThreshold
    Number of days since last sign-in to consider a device stale. Default: 90.

.EXAMPLE
    .\Get-DevicesWithoutLAPS.ps1

.EXAMPLE
    .\Get-DevicesWithoutLAPS.ps1 -StaleDaysThreshold 180
#>

[CmdletBinding()]
param(
    [int]$StaleDaysThreshold = 90
)

$ErrorActionPreference = 'Stop'

# --- Logging setup ---
$logDir = 'C:\GI'
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile        = Join-Path $logDir "LAPSAudit_${timestamp}.log"
$transcriptFile = Join-Path $logDir "LAPSAudit_${timestamp}_transcript.log"

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
Write-Log "Stale threshold: $StaleDaysThreshold days"

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
    $allDevices = Get-MgDevice -Filter $filter -Property $select -All -CountVariable deviceCount `
        -ConsistencyLevel eventual
    Write-Log "Found $($allDevices.Count) hybrid-joined device(s)."
}
catch {
    Write-Log "Failed to retrieve devices: $_" -Level ERROR
    Stop-Transcript | Out-Null
    return
}

if ($allDevices.Count -eq 0) {
    Write-Log "No hybrid-joined devices found in this tenant." -Level WARN
    Stop-Transcript | Out-Null
    return
}

# --- Partition into active and stale ---
$cutoffDate    = (Get-Date).AddDays(-$StaleDaysThreshold)
$activeDevices = $allDevices | Where-Object {
    $_.ApproximateLastSignInDateTime -and $_.ApproximateLastSignInDateTime -gt $cutoffDate
}
$staleDevices  = $allDevices | Where-Object {
    -not $_.ApproximateLastSignInDateTime -or $_.ApproximateLastSignInDateTime -le $cutoffDate
}

Write-Log "Active devices (signed in within $StaleDaysThreshold days) : $($activeDevices.Count)"
Write-Log "Stale  devices (no sign-in in $StaleDaysThreshold+ days)   : $($staleDevices.Count)"

# --- Helper: check LAPS for a list of devices ---
function Get-DevicesWithoutLapsCredential {
    param(
        [object[]]$DeviceList,
        [string]$GroupLabel
    )

    $missing  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $counter  = 0
    $total    = $DeviceList.Count

    foreach ($device in $DeviceList) {
        $counter++
        $pct = [math]::Round(($counter / $total) * 100)
        Write-Progress -Activity "Checking LAPS credentials ($GroupLabel)" `
            -Status "$counter / $total ($pct%)" `
            -PercentComplete $pct `
            -CurrentOperation $device.DisplayName

        $hasLaps = $false
        try {
            $lapsCredential = Get-MgDeviceLocalCredential -Filter "deviceId eq '$($device.DeviceId)'" -ErrorAction Stop
            if ($lapsCredential) { $hasLaps = $true }
        }
        catch {
            $hasLaps = $false
        }

        if (-not $hasLaps) {
            $missing.Add([PSCustomObject]@{
                DeviceName = $device.DisplayName
                DeviceId   = $device.DeviceId
                ObjectId   = $device.Id
                OS         = $device.OperatingSystem
                OSVersion  = $device.OperatingSystemVersion
                Enabled    = $device.AccountEnabled
                LastSignIn = $device.ApproximateLastSignInDateTime
                TrustType  = $device.TrustType
            })
        }
    }

    Write-Progress -Activity "Checking LAPS credentials ($GroupLabel)" -Completed
    return $missing
}

# --- Check active devices ---
Write-Log "Checking LAPS status for $($activeDevices.Count) active device(s)..."
$activeResults = if ($activeDevices.Count -gt 0) {
    Get-DevicesWithoutLapsCredential -DeviceList $activeDevices -GroupLabel 'Active'
} else {
    [System.Collections.Generic.List[PSCustomObject]]::new()
}

# --- Check stale devices ---
Write-Log "Checking LAPS status for $($staleDevices.Count) stale device(s)..."
$staleResults = if ($staleDevices.Count -gt 0) {
    Get-DevicesWithoutLapsCredential -DeviceList $staleDevices -GroupLabel 'Stale'
} else {
    [System.Collections.Generic.List[PSCustomObject]]::new()
}

# --- Console summary ---
Write-Log "========== Results =========="
Write-Log "Total hybrid devices         : $($allDevices.Count)"
Write-Log "  Active checked             : $($activeDevices.Count)"
Write-Log "  Stale  checked             : $($staleDevices.Count)"
Write-Log "Active WITHOUT LAPS          : $($activeResults.Count)"
Write-Log "Stale  WITHOUT LAPS          : $($staleResults.Count)"
Write-Log "Total  WITHOUT LAPS          : $($activeResults.Count + $staleResults.Count)"

if ($activeResults.Count -gt 0) {
    Write-Log "--- Active devices missing LAPS ---" -Level WARN
    $activeResults | Format-Table -Property DeviceName, OS, OSVersion, Enabled, LastSignIn -AutoSize
}
else {
    Write-Log "All active hybrid-joined devices have LAPS credentials in Entra."
}

if ($staleResults.Count -gt 0) {
    Write-Log "--- Stale devices missing LAPS (no sign-in in $StaleDaysThreshold+ days) ---" -Level WARN
    $staleResults | Format-Table -Property DeviceName, OS, OSVersion, Enabled, LastSignIn -AutoSize
}
else {
    Write-Log "All stale hybrid-joined devices have LAPS credentials in Entra."
}

# --- Export to xlsx ---
$totalWithoutLaps = $activeResults.Count + $staleResults.Count
if ($totalWithoutLaps -gt 0) {
    $xlsxFile = Join-Path $logDir "LAPSAudit_${timestamp}.xlsx"
    try {
        $baseExcelParams = @{
            Path         = $xlsxFile
            AutoSize     = $true
            AutoFilter   = $true
            FreezeTopRow = $true
            BoldTopRow   = $true
            TableStyle   = 'Medium6'
        }

        if ($activeResults.Count -gt 0) {
            $activeResults | Export-Excel @baseExcelParams `
                -WorksheetName 'Active Without LAPS' `
                -Title         'LAPS Audit - Active Devices Without Entra LAPS' `
                -TitleBold
        }

        if ($staleResults.Count -gt 0) {
            $staleResults | Export-Excel @baseExcelParams `
                -WorksheetName 'Stale Without LAPS' `
                -Title         "LAPS Audit - Stale Devices Without Entra LAPS (>${StaleDaysThreshold}d)" `
                -TitleBold
        }

        # Summary sheet
        $summary = [PSCustomObject]@{
            'Report Date'                    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            'Tenant ID'                      = (Get-MgContext).TenantId
            'Stale Threshold (Days)'         = $StaleDaysThreshold
            'Total Hybrid Devices'           = $allDevices.Count
            'Active Devices'                 = $activeDevices.Count
            'Stale Devices'                  = $staleDevices.Count
            'Active Without LAPS'            = $activeResults.Count
            'Stale Without LAPS'             = $staleResults.Count
            'Total Without LAPS'             = $totalWithoutLaps
            'Active With LAPS'               = $activeDevices.Count - $activeResults.Count
            'Stale With LAPS'                = $staleDevices.Count  - $staleResults.Count
        }
        $summary | Export-Excel -Path $xlsxFile -WorksheetName 'Summary' -AutoSize -BoldTopRow

        Write-Log "Results exported to: $xlsxFile"
    }
    catch {
        Write-Log "Failed to export xlsx: $_" -Level ERROR
    }
}
else {
    Write-Log "All hybrid-joined devices (active and stale) have LAPS credentials in Entra."
}

Stop-Transcript | Out-Null
Write-Log "Script complete. Transcript: $transcriptFile"

# Return both result sets to the pipeline
[PSCustomObject]@{
    ActiveWithoutLAPS = $activeResults
    StaleWithoutLAPS  = $staleResults
}
