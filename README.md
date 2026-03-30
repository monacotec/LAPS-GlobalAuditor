# LAPS-GlobalAuditor

PowerShell 7 script that audits a hybrid Entra ID (Azure AD) tenant for devices **missing a LAPS (Local Administrator Password Solution) password** stored in Entra.

## What it does

1. Connects to Microsoft Graph with the required scopes.
2. Retrieves all **hybrid Azure AD joined** devices (`trustType eq 'ServerAd'`).
3. Filters out stale devices (no sign-in in 90 days, configurable).
4. Checks each device for a LAPS credential via `Get-MgDeviceLocalCredential`.
5. Outputs a summary to the console and logs everything to `C:\GI`.

## Output

All output is written to `C:\GI\`:

| File | Description |
|------|-------------|
| `LAPSAudit_<timestamp>.log` | Structured log (`[datetime] [LEVEL] message`) |
| `LAPSAudit_<timestamp>_transcript.log` | Full PowerShell transcript |
| `LAPSAudit_<timestamp>.xlsx` | Excel report with **Devices Without LAPS** and **Summary** sheets |

## Prerequisites

- **PowerShell 7+**
- **Microsoft Graph PowerShell SDK** modules:
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Identity.DirectoryManagement`
- **ImportExcel** module

Run the included helper to install everything:

```powershell
.\Install-Prerequisites.ps1
```

## Required Entra Permissions

The script requests these Microsoft Graph scopes at sign-in:

| Scope | Purpose |
|-------|---------|
| `Device.Read.All` | Read device objects |
| `DeviceLocalCredential.Read.All` | Read LAPS credentials |

These are **delegated** permissions. The signed-in user must have sufficient Entra role (e.g., Cloud Device Administrator or Global Reader) or the scopes must be admin-consented for the tenant.

## Usage

```powershell
# Basic run (excludes stale devices)
.\Get-DevicesWithoutLAPS.ps1

# Include devices with no sign-in in the last 180 days
.\Get-DevicesWithoutLAPS.ps1 -IncludeStaleDevices -StaleDaysThreshold 180
```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-IncludeStaleDevices` | Switch | `$false` | Include devices that haven't signed in within the stale threshold |
| `-StaleDaysThreshold` | Int | `90` | Days since last sign-in to consider a device stale |

## License

MIT
