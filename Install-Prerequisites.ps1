#Requires -Version 7.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs the required PowerShell modules for LAPS-GlobalAuditor.

.DESCRIPTION
    Installs Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement,
    and ImportExcel modules from the PowerShell Gallery.
#>

[CmdletBinding()]
param()

$modules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'ImportExcel'
)

foreach ($mod in $modules) {
    if (Get-Module -ListAvailable -Name $mod) {
        Write-Host "[OK] $mod is already installed." -ForegroundColor Green
    }
    else {
        Write-Host "Installing $mod..." -ForegroundColor Cyan
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
        Write-Host "[OK] $mod installed." -ForegroundColor Green
    }
}

Write-Host "`nAll prerequisites installed. You can now run .\Get-DevicesWithoutLAPS.ps1" -ForegroundColor Green
