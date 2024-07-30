# script-three-apply-template.ps1

param (
    [string]$siteAlias,
    [string]$templatePath,
    [pscredential]$adminCredentials
)

# Ensure you are running PowerShell 7
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "Please run this script with PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

# Install necessary modules if they are not already installed
$modules = @('PnP.PowerShell')
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -AllowClobber
    }
}

# Import modules
Import-Module PnP.PowerShell

# Validate parameters
if (-not $siteAlias) {
    Write-Host "Site alias is required." -ForegroundColor Red
    exit 1
}

$siteUrl = "https://cyberforge000.sharepoint.com/sites/$siteAlias"

# Connect to the SharePoint site
Connect-PnPOnline -Url $siteUrl -Credentials $adminCredentials -WarningAction Ignore

# Apply the template
try {
    Apply-PnPProvisioningTemplate -Path $templatePath
    Write-Host "Template applied successfully to $siteUrl."
}
catch {
    Write-Host "Error applying template to $($siteUrl): $_"
    exit 1
}
