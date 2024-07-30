# script-one-setup-environment.ps1

# Ensure you are running PowerShell 7
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "Please run this script with PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

# Install necessary modules if they are not already installed
$modules = @('PSDotEnv', 'PnP.PowerShell')
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -AllowClobber
    }
}

# Import modules
Import-Module PSDotEnv
Import-Module PnP.PowerShell

# Load environment variables from .env file
$envFilePath = ".\.env"
if (Test-Path $envFilePath) {
    $envFileContent = Get-Content $envFilePath -ErrorAction Stop
    $envFileContent | ForEach-Object {
        if ($_ -match "^\s*([^=]+?)\s*=\s*(.+?)\s*$") {
            [Environment]::SetEnvironmentVariable($matches[1], $matches[2])
        }
    }
} else {
    Write-Host "Environment file not found at $envFilePath" -ForegroundColor Red
    exit 1
}

# Set Execution Policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

# Uninstall legacy SharePointPnPPowerShellOnline Module if installed
if (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline) {
    Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force
}

# Retrieve credentials from environment variables
$adminUsername = $env:SHAREPOINT_USERNAME
$adminPassword = $env:SHAREPOINT_PASSWORD | ConvertTo-SecureString -AsPlainText -Force
$adminCredentials = [pscredential]::new($adminUsername, $adminPassword)

# Connect to SharePoint Online
$adminUrl = $env:SHAREPOINT_ADMIN_URL
Connect-PnPOnline -Url $adminUrl -Credentials $adminCredentials -WarningAction Ignore

# Verify connection using a PnP cmdlet
Get-PnPTenantSite
