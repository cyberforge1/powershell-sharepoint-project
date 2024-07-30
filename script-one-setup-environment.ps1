# Filename: setup-environment.ps1

# Load environment variables from .env file
Import-Module PSDotEnv
$envFilePath = ".\.env"
$envFileContent = Get-Content $envFilePath -ErrorAction Stop
$envFileContent | ForEach-Object {
    if ($_ -match "^\s*([^=]+?)\s*=\s*(.+?)\s*$") {
        [Environment]::SetEnvironmentVariable($matches[1], $matches[2])
    }
}

# Set Execution Policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

# Install SharePointPnPPowerShellOnline Module
Install-Module -Name SharePointPnPPowerShellOnline -Force -AllowClobber

# Import SharePointPnPPowerShellOnline Module
Import-Module SharePointPnPPowerShellOnline

# Connect to SharePoint Online
$adminUrl = $env:SHAREPOINT_ADMIN_URL
Connect-PnPOnline -Url $adminUrl -Credentials (Get-Credential)

# Verify connection using a PnP cmdlet
Get-PnPTenantSite
