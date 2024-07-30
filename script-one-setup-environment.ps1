# script-one-setup-environment.ps1

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

# Retrieve credentials from environment variables
$adminUsername = $env:SHAREPOINT_USERNAME
$adminPassword = $env:SHAREPOINT_PASSWORD | ConvertTo-SecureString -AsPlainText -Force
$adminCredentials = New-Object System.Management.Automation.PSCredential($adminUsername, $adminPassword)

# Connect to SharePoint Online
$adminUrl = $env:SHAREPOINT_ADMIN_URL
Connect-PnPOnline -Url $adminUrl -Credentials $adminCredentials

# Verify connection using a PnP cmdlet
Get-PnPTenantSite
