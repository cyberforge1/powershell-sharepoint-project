# script-two-create-site-and-apply-template.ps1

# Install and import the PnP PowerShell module (if not already installed)
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Install-Module -Name PnP.PowerShell -Force -AllowClobber
}
Import-Module PnP.PowerShell

# Define the site name
$siteName = "MySecondNewTestSite"
$siteAlias = "MySecondNewTestSite"

# Load environment variables from .env file
Import-Module PSDotEnv
$envFilePath = ".\.env"
$envFileContent = Get-Content $envFilePath -ErrorAction Stop
$envFileContent | ForEach-Object {
    if ($_ -match "^\s*([^=]+?)\s*=\s*(.+?)\s*$") {
        [Environment]::SetEnvironmentVariable($matches[1], $matches[2])
    }
}

# Variables
$siteUrl = $env:SHAREPOINT_SITE_URL
$templatePath = $env:TEMPLATE_PATH
$owner = $env:OWNER_EMAIL

# Retrieve credentials from environment variables
$adminUsername = $env:SHAREPOINT_USERNAME
$adminPassword = $env:SHAREPOINT_PASSWORD | ConvertTo-SecureString -AsPlainText -Force
$adminCredentials = New-Object System.Management.Automation.PSCredential($adminUsername, $adminPassword)

# Connect to SharePoint Online
Connect-PnPOnline -Url $siteUrl -Credentials $adminCredentials

# Check if the site alias already exists
try {
    $existingSite = Get-PnPTenantSite -Detailed | Where-Object { $_.Url -eq "$siteUrl/sites/$siteAlias" }
    if ($null -ne $existingSite) {
        Write-Host "A site with the alias '$siteAlias' already exists. Please choose a different alias."
        exit 1
    }
}
catch {
    Write-Host "Error checking for existing site: $_"
    exit 1
}

# Create a New Site
try {
    New-PnPSite -Type TeamSite -Title $siteName -Alias $siteAlias -IsPublic -Owner $owner
}
catch {
    Write-Host "Error creating site: $_"
    exit 1
}

# Apply the template
try {
    Apply-PnPProvisioningTemplate -Path $templatePath
}
catch {
    Write-Host "Error applying template: $_"
    exit 1
}
