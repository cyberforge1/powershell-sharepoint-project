# Filename: create-site-and-apply-template.ps1

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
$siteUrl = "https://cyberforge000.sharepoint.com/sites/YourNewSite"
$templatePath = $env:TEMPLATE_PATH
$owner = $env:OWNER_EMAIL

# Create a New Site
New-PnPSite -Type TeamSite -Title "Your New Site" -Alias "YourNewSite" -IsPublic -Owner $owner

# Connect to the new site
Connect-PnPOnline -Url $siteUrl -Credentials (Get-Credential)

# Apply the template
Invoke-PnPProvisioningTemplate -Path $templatePath
