# Filename: create-multiple-sites-and-apply-templates.ps1

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
$templatePath = $env:TEMPLATE_PATH
$adminUrl = $env:SHAREPOINT_ADMIN_URL
$owner = $env:OWNER_EMAIL

# Connect to SharePoint Online
Connect-PnPOnline -Url $adminUrl -Credentials (Get-Credential)

# Create and apply template to 10,000 sites
for ($i=1; $i -le 10000; $i++) {
    $siteUrl = "https://cyberforge000.sharepoint.com/sites/Site$i"
    New-PnPSite -Type TeamSite -Title "Site $i" -Alias "Site$i" -IsPublic -Owner $owner
    Connect-PnPOnline -Url $siteUrl -Credentials (Get-Credential)
    Invoke-PnPProvisioningTemplate -Path $templatePath
}
