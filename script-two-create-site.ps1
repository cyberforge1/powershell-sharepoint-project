
# script-two-create-site.ps1

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "Please run this script with PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

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

$siteName = $env:NEW_SITE_NAME
$siteAlias = $env:SITE_ALIAS
$siteUrl = $env:SHAREPOINT_SITE_URL
$owner = $env:OWNER_EMAIL

Write-Host "siteUrl: $siteUrl"
Write-Host "owner: $owner"

$modules = @('PnP.PowerShell')
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -AllowClobber
    }
}

Import-Module PnP.PowerShell

$adminUsername = $env:SHAREPOINT_USERNAME
$adminPassword = $env:SHAREPOINT_PASSWORD | ConvertTo-SecureString -AsPlainText -Force
$adminCredentials = [pscredential]::new($adminUsername, $adminPassword)

Connect-PnPOnline -Url $siteUrl -Credentials $adminCredentials -WarningAction Ignore

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

try {
    New-PnPSite -Type TeamSite -Title $siteName -Alias $siteAlias -IsPublic -Owner $owner
    Write-Host "Created site $siteUrl/sites/$siteAlias"
}
catch {
    Write-Host "Error creating site: $_"
    exit 1
}
