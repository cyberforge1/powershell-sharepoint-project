# script-four-test-site-access.ps1

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

$adminUsername = $env:SHAREPOINT_USERNAME
$adminPassword = $env:SHAREPOINT_PASSWORD | ConvertTo-SecureString -AsPlainText -Force
$adminCredentials = [pscredential]::new($adminUsername, $adminPassword)

$siteAlias = "NewTestSite"
$siteUrl = "https://cyberforge000.sharepoint.com/sites/$siteAlias"


$modules = @('PnP.PowerShell')
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -AllowClobber
    }
}

Import-Module PnP.PowerShell -Force

try {
    Connect-PnPOnline -Url $siteUrl -Credentials $adminCredentials -WarningAction Ignore
    Write-Host "Connected to SharePoint site: $siteUrl"
}
catch {
    Write-Host "Error connecting to SharePoint site: $_" -ForegroundColor Red
    exit 1
}

try {
    $siteTitle = Get-PnPWeb
    Write-Host "Site Title: $($siteTitle.Title)"
}
catch {
    Write-Host "Error performing site check: $_" -ForegroundColor Red
    exit 1
}


