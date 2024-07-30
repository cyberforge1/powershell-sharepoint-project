# script-three-apply-template.ps1

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "Please run this script with PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

Get-Module | Remove-Module -Force

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

$siteAlias = "SecondTestSite"
$templatePath = $env:TEMPLATE_PATH

if (-not $siteAlias) {
    Write-Host "Site alias is required." -ForegroundColor Red
    exit 1
}
if (-not $templatePath) {
    Write-Host "Template path is required." -ForegroundColor Red
    exit 1
}
if (-not $adminCredentials) {
    Write-Host "Admin credentials are required." -ForegroundColor Red
    exit 1
}

try {
    Install-Module -Name PnP.PowerShell -Force -AllowClobber
    Import-Module PnP.PowerShell -Force
}
catch {
    Write-Host "Error installing or importing PnP.PowerShell module: $($_)" -ForegroundColor Red
    exit 1
}

$siteUrl = "https://cyberforge000.sharepoint.com/sites/$siteAlias"

try {
    Connect-PnPOnline -Url $siteUrl -Credentials $adminCredentials -WarningAction Ignore
}
catch {
    Write-Host "Error connecting to SharePoint site: $($_)" -ForegroundColor Red
    exit 1
}

try {
    Apply-PnPProvisioningTemplate -Path $templatePath
    Write-Host "Template applied successfully to $siteUrl."
}
catch {
    Write-Host "Error applying template to $($siteUrl): $($_)" -ForegroundColor Red
    $_.Exception | Format-List -Force
    exit 1
}
