# script-one-setup-environment.ps1

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "Please run this script with PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

$modules = @('PSDotEnv', 'PnP.PowerShell')
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -AllowClobber
    }
}

Import-Module PSDotEnv
Import-Module PnP.PowerShell

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

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

if (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline) {
    Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force
}

$adminUsername = $env:SHAREPOINT_USERNAME
$adminPassword = $env:SHAREPOINT_PASSWORD | ConvertTo-SecureString -AsPlainText -Force
$adminCredentials = [pscredential]::new($adminUsername, $adminPassword)

$adminUrl = $env:SHAREPOINT_ADMIN_URL
Connect-PnPOnline -Url $adminUrl -Credentials $adminCredentials -WarningAction Ignore

Get-PnPTenantSite
