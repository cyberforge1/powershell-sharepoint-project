# Main.ps1

Import-Module PnP.PowerShell

# Load environment variables
$envVariables = Get-Content -Path "./.env" | Where-Object { $_ -match '=' } | ForEach-Object {
    $name, $value = $_ -split '=', 2
    [PSCustomObject]@{ Name = $name.Trim(); Value = $value.Trim() }
}

$envVariables | ForEach-Object {
    if ($_.Name -and $_.Value) {
        [Environment]::SetEnvironmentVariable($_.Name, $_.Value)
    }
}

# Print out environment variables to ensure they are set correctly
$envVariables | ForEach-Object {
    Write-Host "$($_.Name) = $($_.Value)"
}

# Ensure critical environment variables are set
if (-not $env:SHAREPOINT_SITE_URL) {
    Write-Error "SHAREPOINT_SITE_URL is not set."
    exit 1
}

if (-not $env:NEW_SITE_NAME) {
    Write-Error "NEW_SITE_NAME is not set."
    exit 1
}

if (-not $env:TEMPLATE_PATH) {
    Write-Error "TEMPLATE_PATH is not set."
    exit 1
}

# Prompt the user for the number of sites to create
$siteCount = Read-Host "Enter the number of sites to create"

# Ensure the siteCount is passed correctly
Write-Host "Executing Connect-SharePoint.ps1 script..."
. .\Connect-SharePoint.ps1

Write-Host "Executing Create-Sites.ps1 script with siteCount: $siteCount"
. .\Create-Sites.ps1 -siteCount $siteCount

Write-Host "Main script executed successfully."
