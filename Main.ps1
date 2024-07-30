# Main.ps1

# Import required modules
Import-Module PnP.PowerShell

# Read environment variables from .env file
$envVariables = Get-Content -Path "./.env" | Where-Object { $_ -match '=' } | ForEach-Object {
    $name, $value = $_ -split '=', 2
    [PSCustomObject]@{ Name = $name.Trim(); Value = $value.Trim() }
}

# Set environment variables
$envVariables | ForEach-Object {
    if ($_.Name -and $_.Value) {
        [Environment]::SetEnvironmentVariable($_.Name, $_.Value)
    }
}

# Prompt user for number of sites
$siteCount = Read-Host "Enter the number of sites to create"

Write-Host "Executing Connect-SharePoint.ps1 script..."
# Run the connection script
. .\Connect-SharePoint.ps1

Write-Host "Executing Create-Sites.ps1 script..."
# Run the site creation script
. .\Create-Sites.ps1 -siteCount $siteCount

Write-Host "Main script executed successfully."
