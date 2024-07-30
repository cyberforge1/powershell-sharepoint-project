# Main.ps1

Import-Module PnP.PowerShell

$envVariables = Get-Content -Path "./.env" | Where-Object { $_ -match '=' } | ForEach-Object {
    $name, $value = $_ -split '=', 2
    [PSCustomObject]@{ Name = $name.Trim(); Value = $value.Trim() }
}

$envVariables | ForEach-Object {
    if ($_.Name -and $_.Value) {
        [Environment]::SetEnvironmentVariable($_.Name, $_.Value)
    }
}

$siteCount = Read-Host "Enter the number of sites to create"

Write-Host "Executing Connect-SharePoint.ps1 script..."
. .\Connect-SharePoint.ps1

Write-Host "Executing Create-Sites.ps1 script..."
. .\Create-Sites.ps1 -siteCount $siteCount

Write-Host "Main script executed successfully."
