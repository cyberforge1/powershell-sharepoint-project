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

# Convert plain text password to SecureString
$securePassword = ConvertTo-SecureString $env:SHAREPOINT_PASSWORD -AsPlainText -Force

# Create PSCredential object
$cred = New-Object System.Management.Automation.PSCredential ($env:SHAREPOINT_USERNAME, $securePassword)

# Connect to SharePoint Online
Connect-PnPOnline -Url $env:SHAREPOINT_ADMIN_URL -Credentials $cred

Write-Host "Successfully connected to SharePoint Online."
