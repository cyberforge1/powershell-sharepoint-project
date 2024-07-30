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

# Function to create sites
function New-Sites {
    param (
        [int]$siteCount,
        [string]$sitePrefix,
        [string]$templatePath
    )
    
    for ($i = 1; $i -le $siteCount; $i++) {
        $siteNumber = "{0:D4}" -f $i
        $siteUrl = "$env:SHAREPOINT_SITE_URL/sites/$sitePrefix$siteNumber"
        $siteTitle = "$sitePrefix $siteNumber"
        $siteDescription = "Site $sitePrefix number $siteNumber"
        
        try {
            Write-Host "Creating site: $siteUrl"
            New-PnPSite -Type CommunicationSite -Url $siteUrl -Owner $env:OWNER_EMAIL -Title $siteTitle -Description $siteDescription
            
            Write-Host "Created site: $siteUrl"
            Invoke-Template -siteUrl $siteUrl -templatePath $templatePath
        } catch {
            Write-Error "Error creating site ${siteTitle}: $_"
        }
    }
}

# Function to apply template
function Invoke-Template {
    param (
        [string]$siteUrl,
        [string]$templatePath
    )
    
    try {
        Connect-PnPOnline -Url $siteUrl -Credentials $cred
        Write-Host "Applying template to site: $siteUrl"
        Invoke-PnPSiteTemplate -Path $templatePath
        Write-Host "Template applied to site: $siteUrl"
    } catch {
        Write-Error "Error applying template to site ${siteUrl}: $_"
    }
}

# Main script execution
param (
    [int]$siteCount = 5
)

New-Sites -siteCount $siteCount -sitePrefix $env:NEW_SITE_NAME -templatePath $env:TEMPLATE_PATH

Write-Host "Site creation script executed successfully."
