# Create-Sites.ps1

param (
    [int]$siteCount
)

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

$securePassword = ConvertTo-SecureString $env:SHAREPOINT_PASSWORD -AsPlainText -Force

$cred = New-Object System.Management.Automation.PSCredential ($env:SHAREPOINT_USERNAME, $securePassword)

function Test-DateTimeFields {
    param (
        [string]$templatePath
    )
    
    [xml]$xmlTemplate = Get-Content -Path $templatePath
    
    foreach ($field in $xmlTemplate.SelectNodes("//Field[@Type='DateTime']")) {
        $dateTimeValue = $field.InnerText.Trim()
        if ($dateTimeValue) {
            try {
                [datetime]::Parse($dateTimeValue) | Out-Null
            } catch {
                Write-Error "Invalid DateTime format: $dateTimeValue in field $($field.Name)"
            }
        } else {
            Write-Error "Empty DateTime field found: $($field.Name)"
        }
    }
}

function Invoke-Template {
    param (
        [string]$siteUrl,
        [string]$templatePath
    )
    
    Test-DateTimeFields -templatePath $templatePath
    
    try {
        Connect-PnPOnline -Url $siteUrl -Credentials $cred
        Write-Host "Applying template to site: $siteUrl"
        Invoke-PnPSiteTemplate -Path $templatePath
        Write-Host "Template applied to site: $siteUrl"
    } catch {
        Write-Error "Error applying template to site ${siteUrl}: $_"
        if ($_.Exception -match "String '(.*)' was not recognized as a valid DateTime") {
            Write-Host "Invalid DateTime format found: $($matches[1])"
        }
    }
}

function New-Sites {
    param (
        [int]$siteCount,
        [string]$sitePrefix,
        [string]$templatePath
    )
    
    Write-Host "Starting site creation process with $siteCount sites..."
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

New-Sites -siteCount $siteCount -sitePrefix $env:NEW_SITE_NAME -templatePath $env:TEMPLATE_PATH

Write-Host "Site creation script executed successfully."
