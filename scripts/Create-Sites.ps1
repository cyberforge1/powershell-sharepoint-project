# Create-Sites.ps1

function Initialize-EnvironmentVariables {
    $envVariables = Get-Content -Path "./.env" | Where-Object { $_ -match '=' } | ForEach-Object {
        $name, $value = $_ -split '=', 2
        [PSCustomObject]@{ Name = $name.Trim(); Value = $value.Trim() }
    }

    $envVariables | ForEach-Object {
        if ($_.Name -and $_.Value) {
            [Environment]::SetEnvironmentVariable($_.Name, $_.Value)
        }
    }
}

Initialize-EnvironmentVariables

# Resolve the template path to an absolute path
$resolvedTemplatePath = Resolve-Path $env:TEMPLATE_PATH

# Debug: Print environment variables and resolved template path
Write-Host "SHAREPOINT_SITE_URL: $($env:SHAREPOINT_SITE_URL)"
Write-Host "NEW_SITE_NAME: $($env:NEW_SITE_NAME)"
Write-Host "TEMPLATE_PATH: $resolvedTemplatePath"

$securePassword = ConvertTo-SecureString $env:SHAREPOINT_PASSWORD -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($env:SHAREPOINT_USERNAME, $securePassword)

function Convert-DateTimeFormat {
    param (
        [string]$dateTimeValue
    )
    try {
        $dateTime = [datetime]::ParseExact($dateTimeValue, 'MM/dd/yyyy', $null)
        return $dateTime.ToString('dd/MM/yyyy')
    } catch {
        Write-Error "Error converting DateTime format: $dateTimeValue"
        return $dateTimeValue
    }
}

function Test-And-Convert-DateTimeFields {
    param (
        [string]$templatePath
    )
    
    [xml]$xmlTemplate = Get-Content -Path $templatePath
    
    foreach ($field in $xmlTemplate.SelectNodes("//Field[@Type='DateTime']")) {
        $dateTimeValue = $field.InnerText.Trim()
        if ($dateTimeValue) {
            try {
                [datetime]::ParseExact($dateTimeValue, 'MM/dd/yyyy', $null) | Out-Null
                $convertedDateTimeValue = Convert-DateTimeFormat -dateTimeValue $dateTimeValue
                $field.InnerText = $convertedDateTimeValue
            } catch {
                Write-Error "Invalid DateTime format: $dateTimeValue in field $($field.Name)"
            }
        } else {
            Write-Error "Empty DateTime field found: $($field.Name)"
        }
    }

    $xmlTemplate.Save($templatePath)
}

function Invoke-Template {
    param (
        [string]$siteUrl,
        [string]$templatePath
    )
    
    Test-And-Convert-DateTimeFields -templatePath $templatePath
    
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
        $siteUrl = "$($env:SHAREPOINT_SITE_URL)/sites/$($sitePrefix)$siteNumber"
        $siteTitle = "$sitePrefix $siteNumber"
        $siteDescription = "Site $sitePrefix number $siteNumber"
        
        # Debug: Print site details
        Write-Host "Creating site: $siteUrl with title: $siteTitle and description: $siteDescription"
        Write-Host "SHAREPOINT_SITE_URL: $($env:SHAREPOINT_SITE_URL)"
        Write-Host "sitePrefix: $sitePrefix"
        Write-Host "siteNumber: $siteNumber"
        Write-Host "siteUrl: $siteUrl"

        if ($siteUrl -and $siteUrl -ne "") {
            try {
                New-PnPSite -Type CommunicationSite -Url $siteUrl -Owner $env:OWNER_EMAIL -Title $siteTitle -Description $siteDescription
                Write-Host "Created site: $siteUrl"
                
                # Introduce a delay before applying the template
                Write-Host "Waiting for 10 seconds before applying the template..."
                Start-Sleep -Seconds 10
                
                Invoke-Template -siteUrl $siteUrl -templatePath $templatePath
            } catch {
                Write-Error "Error creating site ${siteTitle}: $_"
            }
        } else {
            Write-Error "The site URL is empty or not properly set: $siteUrl"
        }
    }
}

# Print out the final values of environment variables to ensure they are set
Write-Host "Final Environment Variables:"
Write-Host "SHAREPOINT_SITE_URL: $($env:SHAREPOINT_SITE_URL)"
Write-Host "NEW_SITE_NAME: $($env:NEW_SITE_NAME)"
Write-Host "TEMPLATE_PATH: $resolvedTemplatePath"
Write-Host "OWNER_EMAIL: $($env:OWNER_EMAIL)"
Write-Host "SHAREPOINT_USERNAME: $($env:SHAREPOINT_USERNAME)"

New-Sites -siteCount $siteCount -sitePrefix $env:NEW_SITE_NAME -templatePath $resolvedTemplatePath

Write-Host "Site creation script executed successfully."
