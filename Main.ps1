# Main.ps1

Import-Module PnP.PowerShell

# Hardcoded environment variables
$SHAREPOINT_ADMIN_URL = "https://cyberforge000-admin.sharepoint.com"
$SHAREPOINT_SITE_URL = "https://cyberforge000.sharepoint.com"
$OWNER_EMAIL = "oliver@cyberforge000.onmicrosoft.com"
$TEMPLATE_PATH = ".\contosoworks\source\template.xml"
$SHAREPOINT_USERNAME = "oliver@cyberforge000.onmicrosoft.com"
$SHAREPOINT_PASSWORD = '$i2odroY8K2s'  # Enclosed in single quotes to ensure the value is correct
$NEW_SITE_NAME = "SiteEight"
$SITE_ALIAS = "SiteEight"
$siteCount = 5  # Set the number of sites you want to create

# Debugging: Print out the hardcoded variables
Write-Host "DEBUG: Hardcoded variables:"
Write-Host "SHAREPOINT_ADMIN_URL = $SHAREPOINT_ADMIN_URL"
Write-Host "SHAREPOINT_SITE_URL = $SHAREPOINT_SITE_URL"
Write-Host "OWNER_EMAIL = $OWNER_EMAIL"
Write-Host "TEMPLATE_PATH = $TEMPLATE_PATH"
Write-Host "SHAREPOINT_USERNAME = $SHAREPOINT_USERNAME"
Write-Host "SHAREPOINT_PASSWORD = $SHAREPOINT_PASSWORD"
Write-Host "NEW_SITE_NAME = $NEW_SITE_NAME"
Write-Host "SITE_ALIAS = $SITE_ALIAS"
Write-Host "siteCount = $siteCount"

# Ensure critical environment variables are set
if (-not $SHAREPOINT_SITE_URL) {
    Write-Error "SHAREPOINT_SITE_URL is not set."
    exit 1
}

if (-not $NEW_SITE_NAME) {
    Write-Error "NEW_SITE_NAME is not set."
    exit 1
}

if (-not $TEMPLATE_PATH) {
    Write-Error "TEMPLATE_PATH is not set."
    exit 1
}

if (-not $SHAREPOINT_PASSWORD) {
    Write-Error "SHAREPOINT_PASSWORD is not set."
    exit 1
}

if (-not $SHAREPOINT_USERNAME) {
    Write-Error "SHAREPOINT_USERNAME is not set."
    exit 1
}

if (-not $OWNER_EMAIL) {
    Write-Error "OWNER_EMAIL is not set."
    exit 1
}

# Convert password to secure string
$securePassword = ConvertTo-SecureString $SHAREPOINT_PASSWORD -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($SHAREPOINT_USERNAME, $securePassword)

# Connect to SharePoint
Write-Host "Connecting to SharePoint..."
Connect-PnPOnline -Url $SHAREPOINT_ADMIN_URL -Credentials $cred
Write-Host "Successfully connected to SharePoint Online."

# Resolve the template path to an absolute path
$resolvedTemplatePath = Resolve-Path $TEMPLATE_PATH

# Function to convert DateTime format
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

# Function to test and convert DateTime fields
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

# Function to invoke the template
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

# Function to create new sites
function New-Sites {
    param (
        [int]$siteCount,
        [string]$sitePrefix,
        [string]$templatePath
    )
    
    Write-Host "Starting site creation process with $siteCount sites..."
    for ($i = 1; $i -le $siteCount; $i++) {
        $siteNumber = "{0:D4}" -f $i
        $siteUrl = "$($SHAREPOINT_SITE_URL)/sites/$($sitePrefix)$siteNumber"
        $siteTitle = "$sitePrefix $siteNumber"
        $siteDescription = "Site $sitePrefix number $siteNumber"
        
        # Debug: Print site details
        Write-Host "Creating site: $siteUrl with title: $siteTitle and description: $siteDescription"
        Write-Host "SHAREPOINT_SITE_URL: $SHAREPOINT_SITE_URL"
        Write-Host "sitePrefix: $sitePrefix"
        Write-Host "siteNumber: $siteNumber"
        Write-Host "siteUrl: $siteUrl"

        if ($siteUrl -and $siteUrl -ne "") {
            try {
                New-PnPSite -Type CommunicationSite -Url $siteUrl -Owner $OWNER_EMAIL -Title $siteTitle -Description $siteDescription
                Write-Host "Created site: $siteUrl"
                
                # Introduce a delay before applying the template
                Write-Host "Waiting for 60 seconds before applying the template..."
                Start-Sleep -Seconds 60
                
                Invoke-Template -siteUrl $siteUrl -templatePath $templatePath
            } catch {
                Write-Error "Error creating site ${siteTitle}: $_"
            }
        } else {
            Write-Error "The site URL is empty or not properly set: $siteUrl"
        }
    }
}

# Execute site creation
New-Sites -siteCount $siteCount -sitePrefix $NEW_SITE_NAME -templatePath $resolvedTemplatePath

Write-Host "Script executed successfully."
