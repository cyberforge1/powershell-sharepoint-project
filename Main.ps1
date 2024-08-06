# Main.ps1

Import-Module PnP.PowerShell

# Hardcoded environment variables
$SHAREPOINT_ADMIN_URL = "https://cyberforge000-admin.sharepoint.com"
$SHAREPOINT_SITE_URL = "https://cyberforge000.sharepoint.com"
$OWNER_EMAIL = "oliver@cyberforge000.onmicrosoft.com"
$TEMPLATE_PATH = ".\contosoworks\source\template.xml"
$SHAREPOINT_USERNAME = "oliver@cyberforge000.onmicrosoft.com"
$SHAREPOINT_PASSWORD = '$i2odroY8K2s'  # Enclosed in single quotes to ensure the value is correct
$NEW_SITE_NAME = "SiteFifteen"
$SITE_ALIAS = "SiteFifteen"
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
Write-Host "Resolved template path: $resolvedTemplatePath"

# Function to convert DateTime format within the XML template
function Convert-DateTimeFormatInTemplate {
    param (
        [string]$templatePath
    )

    # Load the XML template
    [xml]$xmlContent = Get-Content -Path $templatePath

    # Define the DateTime format pattern
    $dateTimePattern = '(\d{1,2})/(\d{1,2})/(\d{4})\s(\d{1,2}):(\d{2})\s([APMapm]{2})'

    # Iterate through each node and convert DateTime format
    $xmlContent.SelectNodes("//*[text()]") | ForEach-Object {
        if ($_.InnerText -match $dateTimePattern) {
            $dateTimeString = $_.InnerText
            try {
                $parsedDateTime = [datetime]::ParseExact($dateTimeString, 'M/d/yyyy h:mm tt', $null)
                $_.InnerText = $parsedDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
            } catch {
                Write-Error "Error parsing DateTime: $dateTimeString"
            }
        }
    }

    # Save the modified XML template to a new file
    $modifiedTemplatePath = [System.IO.Path]::ChangeExtension($templatePath, "modified.xml")
    $xmlContent.Save($modifiedTemplatePath)

    return $modifiedTemplatePath
}

# Function to invoke the template
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
        $siteUrl = "$($SHAREPOINT_SITE_URL)/sites/$($sitePrefix)$($siteNumber)"
        $siteTitle = "$sitePrefix $siteNumber"
        $siteDescription = "Site $sitePrefix number $siteNumber"
        
        # Debug: Print site details
        Write-Host "DEBUG: Creating site with the following details:"
        Write-Host "siteUrl: $siteUrl"
        Write-Host "siteTitle: $siteTitle"
        Write-Host "siteDescription: $siteDescription"

        if ($siteUrl -and $siteUrl -ne "") {
            try {
                New-PnPSite -Type CommunicationSite -Url $siteUrl -Owner $OWNER_EMAIL -Title $siteTitle -Description $siteDescription
                Write-Host "Created site: $siteUrl"
                
                # Introduce a delay before applying the template
                Write-Host "Waiting for 30 seconds before applying the template..."
                Start-Sleep -Seconds 30
                
                # Convert DateTime format in the template
                $modifiedTemplatePath = Convert-DateTimeFormatInTemplate -templatePath $templatePath
                Write-Host "Using modified template path: $modifiedTemplatePath"
                
                Invoke-Template -siteUrl $siteUrl -templatePath $modifiedTemplatePath
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
