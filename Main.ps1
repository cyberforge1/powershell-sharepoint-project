# Main.ps1

Import-Module PnP.PowerShell

$SHAREPOINT_ADMIN_URL = "XXXXXXXXXXXXXXXX"
$SHAREPOINT_SITE_URL = "XXXXXXXXXXXXXXXX"
$OWNER_EMAIL = "XXXXXXXXXXXXXXXX"
$TEMPLATE_PATH = ".\EditedTemplate.xml"
$SHAREPOINT_USERNAME = "XXXXXXXXXXXXXXXX"
$SHAREPOINT_PASSWORD = 'XXXXXXXXXXXXXXXX'
$NEW_SITE_NAME = "XXXXXXXXXXXXXXXX"
$SITE_ALIAS = "XXXXXXXXXXXXXXXX"
$siteCount = 5 


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
Write-Host "EVENTS_FILE_PATH = $EVENTS_FILE_PATH"
Write-Host "EVENTS_FILE_PATH = $EVENTS_FILE_PATH"

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

$securePassword = ConvertTo-SecureString $SHAREPOINT_PASSWORD -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($SHAREPOINT_USERNAME, $securePassword)

Write-Host "Connecting to SharePoint..."
Connect-PnPOnline -Url $SHAREPOINT_ADMIN_URL -Credentials $cred
Write-Host "Successfully connected to SharePoint Online."

$resolvedTemplatePath = Resolve-Path $TEMPLATE_PATH
Write-Host "Resolved template path: $resolvedTemplatePath"

function Convert-DateTimeFormatInTemplate {
    param (
        [string]$templatePath
    )

    [xml]$xmlContent = Get-Content -Path $templatePath

    $dateTimePattern = '(\d{1,2})/(\d{1,2})/(\d{4})\s(\d{1,2}):(\d{2})\s([APMapm]{2})'

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

    $modifiedTemplatePath = [System.IO.Path]::ChangeExtension($templatePath, "modified.xml")
    $xmlContent.Save($modifiedTemplatePath)

    return $modifiedTemplatePath
}

function Invoke-Template {
    param (
        [string]$siteUrl,
        [string]$templatePath
    )
    
    try {
        Write-Host "Connecting to site: $siteUrl"
        Connect-PnPOnline -Url $siteUrl -Credentials $cred
        Write-Host "Applying template to site: $siteUrl"
        Invoke-PnPSiteTemplate -Path $templatePath
        Write-Host "Template applied to site: $siteUrl"
    } catch {
        Write-Error "Error applying template to site ${siteUrl}: $_"
    }
}

function Add-CalendarEvents {
    param (
        [string]$siteUrl,
        [string]$eventsFilePath
    )
    
    Write-Host "Adding calendar events to site: $siteUrl"
    
    $events = Import-Csv -Path $eventsFilePath
    
    foreach ($event in $events) {
        $eventTitle = $event.Title
        $startTime = [datetime]::ParseExact($event.'Start Time', 'M/d/yyyy H:mm', $null)
        $endTime = [datetime]::ParseExact($event.'End Time', 'M/d/yyyy H:mm', $null)
        $recurrence = [bool]$event.Recurrence
        
        try {
            Add-PnPListItem -List "Calendar" -Values @{
                "Title" = $eventTitle
                "EventDate" = $startTime
                "EndDate" = $endTime
                "fAllDayEvent" = $true
                "fRecurrence" = $recurrence
            }
            Write-Host "Added event: $eventTitle"
        } catch {
            Write-Error "Error adding event ${eventTitle}: $_"
        }
    }
}

function New-Sites {
    param (
        [int]$siteCount,
        [string]$sitePrefix,
        [string]$templatePath,
        [string]$eventsFilePath
        [string]$templatePath,
        [string]$eventsFilePath
    )
    
    Write-Host "Starting site creation process with $siteCount sites..."
    for ($i = 1; $i -le $siteCount; $i++) {
        $siteNumber = "{0:D4}" -f $i
        $siteUrl = "$SHAREPOINT_SITE_URL/sites/$sitePrefix$siteNumber"
        $siteTitle = "$sitePrefix $siteNumber"
        $siteDescription = "Site $sitePrefix number $siteNumber"
        
        Write-Host "DEBUG: Creating site with the following details:"
        Write-Host "siteUrl: $siteUrl"
        Write-Host "siteTitle: $siteTitle"
        Write-Host "siteDescription: $siteDescription"

        if ($siteUrl -and $siteUrl -ne "") {
            try {
                Write-Host "Creating site: $siteUrl"
                New-PnPSite -Type CommunicationSite -Url $siteUrl -Owner $OWNER_EMAIL -Title $siteTitle -Description $siteDescription
                Write-Host "Created site: $siteUrl"
                
                Write-Host "Waiting for 45 seconds before applying the template..."
                Start-Sleep -Seconds 45
                
                $modifiedTemplatePath = Convert-DateTimeFormatInTemplate -templatePath $templatePath
                Write-Host "Using modified template path: $modifiedTemplatePath"
                
                Invoke-Template -siteUrl $siteUrl -templatePath $modifiedTemplatePath
                
                Add-CalendarEvents -siteUrl $siteUrl -eventsFilePath $eventsFilePath
            } catch {
                Write-Error "Error creating site ${siteTitle}: $_"
            }
        } else {
            Write-Error "The site URL is empty or not properly set: $siteUrl"
        }
    }
}

New-Sites -siteCount $siteCount -sitePrefix $NEW_SITE_NAME -templatePath $resolvedTemplatePath -eventsFilePath $EVENTS_FILE_PATH

Write-Host "Script executed successfully."
