# SharePoint Site Creation with PowerShell

## Description
This project uses PowerShell scripting to automate the creation and application of templates to multiple SharePoint sites.

## Tech Stack
- PowerShell 7
- Microsoft SharePoint
- PnP.PowerShell module

## PowerShell Scripts

### Main.ps1

```powershell
$siteCount = Read-Host "Enter the number of sites to create"

Write-Host "Executing Connect-SharePoint.ps1 script..."
. .\Connect-SharePoint.ps1

Write-Host "Executing Create-Sites.ps1 script..."
. .\Create-Sites.ps1 -siteCount $siteCount

Write-Host "Main script executed successfully."
```

### Connect-SharePoint.ps1

```powershell
$securePassword = ConvertTo-SecureString $env:SHAREPOINT_PASSWORD -AsPlainText -Force

$cred = New-Object System.Management.Automation.PSCredential ($env:SHAREPOINT_USERNAME, $securePassword)

Connect-PnPOnline -Url $env:SHAREPOINT_ADMIN_URL -Credentials $cred
```

### Create-Sites.ps1

```powershell
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
```

## MVP
1) Fork https://github.com/SharePoint/sp-dev-provisioning-templates
2) Create a new site and apply template (https://github.com/SharePoint/sp-dev-provisioning-templates/tree/master/tenant/contosoworks) to it. This can be on dev tenant vbtnd.onmicrosoft.com or your own. Provide link as response
3) Modify template to include the new years for Global Country Holiday, include ANZ public holidays.
4) Capture the site and push to forked RP with PR and merge
5) Create script to apply this template to 10000 sites. Remember title, description and URL will be different for each of these sites.
6) Explain your approach for applying this template on demand via Azure
7) Explain Your approach for integrating a solution in step 6 into other systems

## Structure
- The script `Connect-SharePoint.ps1` handles connecting to SharePoint Online using stored credentials.
- The script `Create-Sites.ps1` creates the sites and applies the template.
- Adjust the `.env` file with your specific environment details.

## Usage
1. Run the `Main.ps1` script:
    ```sh
    pwsh .\Main.ps1
    ```

2. Enter the number of sites you want to create when prompted.

## Setup
1. Install the required PowerShell module:
    ```sh
    Install-Module PnP.PowerShell -Force -AllowClobber
    ```

2. Create a `.env` file in the same directory as the scripts with the following content:
    ```env
    SHAREPOINT_ADMIN_URL=your_admin_url
    SHAREPOINT_SITE_URL=your_sharepoint_site_url
    OWNER_EMAIL=your_owner_email
    TEMPLATE_PATH=your_template_path
    SHAREPOINT_USERNAME=your_sharepoint_username
    SHAREPOINT_PASSWORD=your_sharepoint_password
    NEW_SITE_NAME=your_new_site_name
    SITE_ALIAS=your_site_alias
    ```
