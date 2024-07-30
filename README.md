# SharePoint Site Creation with PowerShell

## Description
This project uses PowerShell scripting to automate the creation and application of templates to multiple SharePoint sites.

## Tech Stack
- PowerShell 7.4.4
- Microsoft SharePoint
- PnP.PowerShell module

## PowerShell Scripting

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

```

```powershell
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

## Notes

### Application of Azure Services

Azure Automation and Functions can be used to run PowerShell scripts for creating SharePoint sites in the cloud. For integration with this project Automation can manage the execution of the `Connect-SharePoint.ps1` and `Create-Sites.ps1` scripts and Functions can handle HTTP triggers to run the automatic creation of sites. Services like Azure Key Vault can also be used to bridge services like these to allow shared access and also secure storage for sensitive project information.

### General Application

The three services - Azure Automation, Functions and Key Vault - work cohesively together to provide automation processes with a trigger, that can be used for many applications. A focus on outsourcing demanding compute tasks to virtual machines in the cloud is incredibly powerful and useful, particularly for tasks that benefit from scaling of resources. Also the serverless nature of Azure Functions allows for cost efficient solutions with having to manage infrastrucutre directly. 


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
