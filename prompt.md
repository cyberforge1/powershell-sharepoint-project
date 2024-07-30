
Here are some requirements for an assignment using powershell 7 and sharepoint:

'''
- Create a new site and apply template (https://github.com/SharePoint/sp-dev-provisioning-templates/tree/master/tenant/contosoworks) to it. This can be on dev tenant vbtnd.onmicrosoft.com or your own. Provide link as response

- Create script to apply this template to 10000 sites. Remember title, description and URL will be different for each of these sites. Add this script as README.md to GitHub repo

'''



I want to write a series of powershell 7 scripts that are stored in different files. When run as a series they will take an input for a number of sites to create, connects to SharePoint online, and generates sequential site names, urls and descriptions for each site, and applies a predefined template to each site. It should apply the template using the Invoke-PnpSiteTemplate command. 

I have attached three scripts that are based around achieving this task for context, and want an optimised version that incorporates the best elements of each script to achieve the final result. It should use commonly used keywords and methods from these scripts, althought it should be written with it's own style and structure, and also not include any batching


'''
#5 Site Generation Script
The script is written within ./contosoworks/script.ps1

The script is written to be able to be used with any template, simply move the script into the folder above /source/ that contains your template.xml
It will take the /source/ folder and create a template.pnp out of it, then apply that to the sites The script first prompts your credentials and creates a credentials.xml file within the same folder as the script, and can then be used on subsequent executions

$templateFolderPath = "./source"
$templatePnpPath = "./template.pnp"
$totalSites = 4
$batchSize = 4
$sitePrefix = "minitest"
$credPath = "./credentials.xml"
$jobs = @()

function Get-Stored-Credential {
    param (
        [string]$credFilePath
    )
    if (-Not (Test-Path -Path $credFilePath)) {
        $cred = Get-Credential
        $spUrl = Read-Host "Enter SharePoint URL"
        $credHash = @{
            Credential = $cred
            SharePointUrl = $spUrl
        }
        $credHash | Export-Clixml -Path $credFilePath
    } else {
        Write-Host "Using stored credentials from $credFilePath"
    }
}

Get-Stored-Credential -credFilePath $credPath

$storedCreds = Import-Clixml -Path $credPath
$cred = $storedCreds.Credential
$spUrl = $storedCreds.SharePointUrl

Import-Module PnP.PowerShell
Import-Module ThreadJob
Connect-PnPOnline -Url $spUrl -Credentials $cred

Convert-PnPFolderToSiteTemplate -Folder $templateFolderPath -Out $templatePnpPath

for ($i = 0; $i -lt $totalSites; $i += $batchSize) {
    $end = [math]::Min($i + $batchSize, $totalSites)
    $jobs += Start-ThreadJob -ScriptBlock {
        param($start, $end, $sitePrefix, $templatePnpPath, $credPath)

        Import-Module PnP.PowerShell
        Import-Module ThreadJob
        $storedCreds = Import-Clixml -Path $credPath
        $cred = $storedCreds.Credential
        $adminEmail = $cred.UserName
        $spUrl = $storedCreds.SharePointUrl
        Connect-PnPOnline -Url $spUrl -Credentials $cred

        for ($j = $start; $j -lt $end; $j++) {
            try {
                $siteNumber = "{0:D4}" -f $j
                $siteUrl = "$spUrl/sites/$sitePrefix$siteNumber"
                $siteTitle = "$sitePrefix $siteNumber"
                $siteDescription = "Site $sitePrefix number $siteNumber"

                Write-Host "Creating site: $siteUrl"
                New-PnPSite -Type CommunicationSite -Url $siteUrl -Owner $adminEmail -Title $siteTitle -Description $siteDescription

                Write-Host "Created site: $siteUrl"
                Connect-PnPOnline -Url $siteUrl -Credentials $cred
                write-host "Connected"
                Invoke-PnPSiteTemplate -Path $templatePnpPath
                Write-Host "Created and applied template to site: $sitePrefix$siteNumber"
            } catch {
                Write-Error "Error processing ${siteTitle}: $_"
            }
        }
    } -ArgumentList $i, $end, $sitePrefix, $templatePnpPath, $credPath
}

$jobs | ForEach-Object { $_ | Receive-Job -Wait }

$jobs | ForEach-Object {
    if ($_.State -eq 'Completed') {
        Write-Host "Job $($_.Id) completed successfully."
    } else {
        Write-Error "Job $($_.Id) failed."
    }
    Remove-Job $_
}

Write-Host "Completed Script and Disconnected"
Script Requirements
Install-Module PnP.Powershell

Install-Module ThreadJob

Edit the script file to output the desired amount of sites and batch amounts

add credentials.xml to .gitignore
'''

'''
The script
The script prompts the user for how many sites they would like to provision. Assuming an input of 5 or greater, a while loop sends a batch job request for 5 new sites, decrements that from the input variable, and checks again. Once less than 5, it sends those as individual batch jobs.
Prompts input for total sites and then fires batch jobs.

$Sitecount = Read-Host "How many sites do you want to create?"

while ($Sitecount -ge 5) {
    Start-Job -FilePath .\maccap\sp-dev-provisioning-templates\scripts\innerbatch5.ps1
    $Sitecount -= 5
}

for ($j = 0; $j -lt $Sitecount; $j++)
{
Start-Job -FilePath .\maccap\sp-dev-provisioning-templates\scripts\innerbatch1.ps1
}
Batch job for leftovers. Could find a way to pass in the remainder as a single job instead of multiple single jobs.

Import-Module NameIT

$TenantUrl = "https://vbtnd-admin.sharepoint.com"
$TemplatePath = "maccap\sp-dev-provisioning-templates\tenant\contosoworks\source\template.xml"

Connect-PnPonline -url $TenantUrl -Interactive

    $Randadd = Invoke-Generate "#####?????"
    $SiteUrl = "https://vbtnd.sharepoint.com/sites/example$Randadd"
    $SiteTitle = "ExampleSite$Randadd"
    $SiteDescription = "This is the description for site $Randadd"

    New-PnpSite -Type CommunicationSite -Title "$SiteTitle" -Url $SiteUrl -Owner candidate04@vbtnd.onmicrosoft.com -description "$SiteDescription"
    Invoke-PnPTenantTemplate -Path $TemplatePath -Parameters @{"SiteTitle"="$SiteTitle";"SiteDescription"="$SiteDescription";"SiteUrl"="$SiteUrl"}


 "Completed batch job"
Batch job for 5

Import-Module NameIT

$TenantUrl = "https://vbtnd-admin.sharepoint.com"
$TemplatePath = "maccap\sp-dev-provisioning-templates\tenant\contosoworks\source\template.xml"

$count = 5

Connect-PnPonline -url $TenantUrl -Interactive

for ($i = 0; $i -lt $count; $i++)
{
    $Randadd = Invoke-Generate "#####?????"
    $SiteUrl = "https://vbtnd.sharepoint.com/sites/example$Randadd"
    $SiteTitle = "ExampleSite$Randadd"
    $SiteDescription = "This is the description for site $Randadd"

    New-PnpSite -Type CommunicationSite -Title "$SiteTitle" -Url $SiteUrl -Owner candidate04@vbtnd.onmicrosoft.com -description "$SiteDescription"
    Invoke-PnPTenantTemplate -Path $TemplatePath -Parameters @{"SiteTitle"="$SiteTitle";"SiteDescription"="$SiteDescription";"SiteUrl"="$SiteUrl"}
 }

 "Completed batch job"

 Import-Module NameIT

$TenantUrl = "https://vbtnd-admin.sharepoint.com"
$TemplatePath = "maccap\sp-dev-provisioning-templates\tenant\contosoworks\source\template.xml"

Connect-PnPonline -url $TenantUrl -Interactive

    $Randadd = Invoke-Generate "#####?????"
    $SiteUrl = "https://vbtnd.sharepoint.com/sites/example$Randadd"
    $SiteTitle = "ExampleSite$Randadd"
    $SiteDescription = "This is the description for site $Randadd"

    New-PnpSite -Type CommunicationSite -Title "$SiteTitle" -Url $SiteUrl -Owner candidate04@vbtnd.onmicrosoft.com -description "$SiteDescription"
    Invoke-PnPTenantTemplate -Path $TemplatePath -Parameters @{"SiteTitle"="$SiteTitle";"SiteDescription"="$SiteDescription";"SiteUrl"="$SiteUrl"}


 "Completed batch job"

 

$Sitecount = Read-Host "How many sites do you want to create?"

while ($Sitecount -ge 5) {
    Start-Job -FilePath .\maccap\sp-dev-provisioning-templates\scripts\innerbatch5.ps1
    $Sitecount -= 5
}

for ($j = 0; $j -lt $Sitecount; $j++)
{
Start-Job -FilePath .\maccap\sp-dev-provisioning-templates\scripts\innerbatch1.ps1
}
'''

'''

Import-Module PnP.PowerShell

$baseUrl = Read-Host "Enter base URL for SharePoint (e.g. https://tenant.sharepoint.com/sites/)"
$baseTitle = Read-Host "Enter Site Name"
$baseDescription = "Description for "
$siteNumber = [int] (Read-Host "Enter # of sites to be generated")
$siteOwner = Read-Host "Enter site owner user"
$adminSiteUrl = Read-Host "Enter Admin SharePoint URL"
$templatePath = Read-Host "Enter Path to template"
$credential = Get-Credential
Connect-PnPOnline -Url $adminSiteUrl -Credentials $credential

for ($i = 1; $i -le $siteNumber; $i++) {
    $siteTitle = "$baseTitle$i"
    $siteUrl = "$baseUrl$siteTitle"
    $siteDescription = "$baseDescription$siteTitle"

    try {

        Write-Host "Creating site: $siteTitle at URL $siteUrl"
        New-PnPSite -Type CommunicationSite -Title $siteTitle -Url $siteUrl -Lcid 1033 -Description $siteDescription -Owner $siteOwner
        Write-Host "Site creation in progress for: $siteUrl"
    }
    catch {
        Write-Host "Failed to create site: $siteUrl. Error: $_"
        continue
    }
    Start-Sleep -Seconds 60
    Connect-PnPOnline -Url $siteUrl -Interactive

    try {

        Write-Host "Applying template to site: $siteUrl"
        Invoke-PnPSiteTemplate -Path $templatePath
        Write-Host "Template applied to site: $siteUrl"
    }
    catch {
        Write-Host "Failed to apply template to site: $siteUrl. Error: $_"
    }
    Start-Sleep -Seconds 10

}

Disconnect-PnPOnline


Approach
To make the script as modular as possible requiring no changes to the script itself when used on different tenants.
Features
Automates the creation and application of templates to sites
Is fully modular and allows the user to input any credential and host site values upon the script running.
'''
'''






The new script should include the variables from the original project code, as is found below:

'''

# .env

SHAREPOINT_ADMIN_URL=https://cyberforge000-admin.sharepoint.com
SHAREPOINT_SITE_URL=https://cyberforge000.sharepoint.com
OWNER_EMAIL=oliver@cyberforge000.onmicrosoft.com
TEMPLATE_PATH=.\template.xml
SHAREPOINT_USERNAME=oliver@cyberforge000.onmicrosoft.com
SHAREPOINT_PASSWORD=$i2odroY8K2s
NEW_SITE_NAME=NewTestSite
SITE_ALIAS=NewTestSite




'''

so that it is completely functional. Please also make the scripts modular and provide comments above each section, explaining what it does. It should be concise and efficient, and work on the first run

