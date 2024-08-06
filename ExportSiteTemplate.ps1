# ExportSiteTemplate.ps1

Import-Module PnP.PowerShell

$SHAREPOINT_ADMIN_URL = "XXXXXXXXXXXXXXXX"
$SHAREPOINT_SITE_URL = "XXXXXXXXXXXXXXXX"
$SHAREPOINT_USERNAME = "XXXXXXXXXXXXXXXX"
$SHAREPOINT_PASSWORD = 'XXXXXXXXXXXXXXXX'

$SITE_URL = "$SHAREPOINT_SITE_URL/sites/XXXXXXXXXXXXXXXX"
$PROJECT_ROOT_PATH = "XXXXXXXXXXXXXXXX"
$OUTPUT_TEMPLATE_PATH = "$PROJECT_ROOT_PATH\EditedTemplate.xml"

$securePassword = ConvertTo-SecureString $SHAREPOINT_PASSWORD -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($SHAREPOINT_USERNAME, $securePassword)

Write-Host "Connecting to SharePoint..."
Connect-PnPOnline -Url $SHAREPOINT_ADMIN_URL -Credentials $cred
Write-Host "Successfully connected to SharePoint Online."

Write-Host "Connecting to site: $SITE_URL"
Connect-PnPOnline -Url $SITE_URL -Credentials $cred
Write-Host "Successfully connected to site: $SITE_URL"

Write-Host "Exporting site template..."
$exportParams = @{
    Out         = $OUTPUT_TEMPLATE_PATH
    PersistBrandingFiles = $true
    IncludeAllTermGroups = $true
    IncludeSiteCollectionTermGroup = $true
    IncludeSearchConfiguration = $true
    Handlers    = "All"
}
Get-PnPSiteTemplate @exportParams

if (Test-Path $OUTPUT_TEMPLATE_PATH) {
    $fileInfo = Get-Item $OUTPUT_TEMPLATE_PATH
    Write-Host "Template exported to: $OUTPUT_TEMPLATE_PATH"
    Write-Host "File size: $($fileInfo.Length) bytes"

    $fileContent = Get-Content $OUTPUT_TEMPLATE_PATH
    if ($fileContent -like "*<pnp:ProvisioningTemplate*") {
        Write-Host "Template file appears to be valid."
    } else {
        Write-Host "Warning: Template file may be incomplete or corrupted."
    }
} else {
    Write-Host "Error: Template file was not created."
}

Write-Host "Script executed successfully."
