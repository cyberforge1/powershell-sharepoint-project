# Import PnP.PowerShell module
Import-Module PnP.PowerShell

# Hardcoded environment variables
$SHAREPOINT_ADMIN_URL = "https://cyberforge000-admin.sharepoint.com"
$SHAREPOINT_SITE_URL = "https://cyberforge000.sharepoint.com"
$SHAREPOINT_USERNAME = "oliver@cyberforge000.onmicrosoft.com"
$SHAREPOINT_PASSWORD = '$i2odroY8K2s'  # Ensure the value is correct

# Site URL to download the template from
$SITE_URL = "$SHAREPOINT_SITE_URL/sites/SiteFifteen0001"
# Path to save the .xml file (project root directory)
$PROJECT_ROOT_PATH = "C:\Users\cyber\Desktop\powershell-sharepoint-project"
$OUTPUT_TEMPLATE_PATH = "$PROJECT_ROOT_PATH\EditedTemplate.xml"

# Convert password to secure string
$securePassword = ConvertTo-SecureString $SHAREPOINT_PASSWORD -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($SHAREPOINT_USERNAME, $securePassword)

# Connect to SharePoint
Write-Host "Connecting to SharePoint..."
Connect-PnPOnline -Url $SHAREPOINT_ADMIN_URL -Credentials $cred
Write-Host "Successfully connected to SharePoint Online."

# Connect to the specific site
Write-Host "Connecting to site: $SITE_URL"
Connect-PnPOnline -Url $SITE_URL -Credentials $cred
Write-Host "Successfully connected to site: $SITE_URL"

# Export the site template with all content
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

# Check if the template file exists and its size
if (Test-Path $OUTPUT_TEMPLATE_PATH) {
    $fileInfo = Get-Item $OUTPUT_TEMPLATE_PATH
    Write-Host "Template exported to: $OUTPUT_TEMPLATE_PATH"
    Write-Host "File size: $($fileInfo.Length) bytes"

    # Basic verification of the file content
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
