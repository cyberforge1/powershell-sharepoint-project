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

# Export the site template
Write-Host "Exporting site template..."
Get-PnPSiteTemplate -Out $OUTPUT_TEMPLATE_PATH
Write-Host "Template exported to: $OUTPUT_TEMPLATE_PATH"

Write-Host "Script executed successfully."


