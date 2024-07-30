# MVP

The main tasks are getting familiar with PowerShell sharepoint commands, writing a PowerShell script, and researching how you'd go about automating template provisioning with Azure services.
Tenant users:

Task Summary:

You'll be working with PowerShell / PnP Templates and deploying them to vbtnd.sharepoint.com. You will need to install PowerShell if you don't have it.

Tips / Resources:

Templates are defined in template.xml files
You can log in with your candidate credentials directly at vbtnd.sharepoint.com
You can set up PowerShell / SharePoint here: https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
You can create new sits via PowerShell using: https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/new-sposite?view=sharepoint-ps
You can apply templates to sites using: https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/applying-pnp-templates
For number 5 you are going to need to create a PowerShell script that creates 10,000 sites using your template. It should be a fairly straight forward googling task.
For some inspiration on task 6/7, google around "deploy sharepoint pnp template from azure", you'll get a lot of articles like this: https://www.sharepointfire.com/2018/04/sharepoint-online-pnp-site-provisioning-using-flow-and-azure-function/




1) Fork https://github.com/SharePoint/sp-dev-provisioning-templates

2) Create a new site and apply template (https://github.com/SharePoint/sp-dev-provisioning-templates/tree/master/tenant/contosoworks) to it. This can be on dev tenant vbtnd.onmicrosoft.com or your own. Provide link as response

3) Modify template to include the new years for Global Country Holiday, include ANZ public holidays.

4) Capture the site and push to forked RP with PR and merge

5) Create script to apply this template to 10000 sites. Remember title, description and URL will be different for each of these sites. Add this script as README.md to GitHub repo

6) Explain your approach for applying this template on demand via Azure

7) Explain Your approach for integrating a solution in step 6 into other systems