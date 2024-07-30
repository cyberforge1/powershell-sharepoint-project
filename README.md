# SharePoint Site Creation Scripts

## Description
This series of PowerShell scripts automates the creation of multiple SharePoint sites and applies a predefined template to each site.

## Prerequisites
- PowerShell 7
- PnP.PowerShell module

## Setup
1. Install the required PowerShell module:
    ```sh
    Install-Module PnP.PowerShell -Force -AllowClobber
    ```

2. Create a `.env` file in the same directory as the scripts with the following content:
    ```env
    SHAREPOINT_ADMIN_URL=XXXXXXXXXXX
    SHAREPOINT_SITE_URL=XXXXXXXXXXX
    OWNER_EMAIL=XXXXXXXXXXX
    TEMPLATE_PATH=XXXXXXXXXXX
    SHAREPOINT_USERNAME=XXXXXXXXXXX
    SHAREPOINT_PASSWORD=XXXXXXXXXXX
    NEW_SITE_NAME=XXXXXXXXXXX
    SITE_ALIAS=XXXXXXXXXXX
    ```

## Usage
1. Run the `Main.ps1` script:
    ```sh
    pwsh .\Main.ps1
    ```

2. Enter the number of sites you want to create when prompted.

## Notes
- The script `Connect-SharePoint.ps1` handles connecting to SharePoint Online using stored credentials.
- The script `Create-Sites.ps1` creates the sites and applies the template.
- Adjust the `.env` file with your specific environment details.
