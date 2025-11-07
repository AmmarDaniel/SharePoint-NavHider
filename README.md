# SharePoint-NavHider
SharePoint Application Customizer to hide the navigation bar only on the home page and show it on other pages.

## How to Use
1. Use Powershell 7 or above
2. Install PnP:
> Install-Module PnP.PowerShell -Scope CurrentUser -Force
> Import-Module PnP.PowerShell
3. > Connect-PnPOnline `
 -Url "https://contoso.sharepoint.com/sites/Site" `
 -ClientId "#" `
 -Tenant "#" `
 -Interactive

 ### Hide navigation bar (add action)
 > Add-PnPCustomAction `
 -Name "NavHider" `
 -Title "NavHider" `
 -Location "ClientSideExtension.ApplicationCustomizer" `
 -ClientSideComponentId "#" `
 -ClientSideComponentProperties '{"hideOnHomeOnly": true, "extraHidePaths": []}' `
 -Scope Web

### Show navgation bar (remove action)
> Get-PnPCustomAction | Where-Object {
 $_.Location -eq 'ClientSideExtension.ApplicationCustomizer' -and $_.Title -eq 'NavHider'
} | Remove-PnPCustomAction
