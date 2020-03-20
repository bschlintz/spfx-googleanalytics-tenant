## Google Analytics for SharePoint

Deploy Google Analytics to all modern sites in your SharePoint Online tenant using a SharePoint Framework extension.

## Setup
### Pre-requisites
- App Catalog: Ensure the [App Catalog](https://docs.microsoft.com/en-us/sharepoint/use-app-catalog) is setup in your SharePoint Online tenant

### Installation
1. Download the SPFx package [googleanalytics-tenant.sppkg](https://github.com/bschlintz/spfx-googleanalytics-tenant/blob/master/sharepoint/solution/googleanalytics-tenant.sppkg) file from Github (or clone the repo and build the package yourself)
2. Upload sppkg file to the 'Apps for SharePoint' library in your Tenant App Catalog
3. Click the 'Make this solution available to all sites in your organization' checkbox and then click Deploy

### Configuration
This solution is deployed using [Tenant Wide Extensions](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions). You can modify the JSON properties configuration via the item in the Tenant Wide Extensions list. The available properties are below.

| Extension Property Name       | Default Value | Description |
| ------------------- | ------------- | ----- |
| googleTrackingId | `""` | Required.<br/>Google Analytics Example: UA-111111111-1<br/>Google Tag Manager Example: GTM-XXXX |
| useGoogleTagManager | `false` | Optional. Use [Google Tag Manager](https://developers.google.com/tag-manager/quickstart) instead of [Google Analytics](https://developers.google.com/analytics/devguides/collection/gtagjs). |
| propertyMappings | `undefined` | Optional (Google Tag Manager only).<br/> Collection of user property mappings from the SharePoint User Profile service to the Google Tag Manager dataLayer. |
| disableCache | `false` | Optional (Google Tag Manager only).<br/>Disable client-side caching of SharePoint User Profile when using Google Tag Manager with property mappings. |
| cacheLifetimeMins | `720`<br/>(12 hours) | Optional (Google Tag Manager only).<br/>Number of minutes the SharePoint User Profile should remain cached in browser localStorage before re-fetching. |

#### Property Mappings
Specify a collection of mappings from SharePoint User Profile Properties to Google Tag Manager dataLayer properties. 

| Mapped Property Name | Default Value | Description |
| ------------------- | ------------- | ----- |
| spoUserPropertyName | `undefined` | Required. Internal name of SharePoint User Profile property. |
| gtmPropertyName | `undefined` | Required. Name of the property to map the User Profile Value into when sent to Google Tag Manager. |
| encryptSHA256 | `false` | Optional. Encrypt with SHA256 and send masked hash value instead of plain text value. |

Example tenant wide extension configuration with property mappings below:

```json
{
  "googleTrackingId": "GTM-XXXX",
  "useGoogleTagManager": true,
  "propertyMappings": [
    { "spoUserPropertyName": "Title", "gtmPropertyName": "jobTitle" },
    { "spoUserPropertyName": "contosoEmployeeID", "gtmPropertyName": "employeeID", "encryptSHA256": true },
    { "spoUserPropertyName": "contosoBusGroup", "gtmPropertyName": "businessGroup" },
    { "spoUserPropertyName": "contosoBusSubGroup", "gtmPropertyName": "businessSubGroup" },
    { "spoUserPropertyName": "contosoManagerInd", "gtmPropertyName": "manager" },
    { "spoUserPropertyName": "contosoExemptStatus", "gtmPropertyName": "expemptStat" },
    { "spoUserPropertyName": "contosoFullTimeInd", "gtmPropertyName": "fullTime" },
    { "spoUserPropertyName": "contosoStartDate", "gtmPropertyName": "startDate" },
    { "spoUserPropertyName": "contosoCity", "gtmPropertyName": "city" },
    { "spoUserPropertyName": "contosoState", "gtmPropertyName": "state" }
  ]
}
```

![Tenant Wide Extension List Item](./docs/TenantWideExtensionItem.png)

## Updates
Follow the same steps as installation. Overwrite the existing package in the 'Apps for SharePoint' library when uploading the new package. 

> __Tip #1__: Be sure to check-in the sppkg file after the deployment if it is left checked-out.

> __Tip #2__: Ensure there aren't duplicate entries in the Tenant Wide Extensions list after deploying an update. Delete any duplicates if there are any.

## Removal

### Uninstall from Tenant
1. Delete the entry in the Tenant Wide Extensions list on the Tenant App Catalog titled `GoogleAnalytics`.
2. Delete the `googleanalytics-tenant.sppkg` file from the 'Apps for SharePoint' library in your Tenant App Catalog.


## Disclaimer

Microsoft provides programming examples for illustration only, without warranty either expressed or implied, including, but not limited to, the implied warranties of merchantability and/or fitness for a particular purpose. We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys' fees, that arise or result from the use or distribution of the Sample Code.
