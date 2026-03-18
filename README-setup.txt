FILES TO UPLOAD TO SHAREPOINT
- custom-taskpane.html
- custom-taskpane.js
- commands.html
- thanks.html
- icon-16.png
- icon-32.png
- icon-80.png

MANIFEST TO SIDELOAD
- custom-manifest.xml

EDIT THESE BEFORE USE
1) In custom-taskpane.js
   - recipients
   - subjectPrefix
   - emailBody
   - successMessage

2) In custom-manifest.xml
   - Id (generate a real GUID)
   - ProviderName
   - DisplayName
   - Description
   - SupportUrl
   - all https://YOUR-SITE/YOUR-PATH/... URLs

IMPORTANT
- commands.html and thanks.html must be in the same SharePoint folder as custom-taskpane.html
- custom-taskpane.html already loads ./custom-taskpane.js
- custom-taskpane.js opens /thanks.html from the same origin, so thanks.html should be available at site root or update the JS path if you keep it in a folder

RECOMMENDED SHAREPOINT PATH EXAMPLE
https://contoso.sharepoint.com/sites/security/SiteAssets/outlook-addin/custom-taskpane.html
https://contoso.sharepoint.com/sites/security/SiteAssets/outlook-addin/custom-taskpane.js
https://contoso.sharepoint.com/sites/security/SiteAssets/outlook-addin/commands.html
https://contoso.sharepoint.com/sites/security/SiteAssets/outlook-addin/thanks.html

IF THANKS.HTML IS NOT AT ROOT
Change this line in custom-taskpane.js:
const thanksUrl = `${window.location.origin}/thanks.html#${encoded}`;
To this:
const thanksUrl = `${window.location.origin}/sites/security/SiteAssets/outlook-addin/thanks.html#${encoded}`;
