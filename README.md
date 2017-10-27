# SPOUtilities

In the spirit of "Sharing is caring" .... Thanks to Vesa Juvonen and the Office Dev PnP Team for the amazing things that you do for the SharePoint/Office 365 community.

This is a Console App (based on .NET Framework). It contains references to the latest SharePoint Online CSOM package.  As of now, it provides the following utilities for managing your SharePoint Online site.

- Enable Major and Minor Versions in Document Library. This enables 5 major versions, and retains drafts for 2 major versions.
- Delete Old Document Versions in "Documents" library. This deletes "older" versions (i.e. greater than 15) documents in the library.
- Get Last Modified Information from all Sites and "Document" Library in a Site Collection. This recurses through all Webs.

The Console App also writes log information to a CSV file.

I will be adding more utilities, and will share them.

To use this:

1. Make sure to add the latest SharePoint Online CSOM package to the project. The NuGet command to install the package is shown below:
Install-Package Microsoft.SharePointOnline.CSOM

2. Add references to System.Web and System.Configuration assemblies.

3. Update appSettings section in 'App.config' file with your SharePoint Online information.

  
Feel free to send your questions/comments to kkakanur@rightpoint.com

I am proud to be working for a values-led intrapreneurial organization, Rightpoint - https://www.rightpoint.com/company 

We believe in the spirit of Makers. 

We are based out of Chicago, with offices in Atlanta, Boston, Dallas, Denver, Detroit, Los Angeles, New York, and Jaipur (India).
We’re driven by innovation, rooted in technology, relentlessly curious and celebrating our 10th anniversary.  


