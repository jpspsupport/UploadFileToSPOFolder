# UploadToSPOFolder

This is a sample CSOM PowerShell Script to upload a large file to SharePoint Online folder.

### Note
If you really want to move large amount of files to SharePoint Online, CSOM call is not the right solution.
In that case, please consider to use SharePoint Migration Tools (SPMT).

## Prerequitesite
You need to download SharePoint Online Client Components SDK to run this script.
https://www.microsoft.com/en-us/download/details.aspx?id=42038

You can also acquire the latest SharePoint Online Client SDK by Nuget as well.

1. You need to access the following site. 
https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM

2. Download the nupkg.
3. Change the file extension to *.zip.
4. Unzip and extract those file.
5. place them in the specified directory from the code. 

## How to Run - parameters

-localfile ... The target local file (or folder) path to upload. 

-spositeurl ... The target SPO site to upload file.

-spofolderpath ... The existing target SPO folder to upload the file. 

-SkipRootFolder ... (OPTIONAL) Skip the root folder creation. Only effective when uploading folder.

-username ... (OPTIONAL) The target user to upload the file.

-password ... (OPTIONAL) The password of the above user.

## Example
.\UploadFileToSPOFolder.ps1 -localfile .\sampledata.txt -spositeUrl https://tenant.sharepoint.com/sites/siteA -spofolderpath Shared%20Documents/myfolder


## Reference
Please also refer to the following docs.
https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/upload-large-files-sample-app-for-sharepoint


