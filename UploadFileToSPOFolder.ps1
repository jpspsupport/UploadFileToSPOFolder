<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 

 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.

Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>
param(
    [Parameter(Mandatory=$true)]
    $localfile,
    [Parameter(Mandatory=$true)]
    $spositeUrl,
    [Parameter(Mandatory=$true)]
    $spofolderpath,
    [switch] $SkipRootFolder,
    [Parameter(Mandatory=$false)]
    $username,
    [Parameter(Mandatory=$false)]
    $password
)

$ErrorActionPreference = "Stop"
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client.Runtime, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")


$script:context = New-Object Microsoft.SharePoint.Client.ClientContext($spositeUrl)
$pwd = $null
if ($username -eq $null)
{
  $cred = Get-Credential
  $username = $cred.UserName
  $secpass = $cred.Password
}
else
{
  $secpass = convertto-securestring $password -AsPlainText -Force
}
$script:context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $secpass)


$script:context.add_ExecutingWebRequest({
    param ($source, $eventArgs);
    $request = $eventArgs.WebRequestExecutor.WebRequest;
    $request.UserAgent = "NONISV|Contoso|Application/1.0";
})

function ExecuteQueryWithIncrementalRetry {
    param (
        [parameter(Mandatory = $true)]
        [int]$retryCount
    );

    $DefaultRetryAfterInMs = 120000;
    $RetryAfterHeaderName = "Retry-After";
    $retryAttempts = 0;

    if ($retryCount -le 0) {
        throw "Provide a retry count greater than zero."
    }

    while ($retryAttempts -lt $retryCount) {
        try {
            $script:context.ExecuteQuery();
            return;
        }
        catch [System.Net.WebException] {
            $response = $_.Exception.Response

            if (($null -ne $response) -and (($response.StatusCode -eq 429) -or ($response.StatusCode -eq 503))) {
                $retryAfterHeader = $response.GetResponseHeader($RetryAfterHeaderName);
                $retryAfterInMs = $DefaultRetryAfterInMs;

                if (-not [string]::IsNullOrEmpty($retryAfterHeader)) {
                    if (-not [int]::TryParse($retryAfterHeader, [ref]$retryAfterInMs)) {
                        $retryAfterInMs = $DefaultRetryAfterInMs;
                    }
                    else {
                        $retryAfterInMs *= 1000;
                    }
                }

                Write-Output ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($retryAfterInMs / 1000))
                #Add delay.
                Start-Sleep -m $retryAfterInMs
                #Add to retry count.
                $retryAttempts++;
            }
            else {
                throw;
            }
        }
    }

    throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}

function UploadFile($inputFile, $spofolder)
{
    $blockSize = 100000000 # 100MB
    $uploadId = (New-Guid)
    $fileSize = $inputFile.Length
    $uploadFile = $null
    $leafname = $inputFile.Name

    if ($fileSize -le $blockSize)
    {
        $fs = New-Object System.IO.FileStream($inputFile.FullName, [System.IO.FileMode]::Open)

        $fileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
        $fileInfo.ContentStream = $fs
        $fileInfo.Url = $leafname
        $fileInfo.Overwrite = $true
        $uploadFile = $spofolder.Files.Add($fileInfo)
        
        $script:context.Load($uploadFile)
        ExecuteQueryWithIncrementalRetry -retryCount 5
        $fs.Dispose()

    }
    else 
    {
        $bytesUploaded = $null
        $fs = $null
        try
        {
            $fs = [System.IO.File]::Open($inputFile.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $br = New-Object System.IO.BinaryReader($fs)
            $buffer = [System.Byte[]]::CreateInstance([System.Byte], $blockSize)
            $lastBuffer = $null
            $fileoffset = 0
            $totalBytesRead = 0
            $bytesRead
            $first = $true
            $last = $false

            while(($bytesRead = $br.Read($buffer, 0, $buffer.Length)) -gt 0)
            {
                $totalBytesRead = $totalBytesRead + $bytesRead;

                if ($totalBytesRead -eq $fileSize)
                {
                    $last = $true
                    $lastBuffer = [System.Byte[]]::CreateInstance([System.Byte], $bytesRead)
                    [System.Array]::Copy($buffer, 0, $lastBuffer, 0, $bytesRead)
                }

                if ($first)
                {
                    $contentStream = New-Object System.IO.MemoryStream

                    $fileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                    $fileInfo.ContentStream = $contentStream
                    $fileInfo.Url = $leafname
                    $fileInfo.Overwrite = $true
                    $uploadFile = $spofolder.Files.Add($fileInfo)

                    $s = New-Object System.IO.MemoryStream
                    $s.Write($buffer, 0, $buffer.Length)
                    $s.Position = 0

                    $bytesUploaded = $uploadFile.StartUpload($uploadId, $s);
                    ExecuteQueryWithIncrementalRetry -retryCount 5
                    $fileoffset = $bytesUploaded.Value
                    $s.Dispose()

                    $first = $false
                }
                else
                {
                    $uploadFile = $script:context.Web.GetFileByServerRelativeUrl(($spofolder.ServerRelativeUrl + "/" + $leafname));

                    if ($last)
                    {
                        $s = New-Object System.IO.MemoryStream
                        $s.Write($lastBuffer, 0, $lastBuffer.Length)
                        $s.Position = 0
                        $uploadFile = $uploadFile.FinishUpload($uploadId, $fileoffset, $s);
                        ExecuteQueryWithIncrementalRetry -retryCount 5 
                        $s.Dispose()
                    }
                    else 
                    {
                        $s = New-Object System.IO.MemoryStream
                        $s.Write($buffer, 0, $buffer.Length)
                        $s.Position = 0

                        $bytesUploaded = $uploadFile.ContinueUpload($uploadId, $fileoffset, $s);
                        ExecuteQueryWithIncrementalRetry -retryCount 5 
                        $fileoffset = $bytesUploaded.Value;
                        $s.Dispose()
                    }
                }
            }
        }
        finally
        {
            if ($fs -ne $null)
            {
                $fs.Dispose()
            }
        }
    }
}

function CreateSPOFolder($localFolderName, $spoParentFolder)
{
    $web = $script:context.Web
    $spofolder = $web.GetFolderByServerRelativeUrl($spoParentFolder.Replace(($spositeUrl + "/"), ""))
    $folder = $spofolder.Folders.Add($localFolderName)
    $script:context.Load($folder)
    ExecuteQueryWithIncrementalRetry -retryCount 5
}

function GetSPOFolder($spofolderpath)
{
    $web = $script:context.Web
    $spofolder = $web.GetFolderByServerRelativeUrl($spofolderpath.Replace(($spositeUrl + "/"), ""))
    $script:context.Load($spofolder)
    ExecuteQueryWithIncrementalRetry -retryCount 5
    return $spofolder
}

function GetLocalParentFolder($currentItem)
{
    if ($currentItem.Mode.StartsWith("d"))
    {
        return $currentItem.Parent.FullName
    }
    else {
        return $currentItem.DirectoryName
    }
}

function GetSPORelativeFolderPath($spoRootFolder, $localRootFolder, $localPath)
{
    return ($spoRootFolder.TrimEnd("/") + $localPath.Replace($localRootFolder, "").Replace("`\", "/"))
}

function UploadFolder($localRootFolder, $spoRootFolderPath, $skipRootFolder)
{
    $localRootFolderPath = $localRootFolder.FullName
    # Get the list of target folders/files
    $childitems = Get-ChildItem $localRootFolderPath -Recurse
    $parentSPOFolder = $spoRootFolderPath
    $parentLocalFolder = $localRootFolder.Parent.FullName

    if (!$skipRootFolder)
    {
        # Create Root Folder
        CreateSPOFolder -localFolderName $localRootFolder.Name -spoParentFolder $spoRootFolderPath
        # Change the Local Base Folder Path
        $localRootFolderPath = $localRootFolder.Parent.FullName
    }

    foreach ($citem in $childitems)
    {
        # Get Parent Folder
        $thisParent = GetLocalParentFolder -currentItem $citem
        if ($parentLocalFolder -ne $thisParent)
        {
            # Convert from local path to SPO destination path.
            $parentSPOFolder = GetSPORelativeFolderPath -spoRootFolder $spoRootFolderPath -localRootFolder $localRootFolderPath -localPath $thisParent
            # Change the SPO Target Folder here.
            $spofolder = GetSPOFolder -spofolderpath $parentSPOFolder
            $parentLocalFolder = $thisParent
        }

        if ($citem.Mode.StartsWith("d"))
        {
            CreateSPOFolder -localFolderName $citem.Name -spoParentFolder $parentSPOFolder
        }
        else {
            UploadFile -inputFile $citem -spofolder $spofolder
        }
    }
}

$inputFile = get-item $localfile
if (!$inputFile.Mode.StartsWith("d"))
{
    #Single File Upload
    $spofolder = GetSPOFolder -spofolderpath $spofolderpath
    UploadFile -inputFile $inputFile -spofolder $spofolder
}
else {
    #Folder Upload
    UploadFolder -localRootFolder $inputFile -spoRootFolderPath $spofolderpath -skipRootFolder $SkipRootFolder
}
