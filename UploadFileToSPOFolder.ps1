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
    $localfile,
    $spositeUrl,
    $spofolderpath,
    $username,
    $password
)

$ErrorActionPreference = "Stop"
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
[void][System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client.Runtime, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")


$script:context = New-Object Microsoft.SharePoint.Client.ClientContext($spositeUrl)
$secpass = ConvertTo-SecureString $password -AsPlainText -Force
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




$web = $context.Web
$spofolder = $web.GetFolderByServerRelativeUrl($spofolderpath)
$script:context.Load($web)
$script:context.Load($spofolder)
ExecuteQueryWithIncrementalRetry -retryCount 5

$leafname = Split-Path -Leaf $localfile
$uploadFile = $null

$inputFile = get-item $localfile

$blockSize = 1000000 # 1MB
$uploadId = (New-Guid)
$fileSize = $inputFile.Length



if ($fileSize -le $blockSize)
{
    $fs = New-Object System.IO.FileStream($localfile, [System.IO.FileMode]::Open)

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
        $fs = [System.IO.File]::Open($localfile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
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

