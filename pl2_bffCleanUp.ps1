# Script for cleaning pl2_bff share from files older than 30 days
# Author: Sławomir Abrich
# Date: 05.02.2024

$year = Get-Date -Format 'yyyy'
$yearOfYesterday = (Get-Date).AddDays(-1).Year.ToString()
$month = Get-Date -Format 'MM'
$monthOfYesterday = (Get-Date).AddDays(-1).Month.ToString()
if ($monthOfYesterday.Length -eq 1)
    {$monthOfYesterday = "0" + $monthOfYesterday}
$day = Get-Date -Format 'dd'
$dayOfYesterday = (Get-Date).AddDays(-1).Day.ToString()
if ($dayOfYesterday.Length -eq 1)
    {$dayOfYesterday = "0" + $dayOfYesterday}
$hour = Get-Date -Format 'HH'
$minute = Get-Date -Format 'mm'
$second = Get-Date -Format 'ss'
$timeStamp = "$year-$month-$day ${hour}:${minute}:${second}"
$sourceHost = $env:COMPUTERNAME
$logFile = "\\aligntech.com\pl2afab\IT\logs\prdpl2scripts01\pl2_bffCleanUp.log"
$destinationPath = '\\aligntech.com\pl2afab\pl2_bff\'
$olderThan = (Get-Date).AddDays(-30)
$fileList = Get-ChildItem -Path $destinationPath -Recurse -File | Select-Object Name,LastAccessTime | Where-Object {$_.LastAccessTime -lt $olderThan}
$fileCount = $fileList.Count
$directoryList = Get-ChildItem -Path $destinationPath -Recurse -Directory | Select-Object Name,LastAccessTime | Where-Object {$_.LastAccessTime -lt $olderThan}
$directoryCount = $directoryList.Count

if ($fileCount -eq 0 -and $directoryCount -eq 0) {
    Write-Output "${timeStamp} - There was no files and folders older than ${olderThan} - nothing to delete." | Out-File $logFile -Append
    }
else {
    Get-ChildItem -Path $destinationPath -Recurse | Where LastAccessTime -lt $olderThan | Remove-Item -Force -Confirm:$false
    Write-Output "${timeStamp} - ${directoryCount} directories older than ${olderThan} deleted." | Out-File $logFile -Append
    Write-Output "${timeStamp} - ${fileCount} files older than ${olderThan} deleted." | Out-File $logFile -Append
    }
