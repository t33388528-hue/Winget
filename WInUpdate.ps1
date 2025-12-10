$retries = 0

while ($retries -lt 2){
Write-Host "------------------------------------------------"
Write-Host "Searching for Windows Updates..."

$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

$SearchResult = $UpdateSearcher.Search("IsInstalled=0")

foreach ($update in $SearchResult.Updates) {
    Write-Host "- $($update.Title)"
}

if ($SearchResult.Updates.Count -eq 0){
Write-Host "No Updates found."
break
}

$UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
foreach ($update in $SearchResult.Updates) {
    $UpdatesToDownload.Add($update) | Out-Null
}

$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
Write-Host "`nStarting download of $($Downloader.Updates.Count)/$($SearchResult.Updates.Count) updates..."
$DownloadResult = $Downloader.Download()

Write-Host "Installing updates..."

$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToDownload
$InstallationResult = $Installer.Install()

$global:retries++
}

Write-Host "Windows Updates completed."
