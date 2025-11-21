Start-Process powershell "gpupdate /force" -WindowStyle Minimized
Start-Process powershell "iwr bit.ly/WinTeams|iex" -WindowStyle Minimized
Start-Process powershell "cscript '\\ikt-drift01\PRODCON\ComputerJobs\DameWare Mini Remote Control Service\v12.2.2.12\Scripts\DameWare Mini Remote Control Service.cis'" -WindowStyle Minimized
Start-Process powershell "Start-Process 'C:\Program Files (x86)\Lenovo\System Update\tvsu.exe' '/CM /Install'" -WindowStyle Minimized

$failedDownloads = 0
$failedInstalls = 0
$retries = 0

while (($failedDownloads -ne 0 -OR $failedInstalls -ne 0) -AND $retries -lt 2){
Write-Host "Searching for Windows Updates..."

# Create update session
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

# Search for updates
$SearchResult = $UpdateSearcher.Search("IsInstalled=0")

# Print the names of all updates
foreach ($update in $SearchResult.Updates) {
    Write-Host "- $($update.Title)"
}

Write-Host "`nStarting download of $($SearchResult.Updates.Count) updates..."

# Collect updates
$UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
foreach ($update in $SearchResult.Updates) {
    $UpdatesToDownload.Add($update) | Out-Null
}

$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$DownloadResult = $Downloader.Download()

# Count failed downloads
$global:failedDownloads = 0
for ($i = 0; $i -lt $DownloadResult.UpdateResult.Count; $i++) {
    if ($DownloadResult.UpdateResult.Item($i).ResultCode -ne 2) { # 2 = succeeded
        $failedDownloads++
        Write-Host "Download failed for: $($SearchResult.Updates.Item($i).Title)"
    }
}
Write-Host "Failed downloads: $failedDownloads"

Write-Host "Installing updates..."

$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToDownload
$InstallationResult = $Installer.Install()

# Count failed installs
$global:failedInstalls = 0
for ($i = 0; $i -lt $InstallationResult.UpdateResult.Count; $i++) {
    if ($InstallationResult.UpdateResult.Item($i).ResultCode -ne 2) { # 2 = succeeded
        $failedInstalls++
        Write-Host "Install failed for: $($SearchResult.Updates.Item($i).Title)"
    }
}
Write-Host "Failed installs: $failedInstalls"
$global:retries++
}

Write-Host "Restarting in 5 minutes."

shutdown -r -t 300
