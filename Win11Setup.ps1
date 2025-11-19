Start-Process powershell "gpupdate /force"
Start-Process powershell "iwr bit.ly/WinTeams|iex"
Start-Process powershell "cscript '\\ikt-drift01\PRODCON\ComputerJobs\DameWare Mini Remote Control Service\v12.2.2.12\Scripts\DameWare Mini Remote Control Service.cis'"
if (-not (Test-Path "C:\Program Files (x86)\Lenovo\System Update\tvsu.exe")) {
    cscript '\\ikt-drift01\PRODCON\ComputerJobs\Lenovo System update\v5.07.0110\Scripts\Lenovo System update.cis'
}
Start-Process powershell "Start-Process 'C:\Program Files (x86)\Lenovo\System Update\tvsu.exe' '/CM /Install'"

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

# Prompt before installation
Write-Host "`nStarting download and installation of $($SearchResult.Updates.Count) updates..."

# Collect updates
$UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
foreach ($update in $SearchResult.Updates) {
    $UpdatesToDownload.Add($update) | Out-Null
}

$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$Downloader.Download()

Write-Host "Installing updates..."

$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToDownload
$InstallationResult = $Installer.Install()

Write-Host "Installation Result: $($InstallationResult.ResultCode)"
Write-Host "Restarting in 5 minutes."

shutdown -r -t 300
