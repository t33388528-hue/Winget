Start-Process powershell "gpupdate /force"
Start-Process powershell "iwr bit.ly/WinTeams|iex"
Start-Process powershell "cscript '\\ikt-drift01\PRODCON\ComputerJobs\DameWare Mini Remote Control Service\v12.2.2.12\Scripts\DameWare Mini Remote Control Service.cis'"
Start-Process powershell "Start-Process 'C:\Program Files (x86)\Lenovo\System Update\tvsu.exe' '/CM /Install'"

Write-Host "Searching for Windows Updates..."

# Create update session
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

# Search for updates
$SearchResult = $UpdateSearcher.Search("IsInstalled=0")

foreach ($update in $SearchResult.Updates) {
    Write-Host "- $($update.Title)"
}

Write-Host "Found and downloading $($SearchResult.Updates.Count) updates..."

# Install all available updates
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
