Start-Process powershell "iwr bit.ly/WinTeams|iex" -WindowStyle Minimized
Start-Process powershell "cscript '\\ikt-drift01\PRODCON\ComputerJobs\DameWare Mini Remote Control Service\v12.2.2.12\Scripts\DameWare Mini Remote Control Service.cis'" -WindowStyle Minimized
Start-Process powershell "Start-Process 'C:\Program Files (x86)\Lenovo\System Update\tvsu.exe' '/CM /Install'" -WindowStyle Minimized
Start-Process powershell "Add-Type -AssemblyName System.Windows.Forms; while (`$true) {[System.Windows.Forms.SendKeys]::SendWait('{SCROLLLOCK}'); Start-Sleep -Seconds 3}" -WindowStyle Minimized

$retries = 0

while ($retries -lt 3){
Write-Host "------------------------------------------------"
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

if ($SearchResult.Updates.Count -eq 0){
Write-Host "No Updates found."
break
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

Write-Host "Installing updates..."

$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToDownload
$InstallationResult = $Installer.Install()

$global:retries++
}

Write-Host "Windows Updates completed."

Start-Process tvsu.exe
Start-Sleep -Seconds 30
$myshell = New-Object -ComObject WScript.Shell
taskkill /IM tvsukernel.exe
Start-Sleep -Seconds 2
$myshell.SendKeys("N")
Start-Sleep -Seconds 100
$myshell.SendKeys("{ENTER}")
Start-Sleep -Seconds 10
$myshell.SendKeys("{Right}")
Start-Sleep -Seconds 2
$myshell.SendKeys("{Right}")
Start-Sleep -Seconds 2
$myshell.SendKeys("{Right}")
Start-Sleep -Seconds 2
$myshell.SendKeys("S")
Start-Sleep -Seconds 2
$myshell.SendKeys("N")
Start-Sleep -Seconds 2
$myshell.SendKeys("D")
Start-Sleep -Seconds 2
$myshell.SendKeys("{Enter}")
shutdown -r -t 600

gpupdate /force

