Start-Process powershell "cscript '\\ikt-drift01\PRODCON\ComputerJobs\DameWare Mini Remote Control Service\v12.2.2.12\Scripts\DameWare Mini Remote Control Service.cis'" -WindowStyle Minimized
Start-Process powershell "Start-Process 'C:\Program Files (x86)\Lenovo\System Update\tvsu.exe' '/CM /Install'" -WindowStyle Minimized
Start-Process powershell "Add-Type -AssemblyName System.Windows.Forms; while (`$true) {[System.Windows.Forms.SendKeys]::SendWait('{SCROLLLOCK}'); Start-Sleep -Seconds 59}" -WindowStyle Minimized
Start-Process powershell "DISM /Online /Add-Package /PackagePath:'\\hk-fil\felles\Personal\IKT\Microsoft-Windows-Client-Language-Pack_x64_nb-no.cab'" -WindowStyle Minimized

#Windows Update stuff
$retries = 0

while ($retries -lt 3){
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

#System update stuff (does 100% of the updates, 50% of the time)
New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\LENOVO\System Update\Preferences\UserSettings\General" -Name "MetricsEnabled" -Value "NO" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\LENOVO\System Update\Preferences\UserSettings\General" -Name "DisplayLicenseNotice" -Value "NO" -PropertyType String -Force | Out-Null
Start-Sleep -Seconds 5
Start-Process tvsu.exe
$myshell = New-Object -ComObject WScript.Shell
Start-Sleep -Seconds 30
$myshell.AppActivate("System Update") 
Start-Sleep -Seconds 2
$myshell.SendKeys("N")
Start-Sleep -Seconds 200
$myshell.AppActivate("System Update") 
$myshell.SendKeys("S")
Start-Sleep -Seconds 5
$myshell.SendKeys("^{TAB}")
Start-Sleep -Seconds 5
$myshell.SendKeys("S")
Start-Sleep -Seconds 5
$myshell.SendKeys("N")
Start-Sleep -Seconds 5
$myshell.SendKeys("D")
Start-Sleep -Seconds 5
$myshell.SendKeys("{Enter}")

#Language stuff
$TaskName = "TempLogonTask"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"schtasks /Delete /TN $TaskName /F; Start-Process powershell 'iwr bit.ly/WinTeams|iex'; Start-Process intl.cpl; msg * 'Copy settings and reboot to apply language'; Set-WinUILanguageOverride nb-NO`""
$Trigger = New-ScheduledTaskTrigger -AtLogon -RandomDelay (New-TimeSpan -Seconds 10)
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Force

shutdown -r -t 600
gpupdate /force

