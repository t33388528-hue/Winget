if ($env:COMPUTERNAME[0] -eq "E"){
Start-Process powershell "cscript '\\capa-edu\PRODCON\ComputerJobs\DameWare Mini Remote Control Service\v12.2.2.12\Scripts\DameWare Mini Remote Control Service.cis'" -WindowStyle Minimized
Start-Process "\\edu-fil01\brukere$\iktadm\system_update_5.08.03.59.exe" -ArgumentList "/VERYSILENT" -Wait
Start-Process powershell "Start-Process 'C:\Program Files (x86)\Lenovo\System Update\tvsu.exe' '/CM /Install'" -WindowStyle Minimized
}else{
Start-Process powershell "cscript '\\ikt-drift01\PRODCON\ComputerJobs\DameWare Mini Remote Control Service\v12.2.2.12\Scripts\DameWare Mini Remote Control Service.cis'" -WindowStyle Minimized
Start-Process powershell "Start-Process 'C:\Program Files (x86)\Lenovo\System Update\tvsu.exe' '/CM /Install'" -WindowStyle Minimized
}
Start-Process powershell "Add-Type -AssemblyName System.Windows.Forms; while (`$true) {[System.Windows.Forms.SendKeys]::SendWait('{SCROLLLOCK}'); Start-Sleep -Seconds 59}" -WindowStyle Minimized

#Windows Update stuff 
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

#Delay cus of winupdate crash i think
if ($env:COMPUTERNAME[0] -eq "E"){
Start-Process powershell "DISM /Online /Add-Package /PackagePath:'\\edu-fil01\brukere`$\iktadm\Microsoft-Windows-Client-Language-Pack_x64_nb-no.cab'" -WindowStyle Minimized
}else{
Start-Process powershell "DISM /Online /Add-Package /PackagePath:'\\hk-fil\felles\Personal\IKT\Microsoft-Windows-Client-Language-Pack_x64_nb-no.cab'" -WindowStyle Minimized
}

#System update stuff (does 100% of the updates, 50% of the time)
taskkill /IM tvsukernel.exe /F
New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\LENOVO\System Update\Preferences\UserSettings\General" -Name "MetricsEnabled" -Value "NO" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\LENOVO\System Update\Preferences\UserSettings\General" -Name "DisplayLicenseNotice" -Value "NO" -PropertyType String -Force | Out-Null
Start-Sleep -Seconds 5
Start-Process tvsu.exe
$myshell = New-Object -ComObject WScript.Shell
Start-Sleep -Seconds 30
$myshell.AppActivate("System Update") 
Start-Sleep -Seconds 2
$myshell.SendKeys("N")
Start-Sleep -Seconds 100
$myshell.AppActivate("System Update") 
$myshell.SendKeys("S")
Start-Sleep -Seconds 5
$myshell.SendKeys("^{TAB}")
Start-Sleep -Seconds 5
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
&{
$TaskName = "TempLogonTask"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"schtasks /Delete /TN $TaskName /F; Start-Process 'ms-settings:windowsupdate'; taskkill /IM tvsukernel.exe /F; Start-Process tvsu.exe; ; Start-Process intl.cpl; Start-Process winver.exe; Set-WinUILanguageOverride nb-NO`""
$Trigger = New-ScheduledTaskTrigger -AtLogon -RandomDelay (New-TimeSpan -Seconds 10)
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Force
}

#Winupdate post reboot
&{
$TaskName = "TempLogonTask2"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -Command `"try{schtasks /Delete /TN $TaskName /F; `$retries = 0; Write-Host 'Searching for Windows Updates...'; while (`$retries -lt 2){`$UpdateSession = New-Object -ComObject Microsoft.Update.Session;`$UpdateSearcher = `$UpdateSession.CreateUpdateSearcher(); `$SearchResult = `$UpdateSearcher.Search('IsInstalled=0'); foreach (`$update in `$SearchResult.Updates) {Write-Host '- `$(`$update.Title)'};if (`$SearchResult.Updates.Count -eq 0){msg * 'No more updates found.'; break;};`$UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl;foreach (`$update in `$SearchResult.Updates) {`$UpdatesToDownload.Add(`$update) | Out-Null};`$Downloader = `$UpdateSession.CreateUpdateDownloader();`$Downloader.Updates = `$UpdatesToDownload;`$DownloadResult = `$Downloader.Download();`$Installer = `$UpdateSession.CreateUpdateInstaller();`$Installer.Updates = `$UpdatesToDownload;`$InstallationResult = `$Installer.Install(); `$global:retries++}; msg * 'Windows Updates completed.'; shutdown -r -t 120 }catch{msg * '`$(`$_.Exception.Message)'}`""
$Trigger = New-ScheduledTaskTrigger -AtLogon -RandomDelay (New-TimeSpan -Seconds 10)
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Force
}

Start-Process powershell "iwr bit.ly/WinTeams|iex" -WindowStyle Minimized
shutdown -r -t 600 -c "Restarting in 10 minutes to apply updates."
gpupdate /force

