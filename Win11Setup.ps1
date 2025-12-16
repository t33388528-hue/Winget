#Getting pw for autologon
$Password = (Get-Credential).GetNetworkCredential().Password
Write-Host "Password = $Password | Abort: Ctrl+C, continuing in 10s..."
Start-Sleep -Seconds 10
clear

#Starting background tasks
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
irm https://raw.githubusercontent.com/t33388528-hue/Winget/refs/heads/main/WinUpdate.ps1|iex

#Autologon
$Username = $env:USERNAME
$Domain   = $env:USERDOMAIN

$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"

Set-ItemProperty -Path $RegPath -Name "AutoAdminLogon" -Value "1"
Set-ItemProperty -Path $RegPath -Name "DefaultUsername" -Value $Username
Set-ItemProperty -Path $RegPath -Name "DefaultPassword" -Value $Password
Set-ItemProperty -Path $RegPath -Name "DefaultDomainName" -Value $Domain

Write-Host "Autologon activated, do not cancel this script!"

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

#Winupdate post reboot
&{
#Thanks Trend...
$str = irm https://raw.githubusercontent.com/t33388528-hue/Winget/refs/heads/main/WinUpdate.ps1
$ubrLink = 'noisreVtnerruC\TN swodniW\tfosorciM\ERAWTFOS\:MLKH'
$ubr = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").UBR
$TaskName = "Win11SetupReboot"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -Command `"`$string = 'noisreVtnerruC\TN swodniW\tfosorciM\ERAWTFOS\:MLKH'; if ((Get-ItemProperty (-join `$string[-1..-(`$string.Length)])).UBR -ne $ubr -or (Get-ItemProperty (-join `$string[-1..-(`$string.Length)])).UBR -eq 7462) {schtasks /Delete /TN $TaskName /F; Enable-ScheduledTask -TaskName 'Win11SetupPost'}; msg * 'Autologon is still running, do not cancel this script!'; Start-Sleep -Seconds 30; $str; shutdown -r -t 10`""
$Trigger = New-ScheduledTaskTrigger -AtLogon -RandomDelay (New-TimeSpan -Seconds 10)
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Force
}

#Language stuff
&{
$TaskName = "Win11SetupPost"
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"schtasks /Delete /TN $TaskName /F; Start-Process 'ms-settings:windowsupdate'; taskkill /IM tvsukernel.exe /F; Start-Process tvsu.exe; ; Start-Process intl.cpl; Remove-ItemProperty -Path '$global:RegPath' -Name 'AutoAdminLogon'; Set-ItemProperty -Path '$global:RegPath' -Name 'DefaultUsername' -Value ''; Remove-ItemProperty -Path '$global:RegPath' -Name 'DefaultPassword'; Start-Sleep -Seconds 10; Start-Process winver.exe; Set-WinUILanguageOverride nb-NO`""
$Trigger = New-ScheduledTaskTrigger -AtLogon -RandomDelay (New-TimeSpan -Seconds 10)
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Force
Disable-ScheduledTask -TaskName $TaskName
}

Start-Process powershell "irm bit.ly/WinTeams|iex" -WindowStyle Minimized
shutdown -r -t 600 -c "Restarting in 10 minutes to apply updates."
gpupdate /force

