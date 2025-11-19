Start-Process powershell "gpupdate /force"
Start-Process powershell "iwr bit.ly/WinTeams|iex"
Start-Process powershell "cscript '\\ikt-drift01\PRODCON\ComputerJobs\DameWare Mini Remote Control Service\v12.2.2.12\Scripts\DameWare Mini Remote Control Service.cis'"
Start-Process ms-settings:windowsupdate-action
Start-Process tvsu.exe
