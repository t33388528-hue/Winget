Start-Process powershell 'Add-Type -AssemblyName System.Windows.Forms; while (`$true) {[System.Windows.Forms.SendKeys]::SendWait(''{SCROLLLOCK}''); Start-Sleep -Seconds 59}' -WindowStyle Minimized
