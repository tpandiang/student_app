$WshShell = New-Object -ComObject WScript.Shell
$StartupPath = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup\ResumeWebServer.lnk")
$Shortcut = $WshShell.CreateShortcut($StartupPath)
$Shortcut.TargetPath = "C:\Users\txp190010\my-webpage\start-server.vbs"
$Shortcut.WorkingDirectory = "C:\Users\txp190010\my-webpage"
$Shortcut.Description = "Resume Website Server on port 9000"
$Shortcut.WindowStyle = 7
$Shortcut.Save()
Write-Host "Startup shortcut created at: $StartupPath"
