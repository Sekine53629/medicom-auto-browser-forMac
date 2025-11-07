# ショートカット作成スクリプト
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\Medicom Auto Browser.lnk")
$Shortcut.TargetPath = "C:\Users\0053629\Documents\GitHub\medicom-auto-browser-forMac\run.bat"
$Shortcut.WorkingDirectory = "C:\Users\0053629\Documents\GitHub\medicom-auto-browser-forMac"
$Shortcut.Description = "Medicom Auto Browser"
$Shortcut.Save()

Write-Host "Shortcut created on Desktop!"
