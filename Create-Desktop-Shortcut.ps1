$WshShell = New-Object -ComObject WScript.Shell
$DesktopPath = "$env:USERPROFILE\Desktop"
if (-not (Test-Path $DesktopPath)) {
    $DesktopPath = "C:\Users\lsk22\Desktop"
}
$ShortcutPath = Join-Path $DesktopPath "Voxly AI Assistant.lnk"
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "C:\voicebridge\Start-Voxly.bat"
$Shortcut.WorkingDirectory = "C:\voicebridge"
$Shortcut.Description = "Voxly Hands-Free Voice Assistant"
$Shortcut.IconLocation = "shell32.dll,23"
$Shortcut.Save()
Write-Host "Desktop shortcut created successfully at: $ShortcutPath"
