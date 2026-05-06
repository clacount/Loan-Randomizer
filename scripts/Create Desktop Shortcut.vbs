Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

repoRoot = fso.GetParentFolderName(WScript.ScriptFullName)
launcherPath = fso.BuildPath(repoRoot, "Open LendingFair.cmd")
desktopPath = shell.SpecialFolders("Desktop")
shortcutPath = fso.BuildPath(desktopPath, "LendingFair.lnk")

If Not fso.FileExists(launcherPath) Then
  MsgBox "Could not find launcher script:" & vbCrLf & launcherPath, vbCritical, "LendingFair"
  WScript.Quit 1
End If

Set shortcut = shell.CreateShortcut(shortcutPath)
shortcut.TargetPath = launcherPath
shortcut.WorkingDirectory = repoRoot
shortcut.WindowStyle = 1
shortcut.Description = "Open LendingFair"
shortcut.IconLocation = "%SystemRoot%\System32\SHELL32.dll,220"
shortcut.Save

MsgBox "Created Desktop shortcut:" & vbCrLf & shortcutPath, vbInformation, "LendingFair"
