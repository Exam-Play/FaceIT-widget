Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "server.exe" & chr(34), 0
Set WshShell = Nothing