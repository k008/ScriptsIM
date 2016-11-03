msgbox EnvironmentVariables("%ProgramFiles%")

ProgrammFiles=EnvironmentVariables("%ProgramFiles%")

RunOutEx("'" & ProgrammFiles & "\7-Zip\7z.exe a -tzip D:\POST\MoveFilesPost.log.zip D:\POST\MoveFilesPost.log '")

Function EnvironmentVariables(fvar)
  Set WshShell = WScript.CreateObject("WScript.Shell")
  EnvironmentVariables=WshShell.ExpandEnvironmentStrings(fvar)
End Function

Function RunOutEx(cmd)
	Dim WshShell
	set WshShell = WScript.CreateObject("WScript.Shell")
	msgbox(cmd)
	WshShell.Run cmd
End Function

---------------------------


'C:\Program Files (x86)\7-Zip\7z.exe a -tzip D:\POST\MoveFilesPost.log.zip D:\POST\MoveFilesPost.log '
