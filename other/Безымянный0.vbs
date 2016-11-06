'msgbox EnvironmentVariables("%Programw6432%")

'ProgrammFiles=EnvironmentVariables("%ProgramFiles%")


Set WSS = CreateObject("Wscript.Shell")
xOS = "x64"
ProgrammFiles=EnvironmentVariables("%Programw6432%")

If WSS.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") = "x86" AND WSS.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%") = "%PROCESSOR_ARCHITEW6432%" Then 
	ProgrammFiles=EnvironmentVariables("%ProgramFiles%")
	xOS = "x86"
End If
'C:\Program Files (x86)\7-Zip\7z.exe a -tzip D:\POST\MoveFilesPost.log.zip D:\POST\MoveFilesPost.log '

RunOutEx("""" & ProgrammFiles & "\7-Zip\7z.exe" & """" & " a -tzip D:\POST\MoveFilesPost.log.zip D:\POST\MoveFilesPost.log")

Function EnvironmentVariables(fvar)
  Set WshShell = WScript.CreateObject("WScript.Shell")
  EnvironmentVariables=WshShell.ExpandEnvironmentStrings(fvar)
End Function

Function RunOutEx(cmd)
	Dim WshShell
	set WshShell = WScript.CreateObject("WScript.Shell")
'	msgbox(cmd)
	WshShell.Run cmd,0
End Function

'ТЗ:
'1 - переименоать лог в дату
'2 - добавить в архив лог
'2а- проверить существование архива, если нет - создать, если есть - добавить файл
