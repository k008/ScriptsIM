'Dim T1
'T1=time
Option explicit
Dim ServerShare, LocalShare, ver, ServerMaptekaIn, ServerMaptekaFarm, PathPostA2, PathMailA2, PathPostSKM1, PathMailSKM1, PathPostSKM2, PathMailSKM2, PathPostSKM3, PathMailSKM3, PathPostPIR, PathMailPIR
Public PathPost, PathMail, iWriteLog, iCheckPath, NameFileLog, WshShell

iWriteLog=1
iCheckPath=0
ver="0.1.5" 'ZAMENA UL

'User = 
'Password = 
ServerShare = "P:\" '"\\Violeta\Install_Violeta\" '"Z:\" '"\\10.20.30.176\ADV\POST_ADV\"
'ServerMaptekaIn = "X:\programs\in\"
'ServerMaptekaFarm = "X:\farm\office\"
LocalShare = "D:\POST\"

PathPostA2="\\10.20.30.1\POST_Office"
'PathPostA2="\\10.20.30.1\POST_A1\A2"
'PathMailA2="\\PRIMERGY\Mail"

PathPostSKM1="\\10.20.30.1\POST_Office"
'PathPostG1="\\10.20.30.1\POST_A1\G1"
'PathMailG1="\\172.27.3.1\Mail_Prima"

PathPostSKM2="\\10.20.30.1\POST_Office"
'PathPostG2="\\10.20.30.1\POST_A1\G2"
'PathMailG2="C:\Mail"

PathPostSKM3="\\10.20.30.1\POST_Office"
'PathPostADV="\\10.20.30.1\POST_A1\ADV"
'PathMailADV="\\172.27.5.1\Mail"

PathPostPIR="\\10.20.30.1\POST_Office"
'PathPostPIR="\\10.20.30.1\POST_A1\PIR"
'PathMailPIR="C:\Mail"

WriteLog("                    ")
WriteLog("                    ")
WriteLog("                    ")
WriteLog("Start:              " & Time & " " & Right(0 & Day(date), 2) & "." & Right(0 & Month(date), 2) & "." & Year(Date))
WriteLog("Version:            " & ver)
'WriteLog("OS:                 " & @OSVersion & " " & @OSArch)
'WriteLog("User:               " & Who)

'server->заведующей
'MoveAllFiles ServerShare & "out\", LocalShare & "in\"
'заведующая-сервер
'MoveAllFiles LocalShare & "in\", ServerMaptekaIn
'сервер->заведующая
'MoveAllFiles ServerMaptekaFarm, LocalShare & "out\"
'заведующей->server
'MoveAllFiles LocalShare & "out\", ServerShare & "in\"

'Главная Ф-ИЯ
CheckIP
WriteLog("Exit")
CheckLogSize
WriteLog("Exit.")


Sub MoveAllFiles(FDir,OutPath)
Dim FSO,FLD,FL,FF, FDirStatus, OutPathStatus
Set FSO = CreateObject("Scripting.FileSystemObject")
WriteLog("                    ") 
WriteLog("In Path: " & FDir)
WriteLog("Out Path: " & OutPath)
FDirStatus=CheckPath(FDir)
OutPathStatus=CheckPath(OutPath)
WriteLog("FDirStatus=" & FDirStatus)
WriteLog("OutPathStatus=" & OutPathStatus)
	If FDirStatus And OutPathStatus Then
		WriteLog("FDir=" & FDir & " OutPath=" & OutPath)
		WriteLog("Folders exists_moveallfiles")
'      msgbox("1" & chr(10) & FDir & chr(10) & OutPath)
		Set FLD = FSO.GetFolder(FDir)
		Set FL = FLD.Files
		For Each FF in FL
		WriteLog("In File: " & FDir&FF.Name)
		WriteLog("Out File: " & OutPath&FF.Name)
'      msgbox(FDir&FF.Name & chr(10) & OutPath&FF.Name)
		FSO.CopyFile FDir&FF.Name, OutPath&FF.Name
		FSO.DeleteFile FDir&FF.Name
		Next
	Else
		WriteLog("Folders NOT exists")
	End If
Set FL = Nothing
Set FLD = Nothing
Set FSO = Nothing
End Sub

Sub CopyAllFiles(FDir,OutPath)
Dim FSO,FLD,FL,FF
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files
For Each FF in FL
	FSO.CopyFile FDir&FF.Name, OutPath&FF.Name
Next
Set FL = Nothing
Set FLD = Nothing
Set FSO = Nothing
End Sub

Sub DeleteFiles(FDir)
Dim FSO,FLD,FL,FF
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files
For Each FF in FL
	FSO.DeleteFile FDir&FF.Name
Next
Set FL = Nothing
Set FLD = Nothing
Set FSO = Nothing
End Sub

Sub CheckDrive(strDriveName, strRemoteShare)
Dim FSO, objNetwork
Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objNetwork = WScript.CreateObject("WScript.Network")

If FSO.DriveExists(strDriveName) Then
	WriteLog("Диск " & strDriveName & " уже подключен")
	'Выводим информацию на экран
	'WScript.Echo "Диск " & strDriveName & " уже подключен"
Else
	'Подключаем диск y:
	WriteLog("Выполняется сетевое подключение сетевого диска: " & strDriveName)
	objNetwork.MapNetworkDrive strDriveName, strRemoteShare
	iCheckPath="1"
	WriteLog("Сетевой Диск: "& strDriveName & " Успешно подключен")
	'Вывод информации на экран:
	'WScript.Echo " Сетевой Диск: "& strDriveName & " Успешно подключен"
End If
End Sub

Sub CheckIP
Dim strComputer, strNetworkConnection, objWMIService, colNics, objNic, colNicConfigs, objNicConfig, strIPAddress, OpenVPNIP
WriteLog("Check IP...")
strComputer  =  "."
' отредактировать под нужное имя сетевого подключения:
strNetworkConnection = "'OpenVPN'"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colNics = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter " _
	& "Where NetConnectionID = " & strNetworkConnection)
For Each objNic in colNics
	Set colNicConfigs = objWMIService.ExecQuery("ASSOCIATORS OF " _
		& "{Win32_NetworkAdapter.DeviceID='" & objNic.DeviceID & "'}" _
		& " WHERE AssocClass=Win32_NetworkAdapterSetting")
	For Each objNicConfig In colNicConfigs
		For Each strIPAddress in objNicConfig.IPAddress
		'Wscript.Echo "IP Address: " & strIPAddress
		OpenVPNIP=strIPAddress
		WriteLog(strNetworkConnection & ": " & OpenVPNIP)
		Next
	Next
Next

Select Case OpenVPNIP
	Case "10.20.30.104"
		PathPost=PathPostA2
		'PathMail=PathMailA2
		WriteLog("Case: 10.20.30.104")
		MoveAllFiles ServerShare & "A2\out\", LocalShare & "in\"
		MoveAllFiles LocalShare & "out\", ServerShare & "A2\in\"
		
	Case "10.20.30.152"
		PathPost=PathPostSKM1
		'PathMail=PathMailG1
		WriteLog("Case: 10.20.30.52")
		'MoveAllFiles ServerShare & "out\", LocalShare & "in\"
		'MoveAllFiles LocalShare & "out\", ServerShare & "in\"
		'A1-G1
		MoveAllFiles ServerShare & "SKM1\out\", LocalShare & "in\"
		'G1-A1
		MoveAllFiles LocalShare & "out\", ServerShare & "SKM1\in\"
		'G1-A1-G2
		'MoveAllFiles LocalShare & "G2\", ServerShare & "G2\out\"
		
	Case "10.20.30.201"
		PathPost=PathPostSKM2
		'PathMail=PathMailG2
		WriteLog("Case: 10.20.30.201")
		'MoveAllFiles ServerShare & "out\", LocalShare & "in\"
		'MoveAllFiles LocalShare & "out\", ServerShare & "in\"
		'A1-G2
		MoveAllFiles ServerShare & "SKM2\out\", LocalShare & "in\"
		'G2-A1-G1
		'MoveAllFiles LocalShare & "out\", ServerShare & "G1\out\"
		'G2-A1
		MoveAllFiles LocalShare & "out\", ServerShare & "SKM2\in\"
	
	Case "10.20.30.170"
		PathPost=PathPostSKM
		'PathMail=PathMailADV
		WriteLog("Case: 10.20.30.170")
		MoveAllFiles ServerShare & "SKM3\out\", LocalShare & "in\"
		MoveAllFiles LocalShare & "out\", ServerShare & "SKM3\in\"
	
	Case "10.20.30.160"
		PathPost=PathPostPIR
		'PathMail=PathMailPIR
		WriteLog("Case: 10.20.30.160")
		MoveAllFiles ServerShare & "PIR\out\", LocalShare & "in\"
		MoveAllFiles LocalShare & "out\", ServerShare & "PIR\in\"
		
	Case Else
		WriteLog("Case: Else")
		If FileExist("D:\POST\keyVBS") = "1" Then
		WriteLog("ALEX")
		'PathPost="\\129.186.1.25\POST\A2"
		PathPost="\\129.186.1.25\POST_Office"
		'PathMail="C:\Mail"
		MoveAllFiles ServerShare & "ALEX\out\", LocalShare & "in\"
		MoveAllFiles LocalShare & "out\", ServerShare & "ALEX\in\"
		End If
	'завершить скрипт!
		'PathPost="Aquarius"
		'PathMail="Aquarius"
End Select
End Sub 

Sub WriteLog(sData)
Dim FSOL, FileLog, PathFileLog
Set FSOL = CreateObject("Scripting.FileSystemObject")
NameFileLog="MoveFilesPost.log"
PathFileLog="\" & NameFileLog

If FSOL.FileExists(LocalShare & PathFileLog) Then
	Set FileLog=FSOL.OpenTextFile(LocalShare & PathFileLog, 8)
End If

If Not FSOL.FileExists(LocalShare & PathFileLog) Then
	SET FileLog=FSOL.CreateTextFile(LocalShare & PathFileLog, True)
End If

	If iWriteLog = 1 Then
		If sData = "                    " Then
			FileLog.WriteLine("                    ")
			'wscript.Echo chr(10) & sData
		Else
			FileLog.WriteLine(Time & " " & Right(0 & Day(date), 2) & "." & Right(0 & Month(date), 2) & "." & Year(Date) & " " & sData)
			'WScript.Echo (HOUR(NOW) & ":" & MINUTE(NOW) & ":" & SECOND(NOW) & " " & DAY(NOW) & "/" & MONTH(NOW) & "/" & YEAR(NOW) & " " & sData)
		End If
	End If
	FileLog.Close
End Sub

Function Who
	Dim colComputer, objWMIService, strComputer
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Set colComputer = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
	For Each objComputer in colComputer
		'WriteLog(objComputer.UserName)
		Who=objComputer.UserName
	Next
End Function

Function CheckPath(Path)
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
WriteLog("CheckPath: " & Path & "=" & FSO.FolderExists(Path))
	If FSO.FolderExists(Path) Then
		WriteLog("Folders exists_checkpatch")
		iCheckPath="1"
	Else
		iCheckPath="0"
		WriteLog("Folder NOT exists")
			
		If Left(Path,2) = "P:" Then
		CheckDrive "P:", PathPost
		End If
		
		If Left(Path,2) = "X:" Then
		WriteLog("REM")
		'CheckDrive "X:", PathMail
		End If
		
		If Left(Path,2) = "D:" Then
		WriteLog("Folder: " & Path & " NOT Exist")
		'проверить структуру локальной папки
		End If
	End If
	CheckPath=iCheckPath
End Function

Function FileExist(fFile)
Dim FSOFE
Set FSOFE = CreateObject("Scripting.FileSystemObject")
		If FSOFE.FileExists(fFile) Then
		FileExist="1"
		Else
		FileExist="0"
		End If
End Function

Function CheckLogSize
Dim FSOL1, FileLog, PathFileLog, LogFile, PathScript, PathFolderScript, LogSize
Set FSOL1 = CreateObject("Scripting.FileSystemObject")
PathFileLog="\" & NameFileLog
If FSOL1.FileExists(LocalShare & PathFileLog) Then
	Set LogFile = FSOL1.GetFile(LocalShare & PathFileLog)
	LogSize = LogFile.Size
	WriteLog ("Размер файла " & WScript.ScriptName & " : " & LogSize &" килобайт")
	If LogSize >= 524288 Then
		WriteLog ("Размер Лога большой=" & LogSize)
		MoveLogFiles
	End if
End If
End Function

Function MoveLogFiles
    Dim FSO, newNameFileLog
    Set FSO = CreateObject("Scripting.FileSystemObject")
    newNameFileLog=LocalShare & NameFileLog & ".old"
    FSO.MoveFile LocalShare & NameFileLog, newNameFileLog
    RunOutEx("""" & ProgrammFiles & "\7-Zip\7z.exe""" & " a -tzip " & newNameFileLog & ".zip " & newNameFileLog)
End Function

Function EnvironmentVariables(fvar)
	Set WshShell = WScript.CreateObject("WScript.Shell")
	EnvironmentVariables=WshShell.ExpandEnvironmentStrings(fvar)
End Function

Function RunOutEx(cmd)
	set WshShell = WScript.CreateObject("WScript.Shell")
	'msgbox(cmd)
	WshShell.Run cmd
End Function

Function ProgrammFiles
	Dim WshProEnv, SysInfo
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set WshProEnv = WshShell.Environment("SYSTEM") 
	SysInfo = WshProEnv.Item("PROCESSOR_ARCHITECTURE")
	If SysInfo = "x86" Then
		ProgrammFiles=EnvironmentVariables("%ProgramFiles%")
	Else
		ProgrammFiles=EnvironmentVariables("%ProgramW6432%")
	End If
End Function