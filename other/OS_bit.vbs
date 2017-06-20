Option Explicit

Dim strShare
Dim strComputer
Dim strComputerName

Dim objSWbemLocator
Dim objSWbemServicesEx
Dim collSWbemObjectSet
Dim objSWbemObjectEx

Dim strCommandLine
Dim lngProcessID

Dim MSUpdate(3,3)
MSUpdate(1,1) = "KB4012212"
MSUpdate(2,1) = "7"
MSUpdate(3,1) = 0

MSUpdate(1,2) = "KB4012215"
MSUpdate(2,2) = "7"
MSUpdate(3,2) = 0

MSUpdate(1,3) = "KB4012598"
MSUpdate(2,3) = "XP"
MSUpdate(3,3) = 0


Dim ver, LocalShare, OSVersion, OSArch, CompName, sProductType, WshShell, OsVer
Public PathPost, PathMail, iWriteLog, iCheckPath, iPing, PathFileLog, PathFullFileLog, ScriptFullName, ScriptName, FSO, FSOL
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FSOL = CreateObject("Scripting.FileSystemObject")
Set ScriptFullName = FSOL.GetFile(Wscript.ScriptFullName)
ScriptName = FSOL.GetFileName(ScriptFullName)
PathFileLog = "\" & ScriptName & ".log"
LocalShare = WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
PathFullFileLog = LocalShare & PathFileLog

iWriteLog=1
iCheckPath=0
ver="0.0.1"

WriteLog("                    ")
WriteLog("                    ")
WriteLog("                    ")
WriteLog("Start:              " & Time & " " & Right(0 & Day(date), 2) & "." & Right(0 & Month(date), 2) & "." & Year(Date))
WriteLog("Version:            " & ver)
'WriteLog("OS:                 " & OSVersion & " " & OSArch)


'msgbox LocalShare

strComputerName = "Magistr"
strShare    = "\\" & strComputerName & "\netlogon"
strComputer = "."
iPing = Ping(strComputerName)

Set objSWbemLocator    = WScript.CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServicesEx = objSWbemLocator.ConnectServer(strComputer, "root\cimv2")
Set collSWbemObjectSet = objSWbemServicesEx.InstancesOf("Win32_OperatingSystem")

For Each objSWbemObjectEx In collSWbemObjectSet
    With objSWbemObjectEx
		'msgbox .Version
		'msgbox "123321" & .Caption
		CompName = .CSName
		
		Select Case .ProductType
            Case "1"
				sProductType = "Work Station"
            Case "2"
				sProductType = "Domain Controller"
            Case "3"
				sProductType = "Server"
		End Select

        Select Case .Version
            Case "5.1.2600"             ' Windows XP X86
                'strCommandLine = strShare & "\VPN_XP_32.exe"
				OSVersion = "XP"
				OSArch = "x32"
            Case "5.2.3790"             ' Windows XP X64 (& Windows Server 2003)
                'strCommandLine = strShare & "\VPN_XP_64.exe"
				OSVersion = "XP"
				OSArch = "x64"
            Case "6.1.7600", "6.1.7601" ' Windows 7 (& Windows Server 2008 R2)
				'msgbox .OSArchitecture
                OSVersion = "7"
				Select Case .OSArchitecture
                    Case "32-bit"       ' X86
                        'strCommandLine = strShare & "\VPN_7_32.exe"
						OSArch = "x32"
                    Case "64-bit"       ' X64
                        'strCommandLine = strShare & "\VPN_7_64.exe"
						OSArch = "x64"
                End Select
            Case Else
                ' Nothing to do
        End Select
    End With
Next

WriteLog("OS:                 " & OSVersion & " " & OSArch)
WriteLog("Type:               " & sProductType)
WriteLog("Computer Name:      " & CompName)

'-------------------------------------------'
If Not IsEmpty(strCommandLine) Then
    If objSWbemServicesEx.Get("Win32_Process").Create("""" & strCommandLine & """", Empty, Nothing, lngProcessID) = 0 Then
        WScript.Echo "Successfully execute [" & strCommandLine & "] on [" & strComputer & "]"
        WScript.Echo "Process Id: [" & lngProcessID & "]"
    Else
        WScript.Echo "Can't execute [" & strCommandLine & "] on [" & strComputer & "]"
    End If
End If

Set collSWbemObjectSet = Nothing
Set objSWbemServicesEx = Nothing
Set objSWbemLocator    = Nothing
'-------------------------------------------'

'Set FSO = CreateObject("Scripting.FileSystemObject")
'Set F = FSO.GetFile(Wscript.ScriptFullName)

'path = FSO.GetParentFolderName(F)

Set WshShell = CreateObject("WScript.Shell")
OsVer = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
'msgbox OsVer

Dim objSWbemServices, colSWbemObjectSet', colSWbemObject
strComputer = "."
Set objSWbemServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("Win32_OperatingSystem")
'For Each objSWbemObject In colSWbemObjectSet
    'Wscript.Echo "Object Path: " & objSWbemObject
'Next

Dim objWMIService, colComputer, objComputer
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colComputer = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
'Set WshShell = CreateObject( "Wscript.Shell")
For Each objComputer in colComputer
    WriteLog("User:               " & objComputer.UserName)
Next

Dim colProcesses, objProcess, strNameOfUser, Return
Set colProcesses = GetObject("winmgmts:" & _
   "{impersonationLevel=impersonate}!\\" & strComputer & _
   "\root\cimv2").ExecQuery("Select * from Win32_Process")

For Each objProcess in colProcesses
    Return = objProcess.GetOwner(strNameOfUser)
    If Return <> 0 Then
        WriteLog "        Could not get owner info for process " & _  
            objProcess.Name & VBNewLine _
            & "Error = " & Return
    Else 
        WriteLog "        Process " _
            & objProcess.Name & " is owned by " _ 
            & "\" & strNameOfUser & "."
    End If
Next

If iPing=0 Then
	WriteLog("Интернет есть")
		If CheckPathFolder(strShare) = 1 Then
			WriteLog("Каталог существует  " & strShare)
			'msgbox PathFullFileLog
			CheckMSUpdates
			'InstallMSUpdates
			'CheckMSUpdates
			If CheckPathFile(PathFullFileLog) = 1 Then
				FSO.CopyFile PathFullFileLog, strShare & "\Logs\" & ScriptName & "_" & CompName & time & " " & date & ".log"
			End If
			If CheckPathFile(PathFullFileLog & ".evt") = 1 Then
				FSO.CopyFile PathFullFileLog & ".evt", strShare & "\Logs\" & ScriptName & "_" & CompName & time & " " & date &  ".evt"
			End If
		End If
Else
	WriteLog("Интернета НЕТ, код ошибки: "& iPing)
	WriteLog("Копьютер Офис-Менеджера ВЫКЛЮЧЕН? Код: " & iPing)
End If

Sub WriteLog(sData)
  Dim FileLog', FSOL, PathFileLog, ScriptFullName, ScriptName
  'Set FSOL = CreateObject("Scripting.FileSystemObject")
  'Set ScriptFullName = FSOL.GetFile(Wscript.ScriptFullName)
  'ScriptName = FSOL.GetFileName(ScriptFullName)
  'PathFileLog = "\" & ScriptName & ".log"
  
  'msgbox PathFullFileLog
  
  If FSOL.FileExists(PathFullFileLog) Then
    Set FileLog=FSOL.OpenTextFile(PathFullFileLog, 8)
  End If

  If Not FSOL.FileExists(PathFullFileLog) Then
    SET FileLog=FSOL.CreateTextFile(PathFullFileLog, True)
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

Function Ping (strTarget)
	Dim objWMIService, colPings, objPing
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strTarget & "'")
	For Each objPing in colPings
		Ping = objPing.StatusCode
	Next
End Function

Function CheckPathFile(Path)
	If FSO.FileExists(Path) Then
		CheckPathFile=1
		WriteLog(Path & " - Файл " & CheckPathFile)
	End If
	If Not FSO.FileExists(Path) Then
		CheckPathFile=0
		WriteLog(Path & " - Файл " & CheckPathFile)
		'FSO.Createfolder Path
		'iCheckPath="1"
	End If
End Function

Function CheckPathFolder(Path)
	If FSO.FolderExists(Path) Then
		CheckPathFolder=1
		WriteLog(Path & " - Директория " & CheckPathFolder)
	End If
	If Not FSO.FolderExists(Path) Then
		CheckPathFolder=0
		WriteLog(Path & " - Директория " & CheckPathFolder)
		'FSO.Createfolder Path
		'iCheckPath="1"
	End If
	'CheckPathFolder=iCheckPath
End Function

Sub CheckMSUpdates()
	'Dim strComputer
	Dim objWmiService
	Dim wmiNS
	Dim wmiQuery
	Dim objItem
	Dim colItems
	Dim objSWbemObjectEx
	'Dim MSUpdate(3,3)
	Dim i, f
	Dim FoundMSUpdate
	
	FoundMSUpdate = 0
	
	WriteLog("Start search Updates")

	wmiNS = "\root\cimv2"
	wmiQuery = "Select * from Win32_QuickFixEngineering"
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS)
	Set colItems = objWMIService.ExecQuery(wmiQuery)
	 
'	WriteLog("End search Updates")
	For i = 1 To UBound(MSUpdate,1)  Step 1
		If MSUpdate(2, i) = OSVersion Then
			WriteLog(MSUpdate(1, i) & " " & MSUpdate(2, i) & " " & MSUpdate(3, i))
			For Each objItem in colItems
			'    Wscript.Echo "Caption: " & objItem.Caption
				'Wscript.Echo "CSName: " & objItem.CSName
				'Wscript.Echo "Description: " & objItem.Description
				'Wscript.Echo "FixComments: " & objItem.FixComments
				'Wscript.Echo "HotFixID: " & objItem.HotFixID
			'    Wscript.Echo "InstallDate: " & objItem.InstallDate
			'    Wscript.Echo "InstalledBy: " & objItem.InstalledBy
			'    Wscript.Echo "InstalledOn: " & objItem.InstalledOn
				'Wscript.Echo "Name: " & objItem.Name
			'    Wscript.Echo "ServicePackInEffect: " & objItem.ServicePackInEffect
			'    Wscript.Echo "Status: " & objItem.Status
				If MSUpdate(3, i) = 1 Then
					If MSUpdate(1, i) = objItem.HotFixID Then
						FoundMSUpdate = 1
						WriteLog("Обновление: " & MSUpdate(1, i) & " " & MSUpdate(3, i) & " уже было ранее успешно установлено: " & objItem.InstalledOn)
					End If
				Else
					If MSUpdate(1, i) = objItem.HotFixID Then
						MSUpdate(3, i) = 1
						FoundMSUpdate = 1
						WriteLog("Обновление: " & MSUpdate(1, i) & " " & MSUpdate(3, i) & " уже было ранее установлено: " & objItem.InstalledOn)
					End If
				End If
			Next
			
			If MSUpdate(3, i) = 1 Then
				If FoundMSUpdate = 0 Then
					MSUpdate(3, i) = 0
					WriteLog("Обновление: " & MSUpdate(1, i) & " " & MSUpdate(3, i) & " не было установлено корректно")
				End If
			Else
				WriteLog("Обновление: " & MSUpdate(1, i) & " " & MSUpdate(3, i) & " ещё не установлено")
			End If
		End If
		FoundMSUpdate = 0
	Next
	WriteLog("End search Updates")
End Sub

Sub InstallMSUpdates()
	WriteLog("Start upgrade")
	Dim i, CodeInstallMSUpdate, PatchMSUpdate, FullPatchMSUpdate
	For i = 1 To UBound(MSUpdate,1)  Step 1
		If MSUpdate(2, i) = OSVersion Then
			If OSVersion = "7" Then
				PatchMSUpdate = strShare & "\KB\" & OSVersion & "-" & MSUpdate(1, i) & "-" & OSArch & ".msu"
				'WriteLog(PatchMSUpdate)
			End If
			If OSVersion = "XP" Then
				PatchMSUpdate = strShare & "\KB\" & OSVersion & "-" & MSUpdate(1, i) & "-" & OSArch & ".exe"
				'WriteLog(PatchMSUpdate)
			End If
			If MSUpdate(3, i) = 0 Then
				If CheckPathFile(PatchMSUpdate) = 1 Then
					FullPatchMSUpdate = PatchMSUpdate & " /quiet /norestart /log:" & PathFullFileLog & ".evt"
					WriteLog(FullPatchMSUpdate)
					CodeInstallMSUpdate = WshShell.Run(FullPatchMSUpdate, 1, True)
					WriteLog ("Обработка завершена! Код возврата - " & CodeInstallMSUpdate)
				Else
					WriteLog ("Файл обновлений - не найден")
				End If
			End If
		End If
	Next
	WriteLog("Finish upgrade")
End Sub

WScript.Quit 0