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

Dim ver, LocalShare, OSVersion, OSArch, CompName, sProductType
Public PathPost, PathMail, iWriteLog, iCheckPath, iPing, FSO
iWriteLog=1
iCheckPath=0
ver="0.0.1"
LocalShare = WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)

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

Dim WshShell, OsVer
Set WshShell = CreateObject("WScript.Shell")
OsVer = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
msgbox OsVer

Dim objSWbemServices, colSWbemObjectSet', colSWbemObject
strComputer = "."
Set objSWbemServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("Win32_OperatingSystem")
'For Each objSWbemObject In colSWbemObjectSet
    'Wscript.Echo "Object Path: " & objSWbemObject
'Next
Set FSO = CreateObject("Scripting.FileSystemObject")


iPing = Ping(strComputerName)

If iPing=0 Then
	WriteLog("Интернет есть")
		If CheckPath(strShare) = 1 Then
			WriteLog("Каталог существует  " & strShare)
			''FSO.CopyFile DirDBF&NameDBF, DirSKM&NameDBF
		End If
Else
	WriteLog("Интернета НЕТ, код ошибки: "& iPing)
	WriteLog("Копьютер Офис-Менеджера ВЫКЛЮЧЕН? Код: " & iPing)
End If

Sub WriteLog(sData)
  Dim FSOL, FileLog, PathFileLog, ScriptFullName, ScriptName
  Set FSOL = CreateObject("Scripting.FileSystemObject")
  Set ScriptFullName = FSOL.GetFile(Wscript.ScriptFullName)
  ScriptName = FSOL.GetFileName(ScriptFullName)
  
  'msgbox ScriptName
  
  PathFileLog = "\" & ScriptName & ".log"
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

Function Ping (strTarget)
	Dim objWMIService, colPings, objPing
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strTarget & "'")
	For Each objPing in colPings
		Ping = objPing.StatusCode
	Next
End Function

Function CheckPath(Path)
  If FSO.FolderExists(Path) Then
    CheckPath=1
  End If
  If Not FSO.FolderExists(Path) Then
    CheckPath=0
    'FSO.Createfolder Path
    'iCheckPath="1"
  End If
  'CheckPath=iCheckPath
End Function

WScript.Quit 0