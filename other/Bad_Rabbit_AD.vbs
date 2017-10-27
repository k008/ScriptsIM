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

Public ver, LocalShare, sTime, sDate, iWriteLog, PathFileLog, PathFullFileLog, iPing, CompName, sProductType, OSVersion, OSArch, OsVer
Public ScriptFullName, ScriptName
Dim DesktopPath, FSOL, WshShell, FSO
Dim File1, File2, User, UserAccess, attrib
ReDim ConvTable(1)
Const TF="128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,240,241,242,243,244,245,246,247,248,249,250,251,252,253,254,255"
Const TT="192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,63,63,63,166,63,63,63,63,63,63,63,63,63,63,63,172,63,63,63,63,63,134,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,135,63,63,63,63,63,63,63,240,241,242,243,244,245,246,247,248,249,250,251,252,253,254,255,168,184,170,186,175,191,161,162,176,149,183,63,185,164,152,160"

iWriteLog=1
Set FSOL = CreateObject("Scripting.FileSystemObject")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ScriptFullName = FSOL.GetFile(Wscript.ScriptFullName)
Set WshShell = CreateObject("WScript.Shell")
DesktopPath = WSHShell.SpecialFolders("Desktop")
'LocalShare = DesktopPath 'WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
ScriptName = FSOL.GetFileName(ScriptFullName)
'PathFileLog = "\" & ScriptName & ".log"
'PathFullFileLog = LocalShare & PathFileLog

PathFileLog = "\" & ScriptName & ".log"
LocalShare = WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
PathFullFileLog = LocalShare & PathFileLog

ver="0.0.1"
sTime = Hour(time) & Minute(time) & Second(time)
sDate = Day(date) & Month(date) & Year(date)

strComputerName = "magistr.mirlekarstw.ru"
strShare    = "\\" & strComputerName & "\netlogon"
strComputer = "."
iPing = Ping(strComputerName)

WriteLog("                    ")
WriteLog("                    ")
WriteLog("                    ")
WriteLog("Start:              " & Time & " " & Right(0 & Day(date), 2) & "." & Right(0 & Month(date), 2) & "." & Year(Date))
WriteLog("Version:            " & ver)
'WriteLog("OS:                 " & OSVersion & " " & OSArch)

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
OsVer = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")

Dim objSWbemServices, colSWbemObjectSet', colSWbemObject
strComputer = "."
Set objSWbemServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("Win32_OperatingSystem")
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



File1="C:\Windows\infpub.dat"
'File="C:\Users\с_барбаянов\Desktop\1 2 3\1"
File2="C:\Windows\cscc.dat"
User="Все"
UserAccess="N" 'N  - нет, R  - чтение, W  - запись, C  - изменение (запись), F  - полный доступ 
attrib=" /t /e /p " & User & ":" & UserAccess

CreateFile(File1)
CreateFile(File2)

fCaclsRead(File1)
fCaclsRead(File2)

fCaclsChange(File1)
fCaclsChange(File2)

If iPing=0 Then
	WriteLog("Интернет есть")
	If CheckPathFolder(strShare) = 1 Then
		WriteLog("Каталог существует  " & strShare)
		'msgbox PathFullFileLog
		'CheckMSUpdates
		'InstallMSUpdates
		'CheckMSUpdates
		If CheckPathFile(PathFullFileLog) = 1 Then
			FSO.CopyFile PathFullFileLog, strShare & "\Logs\" & ScriptName & "_" & CompName & "_" & sTime & "_" & sDate & ".log"
		End If
		If CheckPathFile(PathFullFileLog & sTime & "_" & sDate & ".evt") = 1 Then
			FSO.CopyFile PathFullFileLog & sTime & "_" & sDate & ".evt", strShare & "\Logs\" & ScriptName & "_" & CompName & "_" & sTime & "_" & sDate &  ".evt"
		End If
	End If
Else
	WriteLog("Интернета НЕТ, код ошибки: "& iPing)
	WriteLog("Копьютер Офис-Менеджера ВЫКЛЮЧЕН? Код: " & iPing)
End If

WriteLog("END.")

Sub WriteLog(sData)
  Dim FileLog
  
  If FSOL.FileExists(PathFullFileLog) Then
    Set FileLog=FSOL.OpenTextFile(PathFullFileLog, 8)
  End If

  If Not FSOL.FileExists(PathFullFileLog) Then
    SET FileLog=FSOL.CreateTextFile(PathFullFileLog, True)
  End If
  
	If iWriteLog = 1 Then
		If sData = "                    " Then
			FileLog.WriteLine("                    ")
		Else
			FileLog.WriteLine(Time & " " & Right(0 & Day(date), 2) & "." & Right(0 & Month(date), 2) & "." & Year(Date) & " " & sData)
			'WScript.Echo (HOUR(NOW) & ":" & MINUTE(NOW) & ":" & SECOND(NOW) & " " & DAY(NOW) & "/" & MONTH(NOW) & "/" & YEAR(NOW) & " " & sData)
		End If
	End If
    FileLog.Close
End Sub

Function RunDos(sCommand)
	Dim Result, objShell, objExec
	'msgbox "cmd /C " & sCommand & " > " & chr(34) & result & chr(34)
	Set Result = WshShell.Exec("cmd /C " & sCommand)
	Do While Result.Status = 0
		WScript.Sleep 100
	Loop
	Set objShell = Nothing
	Set objExec = Nothing
	RunDos = Convert866to1251(Result.StdOut.ReadAll)
End Function

Sub MakeConvTable()
  Dim ArrT,ArrF,i
  ReDim ConvTable(256)
  ArrF=Split(TF,",")
  ArrT=Split(TT,",")
  For i=0 to UBound(ArrF)
    ConvTable(ArrF(i))=Chr(ArrT(i))
  Next
End Sub

Function Convert866to1251(St)
  Dim A,i,Ch, StOut
  StOut=""
  if UBound(ConvTable)=1 then MakeConvTable()
  For i=1 to Len(St)
	Ch=Mid(St,i,1)   
	A=ConvTable(Asc(Ch))
	if A="" then A=Ch
	StOut=StOut&A
  Next 
  Convert866to1251=StOut
End Function

Function arrDisplay(arraydisp)
	Dim i
	For i=0 to UBound(arraydisp) - 1
		If arraydisp(i) <> "" Then
			WriteLog(arraydisp(i))
		End If
	Next
End Function

Function CreateFile(FileToCreate)
	If Not FSOL.FileExists(FileToCreate) Then
		FSOL.CreateTextFile FileToCreate, True
		Writelog ("File Create:" & FileToCreate)
	Else
		Writelog ("File: " & FileToCreate & " Exist")
	End If
End Function

Function fCaclsRead(File)
	Dim CommandLineReadAccess, Users, arrUsers, i, boolUserAccess
	CommandLineReadAccess="cacls " & """" & File & """" & " /t"
	RunDos(CommandLineReadAccess)
	WriteLog("CommandLineReadAccess: " & CommandLineReadAccess)
	Users = Replace(RunDos (CommandLineReadAccess), File,"")
	arrUsers=split(Users, chr(13))

	For i=0 To UBound(arrUsers) - 1
		arrUsers(i)=Replace(arrUsers(i), chr(32) & chr(32), " ")
		arrUsers(i)=Replace(arrUsers(i), chr(10), "")
		arrUsers(i)=Replace(arrUsers(i), chr(13), "")
	Next

	For i=0 to UBound(arrUsers)-1
		If instr(arrUsers(i), User) then
			WriteLog("Пользователь найден")
			If instr(arrUsers(i), User & ":" & UserAccess) then
				'msgbox "Ok"
				'WriteLog("Права успешно применились")
				boolUserAccess  = "1"
			else
				'msgbox "Bad"
				'WriteLog("Права не установились")
				 boolUserAccess = "0"
			end if
		end if
	Next
	arrDisplay(arrUsers)
	fCaclsRead = boolUserAccess
End Function

Function fCaclsChange(File)
	Dim CommandLineChangeAccess, boolUserAccess
	CommandLineChangeAccess="cacls " & """" & File & """" & attrib
	WriteLog("CommandLineChangeAccess: " & CommandLineChangeAccess)
	RunDos (CommandLineChangeAccess)
	boolUserAccess = fCaclsRead (File)

	If boolUserAccess = 0 Then
		WriteLog("Права на файл: " & File & " не установились: " & boolUserAccess)
		'msgbox "Права на файл: " & File & " не установились"
	End If

	If boolUserAccess = 1 Then
		WriteLog("Права на файл: " & File & " успешно применились: " & boolUserAccess)
		'msgbox "Права на файл: " & File & " успешно применились"
	End If
End Function

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