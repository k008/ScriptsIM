Option explicit
Dim Dir, PathFullFileLog, PathFileLog, ScriptName, FSO, iWriteLog, WshShell
Dim ver
Dim stribases, str1CEStart, str1CEStartexists, stribasesexists
Dim Read1cefilesplit, Readv8ifilesplit
Dim bom, bomreadv8i, bomread1ce
Dim CompName, sProductType, OSVersion, OSArch, OsVer
Dim ibases(199,4), ibasesi
Dim CEStart(199,1), CEStarti
Dim arCEStart() 'массив записи в файл
Dim strComputerName, strShare, iPing
Set FSO = CreateObject("Scripting.FileSystemObject")

ScriptName = "1C-links"
ver = "0.1.3" ' Отключено создание ссылок
PathFileLog = ScriptName & ".log"
Dir = EnvironmentVariables("%TEMP%") & "\"
PathFullFileLog = Dir & PathFileLog
iWriteLog = 1
stribases = "1C\1CEStart\ibases.v8i"
str1CEStart = "1C\1CEStart\1CEStart.cfg"
str1CEStartexists=0
stribasesexists=0


strComputerName = "192.168.19.3"
strShare    = "\\" & strComputerName & "\links"


WriteLog("                    ")
WriteLog("                    ")
WriteLog("                    ")
WriteLog("Start:              " & Time & " " & Right(0 & Day(date), 2) & "." & Right(0 & Month(date), 2) & "." & Year(Date))
WriteLog("Version:            " & ver)

Call Main

'Основная процедура программы
Sub Main
	Call InfoOS
	
	stribasesexists = FileExist (EnvironmentVariables("%APPDATA%") & "\" & stribases)
	If stribasesexists = 1 Then
		'Call Readv8i
		'call ReadFile(EnvironmentVariables("%APPDATA%") & "\" & stribases)
		'call ReadFile(EnvironmentVariables("%APPDATA%") & "\" & str1CEStart)
		CopyFiles EnvironmentVariables("%APPDATA%") & "\" & stribases
		call Readv8i
	End If
	
	str1CEStartexists = FileExist (EnvironmentVariables("%APPDATA%") & "\" & str1CEStart)
	If str1CEStartexists = 1 Then
		CopyFiles EnvironmentVariables("%APPDATA%") & "\" & str1CEStart
		call Read1ce
	End If
	'Call TestServerLinks
	If stribasesexists = 1 And str1CEStartexists Then
		call SravnenieBases
	Else
		WriteLog("Сравнение баз не выполнено")
	End If
	msgbox "Настройка 1С закончена"
End Sub


'Запись лога, iWriteLog = 1 - запись разрешена
Sub WriteLog(sData)
	Dim FileLog, FSOL
	Set FSOL = CreateObject("Scripting.FileSystemObject")
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

'Переменные среды
Function EnvironmentVariables(fvar)
	Set WshShell = WScript.CreateObject("WScript.Shell")
	EnvironmentVariables=WshShell.ExpandEnvironmentStrings(fvar)
End Function

'Проверка файла на существование
Function FileExist(fFile)
	Dim FSOFE
	Set FSOFE = CreateObject("Scripting.FileSystemObject")
	
	If FSOFE.FileExists(fFile) Then
		FileExist="1"
	Else
		FileExist="0"
	End If
End Function

'Проверка директории на существование
Function FolderExist(Path)
	If FSO.FolderExists(Path) Then
		FolderExist=1
		WriteLog(Path & " - Директория " & FolderExist)
	End If
	If Not FSO.FolderExists(Path) Then
		FolderExist=0
		WriteLog(Path & " - Директория " & FolderExist)
		'FSO.Createfolder Path
		'iCheckPath="1"
	End If
	'CheckPathFolder=iCheckPath
End Function

'Создание директории
Function CreateFolder(FolderToCreate)
	If Not FSO.FolderExists(FolderToCreate) Then
		FSO.CreateFolder  FolderToCreate
		Writelog ("Folder Create:" & FolderToCreate)
	Else
		Writelog ("Folder: " & FolderToCreate & " Exist")
	End If
End Function

'Создание файла
Function CreateFile(FileToCreate)
	If Not FSO.FileExists(FileToCreate) Then
		Writelog ("File Create:" & FileToCreate)
		FSO.CreateTextFile FileToCreate, True
	Else
		Writelog ("File: " & FileToCreate & " Exist")
	End If
End Function

'del
Sub Readv8i1()
	Dim Openv8i, Lenstrv8i, instrsrvtitle, instrsrv, instrsrvsep, str, TristateTrue
	TristateTrue="-2"
	Set Openv8i = FSO.OpenTextFile(EnvironmentVariables("%APPDATA%") & "\" & stribases, 1, False)

	Do While Not Openv8i.AtEndOfStream
		'If Openv8i
		str=Openv8i.ReadLine
		Lenstrv8i = Len(str)
		'msgbox Lenstrv8i
		msgbox mid(str,1,5)
		'msgbox Len(str)
		instrsrvtitle=instr(1,str,"[")
		'msgbox instrsrvtitle
		'msgbox Left(str,instrsrvtitle)
		'msgbox Left(str,Lenstrv8i)
		If instrsrvtitle > 0 Then 
			If mid(str,instrsrvtitle,1) = "[" and Right(str,1) = "]" Then
				msgbox mid(str, instrsrvtitle, Lenstrv8i)
			'перейти на другую строку'
			Else 
			'msgbox str
				instrsrv = instr(1, str, "Connect=Srvr=")
				If instrsrv > 0 Then
					instrsrvsep=instr(1, str, ";")
					msgbox instrsrvsep
					msgbox mid(str, instrsrv+13, instrsrvsep-14)
					instrsrvsep=instr(instrsrvsep+1, str, ";")
					msgbox instrsrvsep
					msgbox mid(str, instrsrv+13, instrsrvsep-14)
				End If
				'msgbox instrsrv
			End If
		End If
			
	Loop
	Openv8i.close

End Sub

'Чтение ibases.v8i, ibases (0,4): 00-имя раздела, 01-Connect=Srvr= (адрес), 02-ref (имя базы), 03-External (внешяя база/ссылка), 04-del-удалить базу из ibases.v8i, add-добавить базу, 1CEStart.cfg, added-база добавлена
Sub Readv8i()
	Dim Lenstrv8i, instrsrvtitle, instrsrv, instrsrvsep, instrsrvbase, instrsrvsep2, instrsrvextexternal, instrsrvsep3, i', readfilesplit, bomreadv8i
	Readv8ifilesplit = split(ReadFile(EnvironmentVariables("%APPDATA%") & "\" & stribases), vbcrlf)
	bomreadv8i = bom
	
	'Чтение файла
	ibasesi=-1
	For i=0 to ubound(Readv8ifilesplit)
		Lenstrv8i = Len(Readv8ifilesplit(i))
		instrsrvtitle=instr(1,Readv8ifilesplit(i),"[")
		'Если есть [, значит это начало раздела базы
		If instrsrvtitle > 0 Then
		'Если есть [ и ], значит это раздел базы
			If mid(Readv8ifilesplit(i),instrsrvtitle,1) = "[" and Right(Readv8ifilesplit(i),1) = "]" Then
				Writelog ""
				'msgbox mid(Readv8ifilesplit(i), instrsrvtitle, Lenstrv8i)
				'WriteLog mid(Readv8ifilesplit(i), instrsrvtitle, Lenstrv8i)
				ibasesi = ibasesi+1
				ibases(ibasesi,0) = mid(Readv8ifilesplit(i), instrsrvtitle, Lenstrv8i)
				WriteLog ibases(ibasesi,0) & " ibasesi=" & ibasesi
			End If
			
			'Нет [ и ], значит это содержимое раздела базы
		Else 
			'msgbox Readv8ifilesplit(i)
			instrsrv = instr(1, Readv8ifilesplit(i), "Connect=Srvr=")
			If instrsrv > 0 Then
				instrsrvsep=instr(instrsrv, Readv8ifilesplit(i), ";")
				'msgbox mid(Readv8ifilesplit(i), instrsrv+13, instrsrvsep-instrsrv-13)
				'WriteLog mid(Readv8ifilesplit(i), instrsrv+13, instrsrvsep-instrsrv-13)
				ibases(ibasesi,1) = mid(Readv8ifilesplit(i), instrsrv+13, instrsrvsep-instrsrv-13)
				WriteLog "Connect=Srvr=" & ibases(ibasesi,1) & " ibasesi=" & ibasesi
			End If
			
			instrsrvbase = instr(1, Readv8ifilesplit(i), "Ref=")
			If instrsrvbase > 0 Then
				instrsrvsep2=instr(instrsrvbase, Readv8ifilesplit(i), ";")
				'От = и до конца строки, если последний ;-убрать, если нет "- то добавить
				'msgbox instrsrvsep2
				'WriteLog Readv8ifilesplit(i) & " " & instrsrvbase+4 & " " & instrsrvsep2-instrsrvbase-4
				'ibases(ibasesi,2) = mid(Readv8ifilesplit(i), instrsrvbase+4, instrsrvsep2-instrsrvbase-4)
				'msgbox "string=" & Readv8ifilesplit(i) & vbcrlf & "Ref=" & instrsrvbase & " Len=" & Len(Readv8ifilesplit(i))
				ibases(ibasesi,2) = mid(Readv8ifilesplit(i), instrsrvbase+4, Len(Readv8ifilesplit(i))-instrsrvbase-4)
				If mid(ibases(ibasesi,2),1,1)="""" Then
					ibases(ibasesi,2)=mid(ibases(ibasesi,2), 2, Len(ibases(ibasesi,2))-1)
				End If
				
				If mid(ibases(ibasesi,2), Len(ibases(ibasesi,2)), 1)=";" Then
					ibases(ibasesi,2)=mid(ibases(ibasesi,2), 1, Len(ibases(ibasesi,2))-1)
				End If
				
				If mid(ibases(ibasesi,2), Len(ibases(ibasesi,2)), 1)="""" Then
					ibases(ibasesi,2)=mid(ibases(ibasesi,2), 1, Len(ibases(ibasesi,2))-1)
				End If
				
				'ibases(ibasesi,2) = mid(ibases(ibasesi,2),2, Len(ibases(ibasesi,2))-2)
				WriteLog "Ref=" & ibases(ibasesi,2) & " ibasesi=" & ibasesi
			End If
			
			instrsrvextexternal = instr(1, Readv8ifilesplit(i), "External=")
			If instrsrvextexternal > 0 Then
				instrsrvsep3=instr(instrsrvextexternal, Readv8ifilesplit(i), "=")
				'msgbox mid(Readv8ifilesplit(i), instrsrvbase+4, instrsrvsep2-instrsrvbase-4)
				'WriteLog mid(Readv8ifilesplit(i), instrsrvextexternal+4, instrsrvsep3-instrsrvextexternal-4)
				'WriteLog mid(Readv8ifilesplit(i), instrsrvbase+10, Len(Readv8ifilesplit(i))-instrsrvbase-8)
				ibases(ibasesi,3) = mid(Readv8ifilesplit(i), instrsrvbase+10, Len(Readv8ifilesplit(i))-instrsrvbase-8)
				WriteLog "External=" & ibases(ibasesi,3) & " ibasesi=" & ibasesi
			End If
		End If
	Next
End Sub

'Чтение 1CEStart.cfg, CEStart(0,0): 00-CommonInfoBases (Путь к базе), 01-Имя базы
Sub Read1ce()
	Dim i, instrCIB, instrsrvsep1, splitbase, j 'readfilesplit-Public!!!'
	Read1cefilesplit = split(ReadFile(EnvironmentVariables("%APPDATA%") & "\" & str1CEStart), vbcrlf)
	bomread1ce=bom
	
	For i=0 to ubound(Read1cefilesplit)
		instrCIB=instr(1,Read1cefilesplit(i),"CommonInfoBases")
		If instrCIB > 0 Then
			instrsrvsep1=instr(instrCIB, Read1cefilesplit(i), "=")
			CEStart(i,0) = mid(Read1cefilesplit(i), instrCIB+16, Len(Read1cefilesplit(i))-instrCIB-15)
			
			splitbase=split (CEStart(i,0), "\")
			For j=0 to ubound(splitbase)
				If Right (splitbase(j), 4) = ".v8i" Then
					CEStart(i,1) = splitbase(j)
					CEStart(i,1) = mid(CEStart(i,1), 1, Len(CEStart(i,1))-4)
				End If
			Next
			'CEStart(i,1) = 
			WriteLog "CommonInfoBases=" & CEStart(i,0) & " Base=" & CEStart(i,1)
		End If
	Next
	
End Sub

'Чтение 1C files
Function ReadFile(filename)
	WriteLog "ReadFile " & filename
	Dim fso1, f, bomtab, stream
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	
	bom = ""
	Set f = fso.OpenTextFile(filename)
	Do Until f.AtEndOfStream Or bom = "y?" or bom = "яю" Or bom = "?y" Or Len(bom) >= 3
		bom = bom & f.Read(1)
	Loop
	f.Close

	Select Case bom
		Case "y?", "?y", "яю"  'UTF-16 text
			Set f = fso1.OpenTextFile(filename, 1, False, -1)
			ReadFile = f.ReadAll
			f.Close
			bomtab=2
			WriteLog "UTF-16" & " bom=" & bom
		Case "i»?", "п»ї"       'UTF-8 text
			Set stream = CreateObject("ADODB.Stream")
			stream.Open
			stream.Type = 2
			stream.Charset = "utf-8"
			stream.LoadFromFile filename
			ReadFile = stream.ReadText
			stream.Close
			bomtab=3
			WriteLog "UTF-8" & " bom=" & bom
		Case Else        'ASCII text
			Set f = fso.OpenTextFile(filename, 1, False, 0)
			ReadFile = f.ReadAll
			f.Close
			bomtab=0
			WriteLog "ASCII" & " bom=" & bom
	End Select

	WriteLog ReadFile
	WriteLog "//ReadFile"
End Function

'Информация ОС: имя ПК, версия, разрядность, тип
Sub InfoOS() 'обновить с рабочего
	Dim strComputer, VersionOS
	Dim objSWbemLocator, objSWbemServicesEx, collSWbemObjectSet, objSWbemObjectEx

	Set objSWbemLocator    = WScript.CreateObject("WbemScripting.SWbemLocator")
	Set objSWbemServicesEx = objSWbemLocator.ConnectServer(strComputer, "root\cimv2")
	Set collSWbemObjectSet = objSWbemServicesEx.InstancesOf("Win32_OperatingSystem")

	For Each objSWbemObjectEx In collSWbemObjectSet
		With objSWbemObjectEx
			'msgbox "123321" & .Caption
			CompName = .CSName
			VersionOS = .Version
			VersionOS = left(VersionOS, 3)
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
					Select Case VersionOS
						Case "8"
							OSVersion = "8"
						Case "8.1"
							OSVersion = "8.1"
						Case "10."
							OSVersion = "10"
						Case Else
							WriteLog "OS not know:        " & .Version & " veros:" & VersionOS & " .OSArchitecture " & .OSArchitecture
					End Select
					
					Select Case .OSArchitecture
						Case "32-разрядная"       ' X86
							'strCommandLine = strShare & "\VPN_7_32.exe"
							OSArch = "x32"
						Case "64-разрядная"       ' X64
							'strCommandLine = strShare & "\VPN_7_64.exe"
							OSArch = "x64"
                End Select
			End Select
		End With
	Next
	
	WriteLog("OS:                 " & OSVersion & " " & OSArch)
	WriteLog("Type:               " & sProductType)
	WriteLog("Computer Name:      " & CompName)
End Sub

'Backup ibases.v8i и 1CEStart.cfg
Sub CopyFiles(path) 'Должно быть простое копирование, откуда и куда!
	Dim fbacfile, fCEpath, fbacpath, fCEfile
	fCEpath=FSO.GetParentFolderName(path)
	fbacpath = fCEpath & "\bac"
	fCEfile = FSO.GetFileName (path)
	CreateFolder fbacpath
	fbacfile = fbacpath & "\" & fCEfile & "." & DatePart("yyyy", now) & DatePart("m", now) & DatePart("d", now) & hour(now) & minute(now) & second(now) 'int(1000*Rnd)' заменить
	'bacfiles = file & "." & DatePart("yyyy", now) & DatePart("m", now) & DatePart("d", now) & hour(now) & minute(now) & second(now) 'int(1000*Rnd)' заменить на время
	
	'If FileExist(fbacfile) = 0 Then
	'	FSO.CopyFile path, fbacfile
	'Else
	'	Randomize
	'	fbacfile = fbacfile & int(10000000000*Rnd)
	'	FSO.CopyFile path, fbacfile
	'End If
	
	FSO.CopyFile path, GenerateFileName(fbacfile)
	
End Sub

'Генератор имени, случайное число+желаемое имя
Function GenerateFileName(filename)
	If FileExist(filename) = 0 Then
		GenerateFileName=filename
	Else
		Randomize
		filename = filename & int(10000000000*Rnd)
	End If
End Function

Sub SravnenieBases() 'сравнение ibases и 1CEStart, на наличие баз
	Dim i, j, addlinkbase, delnolinkbase, allowdel, allowadd
	allowadd=0
	allowdel=0
	For i=0 to ubound(ibases) 'перебор имени базы
		addlinkbase=1 'Маркер добавления базы
		If ibases(i,2) <> "" Then 'если имя базы не пустое 
			'If ibases(i,3)="1" Then 'Если база не по ссылке
				For j=0 to ubound(CEStart) ' перебор имени базы в ссылках
					If CEStart(j,1) <> "" Then 'если имя базы не пустое
						If ibases(i,2) = CEStart(j,1) Then 'если имя баз сходится - 1 условие
							WriteLog ("Найдена совпадающая база:" & ibases(i,0) & " " & ibases(i,2) & "-" & CEStart(j,1))
							'addlinkbase=0
							If ibases(i,1) = """servertsdata""" Or ibases(i,1) = """192.168.19.3""" Or ibases(i, 1) = """servertsdata:2541""" Or ibases(i, 1) = """192.168.19.3:2541""" Then '3 условие - проверка, что 19.3/servertsdata 'работа как обычно. А порт
								addlinkbase=0
								If ibases(i,3) = 0 Then 'Если External=0, то есть база локально прописана - 2 условие
									WriteLog ("База настроена вручную: необходима чистка")
									ibases(i,4) = "del" 'пометка, что эту базу/раздел, необходимо убрать, так как они будут добавлены ссылкой
									allowdel=1
								End If
							Else
							
							End If
						End If
					End If
				Next
				If addlinkbase=1 Then 'Если база не была найдена в ссылках, то добавлям ссылку 'allowadd
					WriteLog ("Не найдена база в 1CEStart.cfg:" & ibases(i,0) & " " & ibases(i,1) &  " " & ibases(i,2))
					ibases(i,4) = "add"
					allowadd=1
				End If
			'End If
		End If
	Next
	If allowadd=1 Then 'addlinkbase
		call addbase()
		allowdel=1
	End If
	
	If allowdel=1 Then
		call removebase()
	End If
End Sub

'Ping
Function Ping (strTarget)
	Dim objWMIService, colPings, objPing
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strTarget & "'")
	For Each objPing in colPings
		Ping = objPing.StatusCode
	Next
End Function

'Проверка на доступность сетового каталога
Function TestServerLinks()
	iPing = Ping(strComputerName)
	If iPing=0 Then
		WriteLog("Есть связь с сервером")
		If FolderExist(strShare) = 1 Then
			WriteLog("Каталог существует  " & strShare & " " & TestServerLinks)
			TestServerLinks=1
		Else
			If FolderExist(strShare) = 0 Then
				WriteLog("Каталог не существует  " & strShare)
				TestServerLinks=0
			End If
		End If
	Else
		WriteLog("Нет связи с сервером, код ошибки: "& iPing)
		WriteLog("Необходимо проверить интернет Код: " & iPing)
	End If
End Function

'Составление нового 1CEStart.cfg, с учётом ibases (указаны базы, которые необходимо добавить) и Read1cefilesplit
Sub addbase()
	Dim i, j, k, l, allowadd, notadd'arfCEStart-глобальный
	j=0
	k=0
	allowadd=0
	notadd=0
	'arCEStart-массив для записи в файл
	ReDim Preserve arCEStart(j)
	For i=0 to ubound (Read1cefilesplit)
	'readfilesplit(i) = "CommonInfoBases="
		If instr(1, Read1cefilesplit(i), "CommonInfoBases=") > 0 Then 'правильный поиск CommonInfoBases+
			ReDim Preserve arCEStart(j)
			arCEStart(j) = Read1cefilesplit(i) 'сохранение всей строки+
			j=j+1
		Else
			If k=0 Then 'Необходимо не допустить повторное выполнение цикла
				For k=0 to ubound(ibases)
					If ibases(k, 4) = "add" Then
						For l=0 to ubound(ibases) 'перебор на наличие одинаковой базы
							If ibases(k, 4) = "add" And ibases(l, 4) = "added" And ibases(k, 2) = ibases(l, 2) & ibases(k, 1) = ibases(l, 1) Then
								WriteLog("Не будет добавлена база:" & " " & ibases(k, 0) & " " & ibases(k, 2) & "=" & ibases(l, 0) & " " & ibases(l, 2))
								notadd=notadd+1
							End If
						Next
						If notadd=0 Then
							allowadd=1
							ReDim Preserve arCEStart(j)
							'Имя и порт
							If ibases(k, 1) = """servertsdata""" Or ibases(k, 1) = """192.168.19.3""" Or ibases(k, 1) = """servertsdata:2541""" Or ibases(k, 1) = """192.168.19.3:2541""" Then
								arCEStart(j) = "CommonInfoBases=" & strShare & "\" & ibases(k, 2) & ".v8i"
								'notadd=0' ?
								ibases(k, 4) = "added"
								WriteLog("Добавление ссылки на сервер")
								createlink strShare & "\" & ibases(k, 2) & ".v8i", ibases(k, 1), ibases(k, 2), ibases(k, 0)
							Else
								arCEStart(j) = "CommonInfoBases=" & strShare & "\" & ibases(k, 2) & "-" & replace(replace(ibases(k, 1),":","_"), """", "") & ".v8i"
								'notadd=0' ?
								ibases(k, 4) = "added"
								WriteLog("Добавление ссылки на сервер:" & strShare & "\" & ibases(k, 2) & "-" & replace(replace(ibases(k, 1),":","_"), """", "") & ".v8i")
								createlink strShare & "\" & ibases(k, 2) & "-" & replace(replace(ibases(k, 1),":","_"), """", "") & ".v8i", ibases(k, 1), ibases(k, 2), ibases(k, 0)
							End If
							j=j+1
						End If
					End If
					notadd=0
				Next
			End If
			ReDim Preserve arCEStart(j) 'удалить?
			arCEStart(j) = Read1cefilesplit(i) 'сохранение всей строки+ 'удалить?
			j=j+1 'удалить?
		End If
	Next
	'call arrwrite(arCEStart)
	If allowadd=1 Then
		bom=bomread1ce
		Writeaddbase arraytofile(arCEStart), EnvironmentVariables("%APPDATA%") & "\" & str1CEStart
		'removebase
	End If
End Sub

Sub removebase()
	Dim i, j, k
	'Readv8ifilesplit
	'ibases
	WriteLog("removebase")
	For i=0 to ubound(ibases)
		If ibases(i,4) = "del" Or ibases(i,4) = "added" Then
			WriteLog("Раздел должен быть удалён " & ibases(i,0) & " " & ibases(i,2))
			For j=0 to ubound(Readv8ifilesplit)
				'msgbox ibases(i,0) & vbcrlf & "[" & Readv8ifilesplit(j) & "]"
				If ibases(i,0) = Readv8ifilesplit(j) Then
					WriteLog("Найден раздел:" & Readv8ifilesplit(j))
					Readv8ifilesplit(j) = ""
					'j=j+1
					For k=0 to Ubound(Readv8ifilesplit)
						j=j+1
						k=j
						If Left(Readv8ifilesplit(k),1) <> "[" Then
							WriteLog(Readv8ifilesplit(j))
							Readv8ifilesplit(j) = ""
						Else
							j=Ubound(Readv8ifilesplit)
							k=j
						End If
					Next
				End If
			Next
			ibases(i,4) = "deleted"
		End If
	Next
	bom=bomreadv8i
	Writeaddbase arraytofile(Readv8ifilesplit), EnvironmentVariables("%APPDATA%") & "\" & stribases
	WriteLog("removebase")
End Sub

'Вывод массива в лог
Function arrwrite(arr)
	Writelog("///arrwrite")
	Dim i, arrtext
	Writelog("arrwrite 1=" & Ubound(arr,1))
	For i=0 To Ubound(arr,1)
		'WriteLog("		" & arr(i,j) & "  Len=" & Len(arr(i,j)))
		arrtext=arr(i)
		writelog (arrtext)
	Next
	Writelog("arrwrite///")
End Function

'конвертация массива в обычный файл. Join
Function arraytofile(ar)
	WriteLog("arraytofile")
	Dim i
	For i=0 to ubound(ar)
		If ar(i) <> "" Then
			arraytofile=arraytofile & ar(i) & vbcrlf
		End If
	Next
End Function

'Создание файла новой ссылки
Sub createlink (cllink, conSRV, cdbase, clNamebase) 'Имя раздела брать от external0
	Dim NewFileText, NewFile
	WriteLog("createlink")
	If TestServerLinks = 1 And iPing=0 Then
		If FileExist(cllink) = 1 Then
			WriteLog("Ссылка существует:" & cllink)
		Else
			WriteLog("Создаётся ссылка:" & cllink)
			NewFileText = clNamebase & vbcrlf & "Connect=Srvr=" & conSRV & ";Ref=""" & cdbase & """;" & vbcrlf & "ClientConnectionSpeed=Normal" & vbcrlf & "App=Auto" & vbcrlf & "WA=1"
			'NewFile = GenerateFileName(EnvironmentVariables("%APPDATA%")) & "\" & FSO.GetFileName(cllink)
			'CreateFile NewFile
			'CreateFile cllink
			bom=bomreadv8i
			'Writeaddbase NewFileText, cllink
			'FSO.CopyFile NewFile, cllink
			'msgbox FSO.GetAbsolutePathName(NewFile) & vbcrlf & NewFile
		End If
	End If
	WriteLog("//createlink")
End Sub

'Выполнить запись в 1C files
Sub Writeaddbase(WriteFile, filename)
	WriteLog("Writeaddbase")
	Dim fso1, f, stream
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	WriteLog("bom:" & bom)
	WriteLog(WriteFile)
	WriteLog(filename)
	Select Case bom
		Case "y?", "?y", "яю"  'UTF-16 text
			Set f = fso1.OpenTextFile(filename, 2, True, -1)'do not create-False
			f.Write WriteFile
			f.Close
			WriteLog "UTF-16" & " bom=" & bom
		Case "i»?", "п»ї"       'UTF-8 text
			Set stream = CreateObject("ADODB.Stream")
			stream.Type = 2
			stream.Charset = "utf-8"
			stream.Open
			stream.WriteText WriteFile
			stream.SaveToFile filename, 2
			stream.Close
			WriteLog "UTF-8" & " bom=" & bom
		Case Else        'ASCII text
			Set f = fso.OpenTextFile(filename, 2, True, 0)
			f.Write WriteFile
			f.Close
			WriteLog "ASCII" & " bom=" & bom
	End Select
	WriteLog("//Writeaddbase")
End Sub