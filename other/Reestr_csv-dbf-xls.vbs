'Option explicit
Dim OApplication, ODesktop, ODocument, srcFile, srcProps() ' ******************
Dim oServiceManager, oCalcDoc, oBook, oSheet, oCells
Dim args(0)

Dim Sep,St, C, n, startcol
Dim DeviceInp,DeviceOut
Dim TableName,dbfPrice,dbfSum,dbfConn, dbfAdd
Dim FSO, FSOL, WshShell, DesktopPath
Dim EAN, ConEnt, EntExl
Dim NameXLS, OpenDiadocCSV, SegmentDiadocCSV
Dim ScriptName, ScriptFullName
Dim SUP(2), i, j, SUPin
Public ver, LocalShare, sTime, sDate, iWriteLog, PathFileLog, PathFullFileLog
Dim iarrsup, arrsup(1000,5)

Const DirDBF= "C:\Users\с_барбаянов\Desktop\vbs\реестр\" '"C:\braki\1\"
Const NameDBF="1CDocs.dbf"
Const DiadocCSV="Diadoc 23.01.18 15.09.csv"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FSOL = CreateObject("Scripting.FileSystemObject")
Set ScriptFullName = FSOL.GetFile(Wscript.ScriptFullName)
Set WshShell = CreateObject("WScript.Shell")

ScriptName = FSOL.GetFileName(ScriptFullName)
PathFileLog = "\" & ScriptName & ".log"
DesktopPath = WSHShell.SpecialFolders("Desktop")
LocalShare = DesktopPath 'WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
PathFullFileLog = LocalShare & PathFileLog

dbfAdd=0
Const TemplateFile="reestr.ots" ' Файл с шаблоном отчета
'Const FDirOut="X:\PRnakl\"   ' Путь куда выкладывать файл с отчетом
'NameXLS="общий склад " & Date & ".xls"

iWriteLog=1
ver="0.0.1"
sTime = Hour(time) & Minute(time) & Second(time)
sDate = Day(date) & Month(date) & Year(date)

WriteLog("                    ")
WriteLog("                    ")
WriteLog("                    ")
WriteLog("Start:              " & Time & " " & Right(0 & Day(date), 2) & "." & Right(0 & Month(date), 2) & "." & Year(Date))
WriteLog("Version:            " & ver)


'********************************************************
SUP(1) = "КАТРЕН"
SUP(2) = "Протек"

Set dbfConn = CreateObject("ADODB.Connection")
With dbfConn
  .Provider = "Microsoft.Jet.OLEDB.4.0"
  .Properties("Extended Properties") = "dBase IV"
  .Properties("Data Source") = DirDBF
  .Open
End With
  
Set dbfPrice = CreateObject("ADODB.Recordset")
'Set dbfPrice.ActiveConnection = dbfConn
'********************************************************
'dbfPrice.CursorType = "adOpenKeyset"
'dbfPrice.LockType = "adLockOptimistic"
'dbfPrice.Open "SELECT * FROM "&NameDBF, dbfConn, 2, 3

If dbfAdd = 1 Then
	dbfPrice.AddNew '"NameSUP", "Test"
	dbfPrice.Fields("NameSUP") = "Тест"
	dbfPrice.Fields("SUMZAK") = "2500"
	dbfPrice.Update
	dbfPrice.Close
End If

If CheckPathFile(DirDBF & DiadocCSV) = 1 Then
	Set OpenDiadocCSV = FSO.OpenTextFile(DirDBF & DiadocCSV, 1, False)
	Writelog ("Diadoc CSV -> dbf")
	Do While Not OpenDiadocCSV.AtEndOfStream
		SegmentDiadocCSV = Split(OpenDiadocCSV.ReadLine, ";")
		If left (SegmentDiadocCSV(0),3) = """=""" Then
			'Writelog(SegmentDiadocCSV(2))
			'msgbox SegmentDiadocCSV(3)
			If Instr(1,SegmentDiadocCSV(3),"счет", 1) Or Instr(1,SegmentDiadocCSV(3),"счёт", 1) Then 
				'WriteLog (SegmentDiadocCSV(2) & " For")
					SUPin=0
					For i = 1 To UBound(SUP) Step 1
						If Instr(1,SegmentDiadocCSV(2), SUP(i), 1) > 0 Then
							SUPin=SUPin + 1
						End If
					Next
					
				If SUPin = 0 Then
					Writelog("		" & SegmentDiadocCSV(2) & " " & SegmentDiadocCSV(5) & " " & SegmentDiadocCSV(6))
					'WriteLog (SegmentDiadocCSV(2) & " " & SUP(i))
					'Writelog (Instr(1,SegmentDiadocCSV(2), SUP(i), 1))
					
					
					dbfPrice.Open "SELECT * FROM "&NameDBF, dbfConn, 2, 3
					dbfPrice.AddNew
					dbfPrice.Fields("NameSUP") = SegmentDiadocCSV(2)
					'MSGBOX mid(SegmentDiadocCSV(4), 5, Len(SegmentDiadocCSV(4))-7)
					dbfPrice.Fields("NUMSUP") = mid(SegmentDiadocCSV(4), 5, Len(SegmentDiadocCSV(4))-7)
					'msgbox fixorder(SegmentDiadocCSV(4)) & "--+" & SegmentDiadocCSV(4)
					dbfPrice.Fields("DateSUP") = SegmentDiadocCSV(5)
					dbfPrice.Fields("SUMZAK") = SegmentDiadocCSV(6)
					dbfPrice.Fields("NDSZ_10") = Replace(fixnuulnum(SegmentDiadocCSV(8)), ".", ",")
					dbfPrice.Fields("NDSZ_18") = Replace(fixnuulnum(SegmentDiadocCSV(9)), ".", ",")
					'msgbox SegmentDiadocCSV(2) & vbcrlf & fixnuulnum(SegmentDiadocCSV(8)) & vbcrlf & fixnuulnum(SegmentDiadocCSV(9) & vbcrlf)
					dbfPrice.Update
					dbfPrice.Close
				End If
			End If
		End If
	Loop
	Writelog ("///Diadoc CSV -> dbf")
	OpenDiadocCSV.close
	
	'********************************************************
	'создаем новый ServiceManager
	Set oServiceManager = CreateObject("com.sun.star.ServiceManager")
	Set oCalcDoc = oServiceManager.createInstance("com.sun.star.frame.Desktop")
	' создаем новую книгу OpenOffice.org Calc
	Set args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
	args(0).Name = "AsTemplate"
	args(0).Value = True
	Set oBook = oCalcDoc.loadComponentFromURL(convertToURL(DirDBF & TemplateFile), "_blank", 0, args)
	'получаем ссылку на второй!!!!!!!!!!!!!!!! лист новой книги
	Set oSheet = oBook.getSheets().getByIndex(0)
	' т.е. чтобы получить ячейку в первом столбце первой строки пишем oSheet.getCellByPosition(0,0)
	'кроме того в getCellByPosition первый аргумент столбец, второй строка (в Excel наоборот)
	'***************************************************************************************************************************************
'dim oCursor, oEndAdr
'CellCursor=oSheet.createCursor()

'Set oSheet = oBook.Sheets(1)
Set oCursor = oSheet.createCursor()
oCursor.gotoEndOfUsedArea(true) 
Set oEndAdr = oCursor.getRangeAddress
'msgbox oEndAdr.EndRow
'msgbox oEndAdr.Endcolumn

	n=0
	startcol=1

	'********************************************************
	dbfPrice.Open "SELECT * FROM ["&NameDBF&"]"
	If CheckPathFile(DirDBF & TemplateFile) = 1 Then
		'Do
			dbfPrice.MoveFirst
			Writelog("Create array")
			Do While Not dbfPrice.Eof
					For iarrsup=1 To Ubound(arrsup,1)
						'WriteLog("debug " & dbfPrice.Fields("DateSUP") & "=" & arrsup(iarrsup,3) & "---" & dbfPrice.Fields("NameSUP") & "=" & arrsup(iarrsup,1) & "---" & iarrsup)
						If dbfPrice.Fields("DateSUP") = arrsup(iarrsup,3) And dbfPrice.Fields("NameSUP") = arrsup(iarrsup,1) Then
							arrsup(iarrsup,2)=arrsup(iarrsup,2)+dbfPrice.Fields("SUMZAK")
							arrsup(iarrsup,4)=arrsup(iarrsup,4) & ", " & dbfPrice.Fields("NUMSUP")
							'writelog ("		Поставщик-старая сумма = 10+18" arrsup(iarrsup,1) & "-" & arrsup(iarrsup,5) & "=" & dbfPrice.Fields("NDSZ_10") & "+" & dbfPrice.Fields("NDSZ_18"))
							writelog("		Поставщик-старая сумма = 10+18: " & arrsup(iarrsup,1) & "-" & fixnuulnum(arrsup(iarrsup,5)) & "=" & fixnuulnum(dbfPrice.Fields("NDSZ_10")) & "+" & fixnuulnum(dbfPrice.Fields("NDSZ_18")) & "=" & fixnuulnum(arrsup(iarrsup,5)) + fixnuulnum(dbfPrice.Fields("NDSZ_10")) + fixnuulnum(dbfPrice.Fields("NDSZ_18")))
							arrsup(iarrsup,5)=arrsup(iarrsup,5) + fixnuulnum(dbfPrice.Fields("NDSZ_10")) + fixnuulnum(dbfPrice.Fields("NDSZ_18"))
							'Writelog("=" & arrsup(iarrsup,5))
							
							Exit For
						Else
							If arrsup(iarrsup,1) = "" Then
								'WriteLog("arrsup(iarrsup,1) = null --- " & arrsup(iarrsup,1))
								arrsup(iarrsup,1)=dbfPrice.Fields("NameSUP")
								arrsup(iarrsup,2)=dbfPrice.Fields("SUMZAK")
								arrsup(iarrsup,3)=dbfPrice.Fields("DateSUP")
								arrsup(iarrsup,4)=dbfPrice.Fields("NUMSUP")
								'writelog (arrsup(iarrsup,1) & "+" & dbfPrice.Fields("NDSZ_10") & "+" & dbfPrice.Fields("NDSZ_18"))
								'writelog (arrsup(iarrsup,1) & "-" & arrsup(iarrsup,5) & "=" & dbfPrice.Fields("NDSZ_10") & "+" & dbfPrice.Fields("NDSZ_18"))
								arrsup(iarrsup,5)=fixnuulnum(arrsup(iarrsup,5)) + fixnuulnum(dbfPrice.Fields("NDSZ_10")) + fixnuulnum(dbfPrice.Fields("NDSZ_18"))
								If arrsup(iarrsup,5) = "" Then 
									arrsup(iarrsup,5)=0
								End If
								'Writelog ("else=" & arrsup(iarrsup,1) & "|" & arrsup(iarrsup,2) & "|" & arrsup(iarrsup,3))
								'WriteLog(arrsup(iarrsup,1))
								Exit For
							End If
						End If
					'________________'
		'					Call oSheet.getCellByPosition(0, startcol+n).SetFormula(dbfPrice.Fields("NameSUP"))
		'					'Call oSheet.getCellByPosition(1, startcol+n).SetFormula(FormatNumber(dbfPrice.Fields("DateSUP"),0))
		'					Call oSheet.getCellByPosition(1, startcol+n).SetFormula(mid(dbfPrice.Fields("DateSUP"), 4, 2) & "." & mid(dbfPrice.Fields("DateSUP"), 1, 2) & "." & mid(dbfPrice.Fields("DateSUP"), 7,4)) '(dbfPrice.Fields("DateSUP"))
		'					Call oSheet.getCellByPosition(3, startcol+n).SetFormula(Replace(dbfPrice.Fields("SUMZAK"), ",", "."))
		'					WriteLog (dbfPrice.Fields("NameSUP") & "-111-" & dbfPrice.Fields("DateSUP") & "-111-" & dbfPrice.Fields("SUMZAK"))
		'					Call oSheet.Rows.insertByIndex(startcol+n+1, 1)
		'					n=n+1
						'____________'
						'iarrsup=iarrsup+1
						
						if Err.Number then Exit Do

					Next
			dbfPrice.MoveNext
			Loop
			Writelog("///Create array")
			Call arrwrite(arrsup)
			Writelog("Array -> xls")
			For iarrsup=1 To Ubound(arrsup,1)
				If arrsup(iarrsup,1) <> "" Then
					Call oSheet.getCellByPosition(0, startcol+n).SetFormula(arrsup(iarrsup,1))
					'Call oSheet.getCellByPosition(1, startcol+n).SetFormula(FormatNumber(dbfPrice.Fields("DateSUP"),0))
					Call oSheet.getCellByPosition(1, startcol+n).SetFormula(mid(arrsup(iarrsup,3), 4, 2) & "." & mid(arrsup(iarrsup,3), 1, 2) & "." & mid(arrsup(iarrsup,3), 7,4)) '(dbfPrice.Fields("DateSUP"))
					Call oSheet.getCellByPosition(3, startcol+n).SetFormula(Replace(arrsup(iarrsup,2), ",", "."))
					writelog("		Order: " & arrsup(iarrsup,4))
					Call oSheet.getCellByPosition(4, startcol+n).SetFormula(arrsup(iarrsup,4))
					writelog("		Поставщик-NDS:" & arrsup(iarrsup,1) & "-" & arrsup(iarrsup,5))
					Call oSheet.getCellByPosition(5, startcol+n).SetFormula(Replace(arrsup(iarrsup,5), ",", "."))
		'			WriteLog (dbfPrice.Fields("NameSUP") & "-111-" & dbfPrice.Fields("DateSUP") & "-111-" & dbfPrice.Fields("SUMZAK"))
					Call oSheet.Rows.insertByIndex(startcol+n+1, 1)
					n=n+1
				Else
					Exit For
				End If
			Next
			Writelog("///Array -> xls")
			'ConEnt=msgbox ("Хотите ввести ещё штрих-код?", vbInformation + vbOKCancel)
			'If ConEnt = 2 Then EntExl = 100
		'Loop While EntExl < 10
	End If
End If







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

Function ConvertToURL(strFile)
	'msgbox strFile & " Finction in"
    strFile = Replace(strFile, "\", "/")
    'strFile = Replace(strFile, ":", "|")
    strFile = Replace(strFile, " ", "%20")
    strFile = "file:///" + strFile
	WriteLog(strFile)
    ConvertToUrl = strFile
End Function

Function ConvertFromURL(strFile1)
	Dim objRegExp
	
	Set objRegExp = CreateObject("VBScript.RegExp")
	objRegExp.Pattern = "file:///"
	strFile1 = objRegExp.Replace(strFile1, "")
    strFile1 = Replace(strFile1, "/", "\")
    'strFile = Replace(strFile, ":", "|")
    strFile1 = Replace(strFile1, "%20", " ")
    'strFile1 = "file:///" + strFile1
	WriteLog(strFile1)
    ConvertFromUrl = strFile1
End Function

Function arrwrite(arr)
	Writelog("arrwrite 1=" & Ubound(arr,1) & " 2=" & Ubound(arr,2))
	For i=1 To Ubound(arr,1)
		For j=1 To Ubound(arr,2)
		'msgbox Ubound(arr,2)
			If arr(i,j) <> "" Then
				WriteLog("		" & arr(i,j) & "  Len=" & Len(arr(i,j)))
			End If
		Next
	Next
	Writelog("///arrwrite")
End Function

Function fixnuulnum(nullnum)
'writelog("nullnum=" & nullnum)
	If nullnum = null Or nullnum = "" Or Len(nullnum)=""  Then
		fixnuulnum=0
		'writelog("0nullnum=" & fixnuulnum)
	Else
		If Len(nullnum)>0 Then
			fixnuulnum=nullnum
			'writelog("1nullnum=" & "'" & fixnuulnum & "'")
		End If
	End If
End Function