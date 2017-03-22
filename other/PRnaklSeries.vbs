Option explicit
Dim OApplication, ODesktop, ODocument, srcFile, srcProps() ' ******************
Dim oServiceManager, oCalcDoc, oBook, oSheet, oCells
Dim args(0)

Dim Sep,St, C, n, startcol
Dim DeviceInp,DeviceOut
Dim TableName,dbfPrice,dbfSum,dbfConn
Dim FSO
Dim EAN, ConEnt, EntExl
Dim NameXLS, FDirOut
Const strLenEAN=12
Const SKM="СК-МедФарм"

Const DirDBF= "X:\PRnakl\" '"C:\braki\1\"
Const NameDBF="Prichod.dbf"
Randomize
Const TemplateFile="X:\PRnakl\PRnakl.ots" ' Файл с шаблоном отчета
FDirOut="X:\PRnakl\"   ' Путь куда выкладывать файл с отчетом
NameXLS="общий1 склад " & Date & " " & Int(Rnd*999999) & ".xls"

Set FSO = CreateObject("Scripting.FileSystemObject")

C=Chr(34) ' Двойные кавычки для строк
'********************************************************
'создаем новый ServiceManager
Set oServiceManager = CreateObject("com.sun.star.ServiceManager")
Set oCalcDoc = oServiceManager.createInstance("com.sun.star.frame.Desktop")
' создаем новую книгу OpenOffice.org Calc
Set args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
args(0).Name = "AsTemplate"
args(0).Value = True
Set oBook = oCalcDoc.loadComponentFromURL(convertToURL(TemplateFile), "_blank", 0, args)
'получаем ссылку на второй!!!!!!!!!!!!!!!! лист новой книги
Set oSheet = oBook.getSheets().getByIndex(1)
' т.е. чтобы получить ячейку в первом столбце первой строки пишем oSheet.getCellByPosition(0,0)
'кроме того в getCellByPosition первый аргумент столбец, второй строка (в Excel наоборот)
'***************************************************************************************************************************************

n=0
startcol=7

'********************************************************
Set dbfConn = CreateObject("ADODB.Connection")
With dbfConn
  .Provider = "Microsoft.Jet.OLEDB.4.0"
  .Properties("Extended Properties") = "dBase IV"
  .Properties("Data Source") = DirDBF
  .Open
End With
  
Set dbfPrice = CreateObject("ADODB.Recordset")
Set dbfPrice.ActiveConnection = dbfConn
'********************************************************

dbfPrice.Open "SELECT * FROM ["&NameDBF&"]"

'msgbox dbfPrice.Fields("NAMESUB2")
If InStr(1,dbfPrice.Fields("NAMESUB2"), SKM) Then
	FDirOut="X:\PRnakl\SKM\"   ' Путь куда выкладывать файл с отчетом
	NameXLS="Склад СК-МедФарм " & Date & ".xls"
Else
	FDirOut="X:\PRnakl\Kor\"   ' Путь куда выкладывать файл с отчетом
	NameXLS="Общий Склад " & Date & ".xls"
End If

Do
	EAN=inputbox("Введите штрих-код:")
	EAN=CheckLenName(EAN)
	'msgbox EAN
	dbfPrice.MoveFirst
	Do While Not dbfPrice.Eof
		If fReadDBFFixNull("BARCODE") = EAN Then
			'        while Len(Trim(xlWbk.getCellByPosition(3,k).String))>0       ' Пока содержимое первой ячейки текущей строки непустое, берем данные
			Call oSheet.getCellByPosition(0, startcol+n).SetFormula(1+n)
			Call oSheet.getCellByPosition(1, startcol+n).SetFormula(fReadDBFFixNull("SUPDATE") & ", №" & fReadDBFFixNull("SUPDOC"))
			Call oSheet.getCellByPosition(2, startcol+n).SetFormula(fReadDBFFixNull("POST"))
			Call oSheet.getCellByPosition(3, startcol+n).SetFormula(fReadDBFFixNull("NAMETOV"))
		'	If dbfPrice.Fields("SERIES") <> "" Then
				Call oSheet.getCellByPosition(4, startcol+n).SetFormula(fReadDBFFixNull("SERIES"))
			'	Call oSheet.getCellByPosition(5, startcol+n).SetFormula(fReadDBFFixNull("SERIES"))
		'	End If
			Call oSheet.getCellByPosition(6, startcol+n).SetFormula(fReadDBFFixNull("KOL"))
			'Call oSheet.getCellByPosition(7, startcol+n).SetFormula(fReadDBFFixNull("KOL"))
			Call oSheet.getCellByPosition(8, startcol+n).SetFormula("Без повреждений")
			Call oSheet.getCellByPosition(9, startcol+n).SetFormula("Соответствует")
			Call oSheet.getCellByPosition(10, startcol+n).SetFormula("Укомплектованно")
			Call oSheet.getCellByPosition(11, startcol+n).SetFormula("-")
			'серия Call oSheet.getCellByPosition(3, startcol+n).SetFormula(Mid(Replace(xlWbk.getCellByPosition(3,k).String,C,"'"),1,200))
			'Call oSheet.getCellByPosition(3, startcol+n).SetFormula(dbfPrice.Fields("CODESUP"))
			Call oSheet.Rows.insertByIndex(startcol+n+1, 1)
			n=n+1
			'        wend
		End If
		dbfPrice.MoveNext
		if Err.Number then Exit Do
	Loop
	ConEnt=msgbox ("Хотите ввести ещё штрих-код?", vbInformation + vbOKCancel)
	If ConEnt = 2 Then EntExl = 100
Loop While EntExl < 10

dbfPrice.Close

Dim mFileType(0), ArgsSave(1)

'Set mFileType(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")

'dummy(0).Name = "Overwrite"
'dummy(0).Value = True

'mFileType(0).Name = "Overwrite"
'mFileType(0).Value = True
'mFileType(0).Name = "FilterName"
'mFileType(0).Value="MS Excel 97"

Set ArgsSave(0)=MakePropertyValue("Overwrite", False)
Set ArgsSave(0)=MakePropertyValue("FilterName", "MS Excel 97")
'msgbox ConvertToURL(FDirOut & NameXLS)
call oBook.storeAsUrl(ConvertToURL(FDirOut & NameXLS),ArgsSave)

'fExportAs oBook, "file:///C:/braki/1/1CDocs.xls"
'Function fExportAs(oDoc, sURL)
'sURL = convertToURL(sFile)
'dim sType
'sType="999"
'if sType="999" then
'	'if isMissing(sType) then
'	   oDoc.storeToURL sURL
'	else
'	  dim mFileType(0)
'	  mFileType(0) = createUnoStruct("com.sun.star.beans.PropertyValue")
'	  mFileType(0).Name = "FilterName"
'	  mFileType(0).Value = sType
'	  oDoc.storeToURL sURL, mFileType()
'	end if
'end Function

msgbox ("OK")

Function fReadDBFFixNull(pr)
	If dbfPrice.Fields(pr) <> "" Then
		fReadDBFFixNull=dbfPrice.Fields(pr)
		else	
		fReadDBFFixNull=" "
	End If
End Function

Function ConvertToURL(strFile)
	'msgbox strFile & " Finction in"
    strFile = Replace(strFile, "\", "/")
    'strFile = Replace(strFile, ":", "|")
    strFile = Replace(strFile, " ", "%20")
    strFile = "file:///" + strFile
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
    ConvertFromUrl = strFile1
End Function

Function DelBigName(LR, strDelBN)
  'Dim FFNew
	If LR = "L" Then
		strDelBN=Left(strDelBN, strLenEAN)
	End If
	If LR = "R" Then
		strDelBN=Right(strDelBN, strLenEAN)
	End If
	DelBigName=strDelBN
	'FFNew=FDir & strDelBN & ".dbf"
  'msgbox(FF & "--" & FFNew)
  'FF.Move (FFNew) 
  'msgbox ("Необходимо уменьшить имя накладной до 8-ми символов перед '.dbg'")
  'msgbox(strDelBN)
End Function


Function CheckLenName(CLN)
  If Len(CLN)>strLenEAN Then
	'msgbox CLN & vbCRLF & DelBigName("L", CLN)
	CheckLenName=DelBigName("L", CLN)
  Else
	CheckLenName=CLN
  End If
  
  If Len(CLN)>strLenEAN Then
    msgbox ("Необходимо уменьшить имя накладной до 8-ми символов перед '.dbf'. Количество символов=" & Len(CLN) & Chr(13)&Chr(10) & "Будет крах, звонить 911 с корпоративного телефона")
  End If
End Function

Function MakePropertyValue(cName, uValue)' As Object
	Dim oStruct', oServiceManager 'as Object
    'Set oServiceManager = CreateObject("com.sun.star.ServiceManager")
    Set oStruct = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    oStruct.Name = cName
    oStruct.Value = uValue
    Set MakePropertyValue = oStruct
End Function 