'Option Explicit
Dim FSO,C,FDir,FLD,FL,FF,Sh,FDirOut,dbfConn,dbfRS,LetLab,Letdate,LetNum,xlsFiles,xlsStrs,x1,x2,MonthYear,ar1
Dim xlglob,Desktop,Document,sheets,xlWbk
Dim Mass()
Dim aNoArgs()
Dim oMyStyle
Dim OApplication, ODesktop, ODocument, srcFile, srcProps() ' ******************
Dim args(0)

FDir="\\129.186.1.24\holdingswap\03 ����������\������ ��������\"      ' ����, ��� �������� ��������� � ����� � ������� Excel
FDirOut="C:\braki\"   ' ���� ���� ����������� ���� � �������
TemplateFile="C:\braki\ReportLSA1.ots" ' ���� � �������� ������

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files

C=Chr(34) ' ������� ������� ��� �����

MonthYear=InputBox("������� �����. ������: ����","������� �����")

'********************************************************
'������� ����� ServiceManager
Set oServiceManager = CreateObject("com.sun.star.ServiceManager")
Set oCalcDoc = oServiceManager.createInstance("com.sun.star.frame.Desktop")
' ������� ����� ����� OpenOffice.org Calc
Set args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
'	Set args(0) = createStruct("com.sun.star.beans.PropertyValue")
args(0).Name = "AsTemplate"
args(0).Value = True
Set oBook = oCalcDoc.loadComponentFromURL("file:///"&TemplateFile, "_blank", 0, args)
'�������� ������ �� ������!!!!!!!!!!!!!!!! ���� ����� �����
Set oSheet = oBook.getSheets().getByIndex(1)
' �.�. ����� �������� ������ � ������ ������� ������ ������ ����� oSheet.getCellByPosition(0,0)
'����� ���� � getCellByPosition ������ �������� �������, ������ ������ (� Excel ��������)
'***************************************************************************************************************************************

xlsFiles=0
xlsStrs=0
n=0
startcol=10	   
For Each FF in FL
  if (InStr(LCase(FF.Name),LCase(MonthYear) & " ��.xls")) then
    Set xlglob = CreateObject("com.sun.star.ServiceManager") 
    Set Desktop = xlglob.createInstance("com.sun.star.frame.Desktop")
    Set Document = Desktop.LoadComponentFromURL("file:///"&FDir&FF.Name, "_blank", 0, mass )
    Set sheets = Document.getSheets()
    Set xlWbk = sheets.getByIndex(0)
    k=0
    
    while xlWbk.getCellByPosition(0,k).String<>"��"
      k=k+1
    wend
    k=k+1
    'n=n+1
    'while Len(Trim(xlWbk.getCellByPosition(0,k).String))>0       ' ���� ���������� ������ ������ ������� ������ ��������, ����� ������
     while Len(Trim(xlWbk.getCellByPosition(2,k).String))>0       ' ���� ���������� ������ ������ ������� ������ ��������, ����� ������
 
'      if InStr(xlWbk.getCellByPosition(7,k).String,":")>0 then
'        LetLab = Mid(Replace(xlWbk.getCellByPosition(7,k).String,C,"'"),1,InStr(xlWbk.getCellByPosition(8,k).String,":")-1)
'      else
'        LetLab = ""
'      end if
      'LetNum = Mid(Replace(xlWbk.getCellByPosition(7,k).String,C,"'"),InStr(xlWbk.getCellByPosition(8,k).String,"�")+2,InStr(xlWbk.getCellByPosition(7,k).String," ��")-InStr(xlWbk.getCellByPosition(7,k).String,"�")-2)
      'LetDate = Mid(Replace(xlWbk.getCellByPosition(7,k).String,C,"'"),InStr(xlWbk.getCellByPosition(7,k).String,"�� ")+3,8)
      'x1 = Mid(LetDate,1,InStr(LetDate,"."))
      'LetDate = Mid(LetDate,InStr(LetDate,".")+1,Len(LetDate))
      'x2 = Mid(LetDate,1,InStr(LetDate,"."))
      'LetDate = Mid(LetDate,InStr(LetDate,".")+1,Len(LetDate))
      'LetDate = x1 & x2 & "20" & LetDate

'MsgBox "file:///"&FDir&FF.Name '��� �������     
'*******'
Call oSheet.getCellByPosition(0, startcol+n).SetFormula(1+n)
Call oSheet.getCellByPosition(1, startcol+n).SetFormula(Mid(Replace("������ ��� " & xlWbk.getCellByPosition(8,k).String,C,"'"),1,200))
'Call oSheet.getCellByPosition(1, startcol+n).SetFormula(LetNum)
Call oSheet.getCellByPosition(2, startcol+n).SetFormula(Mid(Replace(xlWbk.getCellByPosition(0,k).String & ", " & xlWbk.getCellByPosition(1,k).String,C,"'"),1,200))
Call oSheet.getCellByPosition(3, startcol+n).SetFormula(Mid(Replace(xlWbk.getCellByPosition(2,k).String,C,"'"),1,200))
Call oSheet.getCellByPosition(4, startcol+n).SetFormula(Mid(Replace(xlWbk.getCellByPosition(3,k).String,C,"'"),1,200) & ", " & Mid(Replace(xlWbk.getCellByPosition(4,k).String,C,"'"),1,200))
Call oSheet.getCellByPosition(5, startcol+n).SetFormula("�� ������������ �.�.")
Call oSheet.getCellByPosition(7, startcol+n).SetFormula("0")
Call oSheet.getCellByPosition(8, startcol+n).SetFormula("�� ��������")
Call oSheet.Rows.insertByIndex(startcol+n+1, 1)
'Call oSheet.getCellByPosition(8, startcol+n).SetFormula(Mid(Replace(xlWbk.getCellByPosition(7,k).String,C,"'"),1,200))
'*******'
      xlsStrs = xlsStrs+1
      k=k+1
      n=n+1
    wend
    Document.Dispose()
    SET xlWbk = Nothing
    SET sheets = Nothing
    SET Document = Nothing
    SET Desktop = Nothing
    SET xlglob = Nothing
  '  FSO.DeleteFile FDir&FF.Name		' ������� ������������ ����
    xlsFiles = xlsFiles+1
  end if
Next

 ' ��������� �����
 'ar1=Split(MonthYear,".")
 'select case CInt(ar1(0))
 ' case 1 MonthYear=" ������"
 ' case 2 MonthYear=" �������"
 ' case 3 MonthYear=" ����"
 ' case 4 MonthYear=" ������"
 ' case 5 MonthYear=" ���"
 ' case 6 MonthYear=" ����"
 ' case 7 MonthYear=" ����"
 ' case 8 MonthYear=" ������"
 ' case 9 MonthYear=" ��������"
 ' case 10 MonthYear=" �������"
 ' case 11 MonthYear=" ������"
 ' case 12 MonthYear=" �������"
 'end select
 
 'Call oSheet.getCellByPosition(6, 6).SetFormula(MonthYear)
 MonthYear=" "&LCase(MonthYear)&" ����� "
 'if Len(ar1(1))=2 then
 '  MonthYear=MonthYear&"20"
 'end if
 MonthYear=MonthYear&Year(Now)&" �."
 Call oSheet.getCellByPosition(6, 6).SetFormula(MonthYear)
 '-------------------------------------------------------------
'���������� ���������� ���� ����� "osmorStyle" ��� ��������������
'��������� ����� "K1:L10"
'������ �� ��������� �������� �� ����� ������� getCellRangeByName
'-------------------------------------------------------------   
    Set oCells = oSheet.getCellRangeByName("A1:L111")
    'Set oMyStyle = oBook.createInstance("com.sun.star.style.CellStyle")
    'Call oBook.getStyleFamilies().getByName("CellStyles").insertByName("osmorStyle", oMyStyle)
   ' oMyStyle.CellBackColor = RGB(255, 220, 220) ' ���� ����
    'oMyStyle.IsCellBackgroundTransparent = False
   ' oMyStyle.CharColor = RGB(0, 0, 200) ' ����  ������
    'oMyStyle.CharWeight = 150 ' ������� ������
    'Set oCells = oSheet.getCellRangeByName("A1:L111")
    'oCells.CellStyle = "osmorStyle" ' ��������� ����� � ���������� ���������
    oCells.IsTextWrapped = True ' ���������� �� ������
    'Set oMyStyle = Nothing

if xlsStrs=0 then
  MsgBox "���������� " & xlsFiles & " ������!"
else
  MsgBox "���������� " & xlsFiles & " ������! �������� " & xlsStrs & " �������!" & Chr(13) & "��������� ���� � ��������!"    ' ����� ��������� � ����������
end if  
