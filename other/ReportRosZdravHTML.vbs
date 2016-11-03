'Option Explicit
Dim FSO,C,FDir,LetDate,LetNum,xlsFiles,xlsStrs,x1,x2,MonthYear,ar1,DevInp,St,Tovar,Proizv,StopFlag
Dim xlglob,Desktop,Document,sheets,xlWbk
Dim Mass()
Dim aNoArgs()
Dim oMyStyle
Dim OApplication, ODesktop, ODocument, srcFile, srcProps() ' ******************

Dim args(0)
'oPropertyValue.Name = "AsTemplate"
'oPropertyValue.Value = true

TemplateFile="C:\Braki\ReportIMN1.ots" ' ���� � �������� ������
HTMLReport="C:\Braki\IMN\���.htm" ' HTML-���� � ��������

Set FSO = CreateObject("Scripting.FileSystemObject")

C=Chr(34) ' ������� ������� ��� �����
StopFlag=0

MonthYear=InputBox("������� �����. ������: 06.2013 - ���� 2013�.","������� �����",Mid(CStr(Date),4,2)&"."&Mid(CStr(Date),7,4))

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

xlsFiles=1
xlsStrs=0
n=0
startcol=10	   

Set DevInp = FSO.OpenTextFile(HTMLReport)
' ���� ������ �������
Do While not DevInp.AtEndOfStream
 St=DevInp.ReadLine
 if InStr(LCase(St),"<tbody>") then
   n=n+1
  end if
  if n=2 then
    exit do
  end if
Loop

n=0

Do While not DevInp.AtEndOfStream
 Do While not DevInp.AtEndOfStream
  St=DevInp.ReadLine
  if InStr(LCase(St),"</span>") then
   exit do
  end if
 Loop
 LetNum=Mid(St,InStr(St,"8pt")+5,InStr(St,"</SPAN>")-InStr(St,"8pt")-5)
 
 Do While not DevInp.AtEndOfStream
  St=DevInp.ReadLine
  if InStr(LCase(St),"</span>") then
   exit do
  end if
 Loop
 LetDate=Mid(St,InStr(St,"8pt")+5,InStr(St,"</SPAN>")-InStr(St,"8pt")-5)
 
 Do While not DevInp.AtEndOfStream
  St=DevInp.ReadLine
  if InStr(LCase(St),"title=") then
   exit do
  end if
 Loop
 Tovar=St
 Do While not (InStr(LCase(Tovar),"><nobr>")>0)
  St=DevInp.ReadLine
  Tovar=Tovar & St
 Loop
 St = Tovar
 Tovar=Mid(St,InStr(St,"title=")+7,InStr(St,"><NOBR>")-InStr(St,"title=")-8)
 
 Do While not DevInp.AtEndOfStream
  St=DevInp.ReadLine
  if InStr(LCase(St),"title=") then
   exit do
  end if
 Loop
 Proizv=St
 Do While not (InStr(LCase(Proizv),"><nobr>")>0)
  St=DevInp.ReadLine
  Proizv=Proizv & St
 Loop
 St = Proizv
 Proizv=Mid(St,InStr(St,"title=")+7,InStr(St,"><NOBR>")-InStr(St,"title=")-8)
 
 Do While not DevInp.AtEndOfStream
  St=DevInp.ReadLine
  if InStr(LCase(St),"</td></tr>") then
   if InStr(LCase(St),"</tbody") then
    StopFlag=1
   end if
   exit do
  end if
 Loop
 
 if InStr(LetDate,MonthYear)>0 then
  Call oSheet.getCellByPosition(0, startcol+n).SetFormula(1+n)
  Call oSheet.getCellByPosition(1, startcol+n).SetFormula(LetNum & " �� " & LetDate)
  Call oSheet.getCellByPosition(2, startcol+n).SetFormula(Tovar)
  Call oSheet.getCellByPosition(3, startcol+n).SetFormula("")
  Call oSheet.getCellByPosition(4, startcol+n).SetFormula(Proizv)
  Call oSheet.getCellByPosition(5, startcol+n).SetFormula("�� ������������ �.�., ������ 1, �.������, 5")
  Call oSheet.getCellByPosition(7, startcol+n).SetFormula("0")
  Call oSheet.getCellByPosition(8, startcol+n).SetFormula("�� ��������")
  Call oSheet.Rows.insertByIndex(startcol+n+1, 1)
  xlsStrs = xlsStrs+1
  k=k+1
  n=n+1
 end if
 if Stopflag then
  exit do
 end if
Loop

if xlsStrs=0 then
  MsgBox "���������� " & xlsFiles & " ������!"
else
  MsgBox "���������� " & xlsFiles & " ������! �������� " & xlsStrs & " �������!" & Chr(13) & "��������� ���� �� ������: fs@rznkk.org"    ' ����� ��������� � ����������
end if  
