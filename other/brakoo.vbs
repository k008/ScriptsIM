'Option Explicit
Dim FSO,C,FDir,FLD,FL,FF,Sh,FDirOut,dbfConn,dbfRS,LetLab,Letdate,LetNum,xlsFiles,xlsStrs,x1,x2
Dim xlglob,Desktop,Document,sheets,xlWbk
Dim Mass()

FDir="C:\braki\2010\"      ' ����, ��� �������� ����� � ������� Excel
FDirOut="X:\brak\"   ' ���� ���� ����������� ���� � �������

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files

Set dbfConn = CreateObject("ADODB.Connection")
  With dbfConn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Extended Properties") = "dBase IV"
    .Properties("Data Source") = FDirOut
    .Open
  End With
Set dbfRS = CreateObject("ADODB.Recordset")
Set dbfRS.ActiveConnection = dbfConn

if FSO.FileExists(FDir&"brak.dbf") then    ' ���� ���� ����������
  FSO.DeleteFile FDir&"brak.dbf"           ' �� ������� ���
end if
C=Chr(34) ' ������� ������� ��� �����

dbfRS.Open "CREATE TABLE brak(MNFGRNX INT, MNFNX INT, MNFNMR VARCHAR(200), COUNTRYR VARCHAR(200), DRUGTXT VARCHAR(200)," &_
           "  SERNM VARCHAR(200), LETTERSNR VARCHAR(200), LETTERSDT VARCHAR(20), LABNMR VARCHAR(200), QUALNMR VARCHAR(250), " &_
           "TRADENMNX INT, INNNX INT)"     ' ������� ���� DBF

xlsFiles=0
xlsStrs=0	   
For Each FF in FL
  if InStr(LCase(FF.Name),".xls") then
    Set xlglob = CreateObject("com.sun.star.ServiceManager") 
    Set Desktop = xlglob.createInstance("com.sun.star.frame.Desktop")
    Set Document = Desktop.LoadComponentFromURL("file:///"&FDir&FF.Name, "_blank", 0, mass )
    Set sheets = Document.getSheets()
    Set xlWbk = sheets.getByIndex(0)
    k=0
    while xlWbk.getCellByPosition(8,k).String<>"���������"
      k=k+1
    wend
    k=k+1
    'while Len(Trim(xlWbk.getCellByPosition(0,k).String))>0       ' ���� ���������� ������ ������ ������� ������ ��������, ����� ������
     while Len(Trim(xlWbk.getCellByPosition(2,k).String))>0       ' ���� ���������� ������ ������ ������� ������ ��������, ����� ������
    
    ' �������� �������
'		   MNFGRNX		' ��������� ����
'         		   MNFNX		                   ' ��������� ����
'		   MNFNMR		' �������������
'		   COUNTRYR		' ������ �������������
'		   DRUGTXT		' ������������ ���������
'                                          SERNM		                    ' �����
'		   LETTERSNR		' ����� ������
'		   LETTERSDT_		' ���� ������
'		   LABNMR		' �����������
'		   QUALNMR		' �������� ����������
'	                      TRADENMNX		' ��������� ����
'		   INNNX		                    ' ��������� ����
      if InStr(xlWbk.getCellByPosition(8,k).String,":")>0 then
        LetLab = Mid(Replace(xlWbk.getCellByPosition(8,k).String,C,"'"),1,InStr(xlWbk.getCellByPosition(8,k).String,":")-1)
      else
        LetLab = ""
      end if
      LetNum = Mid(Replace(xlWbk.getCellByPosition(8,k).String,C,"'"),InStr(xlWbk.getCellByPosition(8,k).String,"�")+2,InStr(xlWbk.getCellByPosition(8,k).String," ��")-InStr(xlWbk.getCellByPosition(8,k).String,"�")-2)
      LetDate = Mid(Replace(xlWbk.getCellByPosition(8,k).String,C,"'"),InStr(xlWbk.getCellByPosition(8,k).String,"�� ")+3,8)
      x1 = Mid(LetDate,1,InStr(LetDate,"."))
      LetDate = Mid(LetDate,InStr(LetDate,".")+1,Len(LetDate))
      x2 = Mid(LetDate,1,InStr(LetDate,"."))
      LetDate = Mid(LetDate,InStr(LetDate,".")+1,Len(LetDate))
      LetDate = x1 & x2 & "20" & LetDate
      dbfRS.Open "INSERT INTO brak(" &_
                   "MNFGRNX," &_
                   "MNFNX," &_
		   "MNFNMR," &_
		   "COUNTRYR," &_
		   "DRUGTXT," &_
                   "SERNM," &_
		   "LETTERSNR," &_
		   "LETTERSDT," &_
		   "LABNMR," &_
		   "QUALNMR," &_
	           "TRADENMNX," &_
		   "INNNX" &_
		   ") Values (" &_
		   "0," &_
		   "0," &_
		   C & Mid(Replace(xlWbk.getCellByPosition(2,k).String,C,"'"),1,200) & C & "," &_
		   C & Mid(Replace(xlWbk.getCellByPosition(3,k).String,C,"'"),1,200) & C & "," &_
		   C & Mid(Replace(xlWbk.getCellByPosition(1,k).String,C,"'"),1,200) & C & "," &_
		   C & Mid(Replace(xlWbk.getCellByPosition(6,k).String,C,"'"),1,200) & C & "," &_
		   C & LetNum & C & "," &_
		   C & LetDate & C & "," &_
		   C & LetLab & C & "," &_
		   C & Mid(Replace(xlWbk.getCellByPosition(7,k).String,C,"'"),1,250) & C & "," &_
		   "0," &_
		   "0" &_
		   ")"
      xlsStrs = xlsStrs+1
      k=k+1
    wend
    Document.Dispose()
    SET xlWbk = Nothing
    SET sheets = Nothing
    SET Document = Nothing
    SET Desktop = Nothing
    SET xlglob = Nothing
    FSO.DeleteFile FDir&FF.Name		' ������� ������������ ����
    xlsFiles = xlsFiles+1
  end if
Next

if xlsStrs=0 then
  MsgBox "���������� " & xlsFiles & " ������!"
else
  MsgBox "���������� " & xlsFiles & " ������! �������� " & xlsStrs & " �������!" & Chr(13) & "������������ ����� � �-������+!"    ' ����� ��������� � ����������
end if  

dbfConn.Close
dbfRS = null
dbfConn = null
