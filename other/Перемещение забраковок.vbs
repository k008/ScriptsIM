Option explicit
Dim dbfConn, dbfPrice, FSO, DirKor, DirSKM, iPing

Const DirDBF = "X:\PRnakl\" '"C:\braki\1\"
Const NameDBF = "Prichod.dbf"
Const sHoldingSwapIP = "129.186.1.24"

DirKor="\\" & sHoldingSwapIP & "\HoldingSwap\05 �� ������������ �.�\����������\"
DirSKM="\\" & sHoldingSwapIP & "\HoldingSwap\11 ��� ��-�������\����������\"
Const SKM="��-�������"

Set FSO = CreateObject("Scripting.FileSystemObject")

Set dbfConn = CreateObject("ADODB.Connection")
  With dbfConn
  .Provider = "Microsoft.Jet.OLEDB.4.0"
  .Properties("Extended Properties") = "dBase IV"
  .Properties("Data Source") = DirDBF
  .Open
End With
Set dbfPrice = CreateObject("ADODB.Recordset")
Set dbfPrice.ActiveConnection = dbfConn
dbfPrice.Open "SELECT * FROM ["&NameDBF&"]"

iPing = Ping(sHoldingSwapIP)

If iPing=0 Then
	'WScript.Echo "�������� ����"
	If InStr(1,dbfPrice.Fields("NAMESUB2"), SKM) Then
		If CheckPath(DirSKM) = 1 Then
			FSO.CopyFile DirDBF&NameDBF, DirSKM&NameDBF
		End If
	else
		If CheckPath(DirKor) = 1 Then
			FSO.CopyFile DirDBF&NameDBF, DirKor&NameDBF
		End If  
	End If
Else
	WScript.Echo "��������� ���, ��� ������: "& iPing
	msgbox "�������� ����-��������� ��������? ���: " & iPing 
End If

dbfPrice.Close

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

msgbox ("OK")