Option explicit
Dim dbfConn, dbfPrice, FSO, DirKor, DirSKM, Elenka1C, Eva1C, MAXI1C, NAINA1C, Svetlana1C, Elenka, Eva, MAXI, NAINA, Svetlana

Const DirDBF="\\MAGISTR\MAExport\"
Const NameDBF="1CDocs.dbf"
 DirKor=DirDBF & "KOR\"
 DirSKM=DirDBF & "SKM\"
Const SKM="ÑÊ-ÌוהÔאנל"

Elenka="Elenka"
Eva="Eva"
MAXI="MAXI"
NAINA="NAINA"
Svetlana="Svetlana"

Elenka1C="\\" & Elenka & "\1c\"
Eva1C="\\" & Eva & "\1c\"
MAXI1C="\\" & MAXI & "\1c\"
NAINA1C="\\" & NAINA & "\1c\"
Svetlana1C="\\" & Svetlana & "\1c\"

Set FSO = CreateObject("Scripting.FileSystemObject")
'CheckPath(DirKor)
'CheckPath(DirSKM)
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
'msgbox dbfPrice.Fields("NAMESUB2")
'msgbox (InStr(1,dbfPrice.Fields("NAMESUB2"), SKM))

If InStr(1,dbfPrice.Fields("NAMESUB1"), SKM) Or InStr(1,dbfPrice.Fields("NAMESUB2"), SKM) Then
  'msgbox "SKM=" & dbfPrice.Fields("NAMESUB1") & dbfPrice.Fields("NAMESUB2")
  'Wscript.Echo "SKM" & chr(13)
  If Ping(Elenka) = 0 Then
    If CheckPath(Elenka1C) = 1 Then
      FSO.CopyFile DirDBF&NameDBF, Elenka1C&NameDBF
      'Wscript.Echo "Elenka On"
      'FSO.DeleteFile DirDBF&NameDBF
    End If
  End If
 
  If Ping(Eva) = 0 Then
'   msgbox "ping"
    If CheckPath(Eva1C) = 1 Then
      FSO.CopyFile DirDBF&NameDBF, Eva1C&NameDBF
      'Wscript.Echo "Elenka On"
      'FSO.DeleteFile DirDBF&NameDBF
    End If
  End If

  If Ping(Svetlana) = 0 Then
    If CheckPath(Svetlana1C) = 1 Then
      FSO.CopyFile DirDBF&NameDBF, Svetlana1C&NameDBF
      'Wscript.Echo "Svetlana On"
      'FSO.DeleteFile DirDBF&NameDBF
    End If
  End If

else
  'msgbox "KOR=" & dbfPrice.Fields("NAMESUB1") & dbfPrice.Fields("NAMESUB2")
    'Wscript.Echo "KOR" & chr(13)
  If Ping(MAXI) = 0 Then
    If CheckPath(MAXI1C) = 1 Then
      FSO.CopyFile DirDBF&NameDBF, MAXI1C&NameDBF
      'Wscript.Echo "MAXI On"
      'FSO.DeleteFile DirDBF&NameDBF
    End If
  End If
  
  If Ping(NAINA) = 0 Then
    If CheckPath(NAINA1C) = 1 Then
      FSO.CopyFile DirDBF&NameDBF, NAINA1C&NameDBF
      'Wscript.Echo "NAINA On"
      'FSO.DeleteFile DirDBF&NameDBF
    End If
  End If

End If

msgbox ("OK")

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

Function Ping (strTarget)
	Dim objWMIService, colPings, objPing
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strTarget & "'")
	For Each objPing in colPings
		Ping = objPing.StatusCode
        'Wscript.Echo strTarget & Ping & chr(13)
	Next
End Function
