Const DirDBF= "C:\braki\1\"
Const NameDBF="P1.dbf"


Set dbfConn = CreateObject("ADODB.Connection")
With dbfConn
'  .Provider = "Microsoft.Jet.OLEDB.4.0"
  .Provider = "Microsoft.ACE.OLEDB.12.0"
  .Properties("Extended Properties") = "dBase IV"
  .Properties("Data Source") = DirDBF
  .Open
End With
  
Set dbfPrice = CreateObject("ADODB.Recordset")
Set dbfPrice.ActiveConnection = dbfConn


dbfPrice.Open "SELECT * FROM ["&NameDBF&"]"

msgbox dbfPrice.Fields("NAMETOV")

dbfPrice.Close
