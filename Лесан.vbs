Option explicit
'On Error Resume Next
Dim FSO,Sep,FDir,FLD,ArrayInp,FF,FL,St
Dim DeviceInp,DeviceOut
Dim TableName,dbfPrice,dbfSum,dbfConn
Const InExt="dbf"
Const OutExt="sas"
Const OutPath="X:\Programs\In\"
Set FSO = CreateObject("Scripting.FileSystemObject")
FDir="C:\Mail\Invoice\"&Mid(WScript.ScriptName,1,InStr(LCase(WScript.ScriptName),".vbs")-1)&"\"
ArrayInp=Split(FDir,";")
FDir=ArrayInp(0)
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files

For Each FF in FL 'Смотрим все файлы и убираем год
 if InStr(LCase(FF.Name),"."&InExt) then
   if InStr(FF.Name,Year(Now)) then
     FSO.CopyFile FDir&FF.Name, FDir&Mid(FF.Name,InStr(FF.Name,Year(Now))+Len(Year(Now)),Len(FF.Name))
     FSO.DeleteFile FDir&FF.Name
   end if
 end if
next
Set FLD = Nothing
SET FL = Nothing
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files

For Each FF in FL
 if InStr(LCase(FF.Name),"."&InExt) AND (FF.Size<50000) then

  TableName=Mid(LCase(FF.Name),1,InStr(LCase(FF.Name),"."&InExt)-1)
  Set dbfConn = CreateObject("ADODB.Connection")
  
  With dbfConn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Extended Properties") = "dBase IV"
    .Properties("Data Source") = FDir
    .Open
  End With

  
  Set dbfPrice = CreateObject("ADODB.Recordset")
  Set dbfPrice.ActiveConnection = dbfConn
  Set dbfSum = CreateObject("ADODB.Recordset")
  Set dbfSum.ActiveConnection = dbfConn

  
  
  Sep = ";"
  
  dbfPrice.Open "SELECT * FROM ["&TableName&"]"
  dbfSum.Open "SELECT Sum(SUMMA) AS Summ FROM ["&TableName&"]"
  ' St= Year(Now)&TableName&Sep & Date & Sep & dbfSum("Summ")
  St= dbfPrice.Fields("NUMDOC")&Sep & dbfPrice.Fields("DATEDOC") & Sep & dbfSum("Summ")
  
   Set DeviceOut = FSO.CreateTextFile(FDir&"\"&TableName&"."&OutExt, True)
   DeviceOut.WriteLine "[Header]"
   DeviceOut.WriteLine(St)
   DeviceOut.WriteLine "[Body]"
   
    Do While Not dbfPrice.Eof
       St=DecodeCode(Trim(dbfPrice.Fields("KOD"))) & Sep & Trim(dbfPrice.Fields("NAME")) & Sep & Trim(dbfPrice.Fields("PROIZV")) & Sep & Trim(dbfPrice.Fields("COUNTRY")) & Sep &_
  	    dbfPrice.Fields("KOLVO") & Sep & dbfPrice.Fields("SUMMA")/dbfPrice.Fields("KOLVO") & Sep & dbfPrice.Fields("CENAPROIZ") & Sep & dbfPrice.Fields("SUMMA")/dbfPrice.Fields("KOLVO")/(dbfPrice.Fields("NDSPOSTAV")+100)*100 & Sep &_
		Sep & Sep & Sep & Trim(dbfPrice.Fields("N_DECLAR")) & Sep & Trim(dbfPrice.Fields("SERTIF")) & "^" & Trim(dbfPrice.Fields("DATAREGSE")) & " " & Trim(dbfPrice.Fields("DATAREGCR")) & Sep & Trim(dbfPrice.Fields("SERII")) & Sep & Sep & Trim(dbfPrice.Fields("DATAEND")) & Sep & Trim(dbfPrice.Fields("BARCODE")) & Sep & Sep & Trim(dbfPrice.Fields("REESTR")) & Sep & Sep & Trim(dbfPrice.Fields("SUMMA")) & Sep & ""
	  DeviceOut.WriteLine (St)
	  dbfPrice.MoveNext
          if Err.Number then Exit Do
    Loop
     
   dbfConn.Close
   DeviceOut.close
   'если возникли ошибки то удалим созданный файл
   if Err.Number then 
    FSO.DeleteFile (FDir&"\"&TableName&"."&OutExt)
    Dim FOut
    if not FSO.FileExists("error.log") then 
     Set FOut=FSO.CreateTextFile("error.log")
    else          
     Set Fout=FSO.OpenTextFile("error.log",8,true)
    end if
    FOut.WriteLine("["&Now()&"]	"&Err.Description&" ->"&WScript.ScriptName)
    FOut.Close() : FOut=nothing
   end if
   Err.clear()
 end if
next
   
dbfPrice = null
dbfSum = null
dbfConn = null
CopyFiles()
DeviceOut = null
fso = null

Function GetParm()
  Set DeviceInp = FSO.OpenTextFile("mail.tmp")
  GetParm=DeviceInp.ReadLine()
  DeviceInp.Close: Set DeviceInp = nothing
End Function

Function CopyFiles()
Set FL = FLD.Files
For Each FF in FL
  if InStr(LCase(FF.Name),"."&OutExt) then
    FSO.CopyFile FDir&FF.Name, OutPath&FF.Name
    FSO.DeleteFile FDir&FF.Name
  else
    FSO.DeleteFile FDir&FF.Name
  end if
Next
End Function

Function ReplaceStr(S,S1,S2)
  if IsNull(S) then
    ReplaceStr=S
  else
    ReplaceStr=Replace(S,S1,S2)
  end if
End Function

Function DecodeCode(Value)
 Dim Value1,i,St
 Value=ReplaceStr(Value,"-","") ' Убираем тире, если есть
 Value1=""
 For i = 1 to Len(Value)
  St=Mid(Value,i,1)
  if Asc(St)<65 then  ' Если код символа меньше A, то берем его
   Value1=Value1 & St
  end if
  if Asc(St)>90 then  ' Если код символа больше Z, то берем его
   Value1=Value1 & St
  end if
  if ((Asc(St)>=65) and (Asc(St)<=90)) then
   Value1=Value1 & Asc(St)
  end if
 Next
 DecodeCode=Value1
End Function
