Option explicit
On Error Resume Next
Dim FSO,Sep,FDir,FLD,ArrayInp,FF,FL,St,i,k,zv
Dim DeviceInp,DeviceOut,ArrayOut
Dim TableName,dbfRS,dbfSum,dbfConn

ReDim ConvTable(1)
Const TF="128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,240,241,242,243,244,245,246,247,248,249,250,251,252,253,254,255"
Const TT="192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,63,63,63,166,63,63,63,63,63,63,63,63,63,63,63,172,63,63,63,63,63,134,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,135,63,63,63,63,63,63,63,240,241,242,243,244,245,246,247,248,249,250,251,252,253,254,255,168,184,170,186,175,191,161,162,176,149,183,63,185,164,152,160"
Const InExt="dbf"
Const OutExt="sun"
Const OutPath="X:\Programs\In\"

Set FSO = CreateObject("Scripting.FileSystemObject")
FDir="C:\Mail\Invoice\"&Mid(WScript.ScriptName,1,InStr(LCase(WScript.ScriptName),".vbs")-1)&"\"
ArrayInp=Split(FDir,";")
FDir=ArrayInp(0)

Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files

 
ReDim ArrayOut(22)

For Each FF in FL
 
 if InStr(LCase(FF.Name),"."&InExt) then
  TableName=Mid(LCase(FF.Name),1,InStr(LCase(FF.Name),"."&InExt)-1)
  Set DeviceOut = FSO.CreateTextFile(FDir&"\"&TableName&"."&OutExt, True)
  Set dbfConn = CreateObject("ADODB.Connection")
  With dbfConn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Extended Properties") = "dBase IV"
    .Properties("Data Source") = FDir
    .Open
  End With 
  
  Set dbfRS = CreateObject("ADODB.Recordset")
  Set dbfRS.ActiveConnection = dbfConn
  Set dbfSum = CreateObject("ADODB.Recordset")
  Set dbfSum.ActiveConnection = dbfConn
  
    
  
  Sep =";"

  dbfRS.Open "SELECT * FROM "&TableName
  k = 1
  dbfSum.Open "SELECT Count(NAME_TV) AS Summ FROM "&TableName
  
  Do While Not dbfRS.Eof
     if k = dbfSum("Summ") then
	Exit Do
     end if
     k = k+1
     dbfRS.MoveNext
  loop   
  dbfSum.Close
  dbfSum.Open "Select Sum ( SUMMA ) AS SUMMADOC1 from "&TableName
  
  DeviceOut.WriteLine "[Header]"
  DeviceOut.WriteLine dbfRS.Fields("NAME_TV") & Sep & Mid(dbfRS.Fields("DATA_GOOD"),7,2) & "." & Mid(dbfRS.Fields("DATA_GOOD"),5,2) & Mid(dbfRS.Fields("DATA_GOOD"),1,4) & Sep & dbfSum.Fields("SUMMADOC1") 
  DeviceOut.WriteLine "[Body]" 
  dbfRS.Close
  k=1
  dbfSum.Close
  dbfSum.Open "SELECT Count(NAME_TV) AS Summ FROM "&TableName
  dbfRS.Open "SELECT * FROM "&TableName  
    Do While Not dbfRS.Eof
     if k = dbfSum("Summ") then
	Exit Do
     end if
     k = k+1 
     if dbfRS.Fields("GV")=True then
       ZV=1
     else
       ZV=0
     end if       
     ArrayOut(0)= Trim(dbfRS.Fields("NAME_TV"))    ' Код товара
     ArrayOut(1)= Trim(dbfRS.Fields("NAME_TV"))    ' Наименование товара (поставщика)
     ArrayOut(2)= Trim(dbfRS.Fields("IZGOT"))     ' Производитель 
     ArrayOut(3)= Trim(dbfRS.Fields("STRANA"))              ' Страна
     ArrayOut(4)= Trim(dbfRS.Fields("KOL"))   ' Кол-во
     ArrayOut(5)= Trim(dbfRS.Fields("CENA_S_NAC"))    ' Цена зак.
     ArrayOut(6)= Trim(dbfRS.Fields("CENA_Z"))     ' Цена произв.
     ArrayOut(7)= Trim(dbfRS.Fields("CENA"))     ' Цена зак. без НДС
     ArrayOut(8)=""    
     ArrayOut(9)= ""        ' Наценка посредника
     ArrayOut(10)=""        
     ArrayOut(11)=tRIM(dbfRS.Fields("GTD"))   ' ГТД
     ArrayOut(12)=Trim(dbfRS.Fields("SER"))&"^"&Trim(dbfRS.Fields("PROT"))&"^"&Trim(dbfRS.Fields("NOM_REG")) 'сертификаты
     ArrayOut(13)=Trim(dbfRS.Fields("SER"))  ' Серия
     ArrayOut(14)=""  'Резерв
     ArrayOut(15)=Trim(dbfRS.Fields("DATA_GOOD"))    ' Срок годности (дата истечения)
     ArrayOut(16)=""           ' Заводской штрих-код
     ArrayOut(17)=Trim(dbfRS.Fields("DATA_REG"))            ' Дата регистрации
     ArrayOut(18)=Trim(dbfRS.Fields("CENA_REE"))            ' Цена реестра
     ArrayOut(19)=""            ' Торг.наценка импортера
     ArrayOut(20)= Trim(dbfRS.Fields("SUMMA"))   ' Сумма по строке
     ArrayOut(21)=ZV              ' Признак ЖВЛС

     St="" ' Join не хотел работать???
     for i=0 to 21
      St=St&ArrayOut(i)&Sep
     next
     St=Convert866to1251(St)
     dbfRS.MoveNext()
     
          if Err.Number then Exit Do
          DeviceOut.WriteLine(St) 
  Loop
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
   
dbfRS = null
dbfSum = null
dbfConn = null
DeviceOut = null
CopyFiles()
fso = null

Function GetParm()
  Set DeviceInp = FSO.OpenTextFile("mail.tmp")
  GetParm=DeviceInp.ReadLine()
  DeviceInp.Close: Set DeviceInp = nothing
End Function

Sub MakeConvTable()
  Dim ArrT,ArrF,i
  ReDim ConvTable(256)
  ArrF=Split(TF,",")
  ArrT=Split(TT,",")
  For i=0 to UBound(ArrF)
    ConvTable(ArrF(i))=Chr(ArrT(i))
  Next
End Sub

Function Convert866to1251(St)
  Dim A,i,Ch, StOut
  StOut=""
  if UBound(ConvTable)=1 then MakeConvTable()
  For i=1 to Len(St)
    Ch=Mid(St,i,1)   
    A=ConvTable(Asc(Ch))
    if A="" then A=Ch
    StOut=StOut&A
  Next 
  Convert866to1251=StOut
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
