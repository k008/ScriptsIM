Option explicit
On Error Resume Next
Dim FSO,Sep,FDir,FLD,ArrayInp,FF,FL,St,i
Dim DeviceInp,DeviceOut,ArrayOut
Dim TableName,dbfRS,dbfSum,dbfConn


ReDim ConvTable(1)
Const TF="128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,240,241,242,243,244,245,246,247,248,249,250,251,252,253,254,255"
Const TT="192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,63,63,63,166,63,63,63,63,63,63,63,63,63,63,63,172,63,63,63,63,63,134,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,63,135,63,63,63,63,63,63,63,240,241,242,243,244,245,246,247,248,249,250,251,252,253,254,255,168,184,170,186,175,191,161,162,176,149,183,63,185,164,152,160"
Const InExt="dbf"
Const OutExt="sml"
Const OutPath="C:\Mail\Programs\In\"

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
  

  
  
  Sep = ";"

  dbfRS.Open "SELECT * FROM "&TableName
  dbfSum.Open "Select Sum(SUMMA) AS SUMMADOC from "&TableName

  DeviceOut.WriteLine "[Header]"
  DeviceOut.WriteLine Convert866to1251(DelZero(dbfRS.Fields("NOM_DOC"))) & Sep & dbfRS.Fields("DATE_DOC") & Sep & dbfSum.Fields("SUMMADOC")
  DeviceOut.WriteLine "[Body]"
  
    Do While Not dbfRS.Eof

     
     ArrayOut(0)=DelZero(Trim(dbfRS.Fields("CODE")))    ' Код товара
     ArrayOut(1)=Trim(dbfRS.Fields("NAME_TOVAR"))    ' Наименование товара (поставщика)
     ArrayOut(2)=Trim(dbfRS.Fields("PROIZ"))    ' Производитель
     ArrayOut(3)=Trim(dbfRS.Fields("COUNTRY"))             ' Страна
     ArrayOut(4)=Trim(dbfRS.Fields("VOLUME"))   ' Кол-во
     ArrayOut(5)=Trim(dbfRS.Fields("PRICE"))   ' Цена зак.
     ArrayOut(6)=Trim(dbfRS.Fields("PRICE_PROI"))    ' Цена произв.
     ArrayOut(7)=Trim(dbfRS.Fields("PRICEWNDS"))   ' Цена зак. без НДС
     ArrayOut(8)=""   
     ArrayOut(9)=""             ' Наценка посредника
     ArrayOut(10)=""            
     ArrayOut(11)=Trim(dbfRS.Fields("GTD"))  ' ГТД
     ArrayOut(12)=Trim(dbfRS.Fields("DOCUMENT"))&"^"&Trim(dbfRS.Fields("VYDAN"))&"^"&Trim(dbfRS.Fields("END_D"))  ' Сертификаты (серия^рег.номер^дата выдачи и орган^срок действия)
     ArrayOut(13)=Trim(dbfRS.Fields("SERIA"))  ' Серия
     ArrayOut(14)=""            
     ArrayOut(15)=Trim(dbfRS.Fields("SROK"))  ' Срок годности (дата истечения)
     ArrayOut(16)=""            ' Заводской штрих-код
     ArrayOut(17)=""            ' Дата регистрации
     ArrayOut(18)=""            ' Цена реестра
     ArrayOut(19)=""            ' Торг.наценка импортера
     ArrayOut(20)=Trim(dbfRS.Fields("SUMMA"))  ' Сумма по строке
     ArrayOut(21)=""            ' Признак ЖВЛС

     St="" ' Join не хотел работать???
     for i=0 to 21
      St=St&ArrayOut(i)&Sep
     next
     
     dbfRS.MoveNext()
          if Err.Number then Exit Do
          St=Convert866to1251(St)
          DeviceOut.WriteLine(St) 
  Loop
  DeviceOut.close
  'если возникли ошибки то удалим созданный файл
   if Err.Number then
   ' FSO.DeleteFile (FDir&"\"&TableName&"."&OutExt)
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

'------перекодировка из Dos в Windows
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

Function DelZero(St)
Dim i,Stmp
For i=1 to Len(St) 
 Stmp=Mid(St,i,1)
 if Stmp<>"0" then exit for
next 
DelZero=Mid(St,i,Len(St))
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
