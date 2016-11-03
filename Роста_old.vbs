Option Explicit
On Error Resume Next
Dim DeviceOut, DeviceInp, FSO, FName, FDir, FLD, FF, FL, FN, S, NN,i,j,ZV,Doc,Docdat
Dim ArrayInp,ArrayOut, St
ReDim ArrayOut(21)
Const InExt="sst"
Const OutExt="sst"
Const OutPath="X:\Programs\In\"


Set FSO = CreateObject("Scripting.FileSystemObject")
'FDir=GetParm()
FDir="C:\Mail\Invoice\"&Mid(WScript.ScriptName,1,InStr(LCase(WScript.ScriptName),".vbs")-1)&"\"
ArrayInp=Split(FDir,";")
FDir=ArrayInp(0)
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files

For Each FF in FL

 If FSO.GetExtensionName(LCase(FF.Name))=InExt then
  i=0
  FName=FDir&"\"&FF.Name
    Set DeviceOut = FSO.CreateTextFile(FDir&"\"&Replace(FF.Name,".","_")&"pr."&OutExt, True)
    Set DeviceInp = FSO.OpenTextFile(FName)
    St=DeviceInp.ReadLine
	DeviceOut.WriteLine(St)
    St=DeviceInp.ReadLine
	DeviceOut.WriteLine(St)
    St=DeviceInp.ReadLine
	DeviceOut.WriteLine(St)
	
    Do While not DeviceInp.AtEndOfStream     ' товарные строки
     St=DeviceInp.ReadLine
	 if St="" then Exit Do
	 ArrayInp=Split(St,";")
     ArrayOut(0)=ArrayInp(0)					' Код товара 
     ArrayOut(1)=ArrayInp(1)		' Наименование товара (поставщика)
     ArrayOut(2)=ArrayInp(2)		' Производитель
     ArrayOut(3)=ArrayInp(3)						' Страна
     ArrayOut(4)=ArrayInp(4)		' Кол-во
     ArrayOut(5)=ArrayInp(5)		' Цена зак. ????????????
     ArrayOut(6)=ArrayInp(6)		' Цена произв.
     ArrayOut(7)=ArrayInp(7)		' Цена зак. без НДС
     ArrayOut(8)=ArrayInp(8)
     ArrayOut(9)=ArrayInp(9)		' Наценка посредника
     ArrayOut(10)=ArrayInp(10)					' Резерв
     ArrayOut(11)=ArrayInp(11)					' ГТД
     ArrayOut(12)=ArrayInp(12)	'сертификаты
     ArrayOut(13)=ArrayInp(13)		' Серия
     ArrayOut(14)=ArrayInp(14)					' резерв
     ArrayOut(15)=ArrayInp(15)		' Срок годности (дата истечения)
     ArrayOut(16)=Trim(ArrayInp(22))					' Заводской штрих-код
     ArrayOut(17)=ArrayInp(17)	' Дата регистрации
     ArrayOut(18)=ArrayInp(18)	' Цена реестра
     ArrayOut(19)=ArrayInp(19)					' Торг.наценка импортера
     ArrayOut(20)=ArrayInp(20)     ' Сумма по строке
     ArrayOut(21)=ArrayInp(21)					' Признак ЖВЛС
     
     DeviceOut.WriteLine(Join(ArrayOut,";"))
     if Err.Number then Exit Do      
   Loop
   
   DeviceOut.Close: Set DeviceOut = nothing
   DeviceInp.Close: Set DeviceInp = Nothing
   FSO.DeleteFile (FName)
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
 End If
Next
CopyFiles()
Set FL = nothing
Set FLD = nothing
Set FSO = nothing

Function GetParm()
  Set DeviceInp = FSO.OpenTextFile("mail.tmp")
  GetParm=DeviceInp.ReadLine()
  DeviceInp.Close: Set DeviceInp = nothing
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
