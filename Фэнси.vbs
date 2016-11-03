Option explicit
On Error Resume Next
Dim FSO,Sep,FDir,FLD,ArrayInp,FF,FL,St,St1
Dim DeviceInp,DeviceOut
Dim TableName,ZV,i,xlglob,Desktop,Document,sheets,xlWbk
Dim mass()

Const InExt="xls"
Const OutExt="sfy"
Const OutPath="H:\programs\detstvo\"

Set FSO = CreateObject("Scripting.FileSystemObject")
'FDir=GetParm()
FDir="C:\Mail\Invoice\"&Mid(WScript.ScriptName,1,InStr(LCase(WScript.ScriptName),".vbs")-1)&"\"
ArrayInp=Split(FDir,";")
FDir=ArrayInp(0)
Set FLD = FSO.GetFolder(FDir)
Set FL = FLD.Files

For Each FF in FL
 if InStr(LCase(FF.Name),"."&InExt) then

  TableName=Mid(LCase(FF.Name),1,InStr(LCase(FF.Name),"."&InExt)-1)
  Set DeviceOut = FSO.CreateTextFile(FDir&"\"&TableName&"."&OutExt, True)
  

  DeviceOut.WriteLine "[Header]"
  
  Sep = ";"
  ZV = 0

  Set xlglob = CreateObject("com.sun.star.ServiceManager") 
  Set Desktop = xlglob.createInstance("com.sun.star.frame.Desktop")
  Set Document = Desktop.LoadComponentFromURL("file:///"&FDir&FF.Name, "_blank", 0, mass )
  Set sheets = Document.getSheets()
  Set xlWbk = sheets.getByIndex(0) 
  St=xlWbk.getCellByPosition(1,0).String
  St1=Trim(Mid(St,InStr(St,"№")+1,InStr(St,"от")-InStr(St,"№")-1)) & Sep & Trim(Mid(St,InStr(St,"от")+3,10)) & Sep

  DeviceOut.WriteLine(St1)
  DeviceOut.WriteLine "[Body]"
  i=10 
  Do While xlWbk.getCellByPosition(29,i).String<>""
	
        St=xlWbk.getCellByPosition(3,i).String & Sep & Replace(Replace(xlWbk.getCellByPosition(7,i).String,chr(10),""),chr(13),"") & Sep &  Sep &_
	    Replace(xlWbk.getCellByPosition(33,i).String," ","") & Sep & Replace(xlWbk.getCellByPosition(20,i).String," ","") & Sep & Replace(xlWbk.getCellByPosition(25,i).String," ","") & Sep & Sep & xlWbk.getCellByPosition(25,i).String*100/110 & Sep &_
		Sep & "0" & Sep & Sep & Replace(xlWbk.getCellByPosition(37,i).String," ","") & Sep & Sep & Sep & Sep & Sep &_
		Replace(xlWbk.getCellByPosition(41,i).String," ","") & Sep & Sep & Sep & Sep & Replace(xlWbk.getCellByPosition(29,i).String," ","") & Sep
        i = i+1
        if Err.Number then Exit Do 
        DeviceOut.WriteLine (St)
    Loop
  Document.Dispose()
  SET xlWbk = Nothing
  SET sheets = Nothing
  SET Document = Nothing
  SET Desktop = Nothing
  SET xlglob = Nothing
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
   
DeviceOut = null
CopyFiles()
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
