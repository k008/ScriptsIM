Option explicit
On Error Resume Next
Dim FSO,Sep,FDir,FLD,ArrayInp,FF,FL,St
Dim DeviceInp,DeviceOut
Dim TableName,ZV,i,xlglob,Desktop,Document,sheets,xlWbk
Dim mass()

Const InExt="xls"
Const OutExt="svb"
Const OutPath="C:\Mail\Programs\In\"

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
  St=xlWbk.getCellByPosition(6,12).String & Sep & xlWbk.getCellByPosition(9,12).String & Sep & ""

  DeviceOut.WriteLine(St)
  DeviceOut.WriteLine "[Body]"
  i=18 
  Do While xlWbk.getCellByPosition(15,i).String<>""
      if InStr(xlWbk.getCellByPosition(15,i).String,"����� �") then
        i=i+3
      end if	
      if Not ((xlWbk.getCellByPosition(13,i).String="�") or InStr(xlWbk.getCellByPosition(15,i).String,"��������")>0) then
        St=xlWbk.getCellByPosition(1,i).String & Sep & xlWbk.getCellByPosition(1,i).String & Sep & Sep & Sep &_
	    xlWbk.getCellByPosition(10,i).String & Sep & xlWbk.getCellByPosition(15,i).String/xlWbk.getCellByPosition(10,i).String & Sep & Sep & xlWbk.getCellByPosition(12,i).String/xlWbk.getCellByPosition(10,i).String & Sep &_
		Sep & Sep & Sep & Sep & Sep & Sep & Sep & Sep &_
		Sep & Sep & Sep & Sep & Replace(xlWbk.getCellByPosition(15,i).String,"�","") & Sep
        i = i+1
        if Err.Number then Exit Do 
        DeviceOut.WriteLine (St)
      else
	i=i+1
      end if
    Loop
  Document.Dispose()
  SET xlWbk = Nothing
  SET sheets = Nothing
  SET Document = Nothing
  SET Desktop = Nothing
  SET xlglob = Nothing
  DeviceOut.close
  '���� �������� ������ �� ������ ��������� ����
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
