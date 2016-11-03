Option Explicit
On Error Resume Next
Dim DeviceOut, DeviceInp, FSO, FName, FDir, FLD, FF, FL, FN, S, NN,i,j,ZV,Doc,Docdat
Dim ArrayInp,ArrayOut, St
ReDim ArrayOut(21)
Const InExt="txt"
Const OutExt="syu"
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
   FSO.CopyFile FDir&FF.Name, FDir&Replace(FF.Name,".","_")&"."&OutExt
 End If
Next
CopyFiles()
Set FL = nothing
Set FLD = nothing
Set FSO = nothing

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
