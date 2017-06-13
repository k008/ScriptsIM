Option Explicit 
On Error Resume Next
Dim strComputer
Dim objWmiService
Dim wmiNS
Dim wmiQuery
Dim objItem
Dim colItems
 
strComputer = "."
wmiNS = "\root\cimv2"
wmiQuery = "Select * from Win32_QuickFixEngineering"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS)
Set colItems = objWMIService.ExecQuery(wmiQuery)
 
For Each objItem in colItems
'    Wscript.Echo "Caption: " & objItem.Caption
'    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FixComments: " & objItem.FixComments
    Wscript.Echo "HotFixID: " & objItem.HotFixID
'    Wscript.Echo "InstallDate: " & objItem.InstallDate
'    Wscript.Echo "InstalledBy: " & objItem.InstalledBy
'    Wscript.Echo "InstalledOn: " & objItem.InstalledOn
'    Wscript.Echo "Name: " & objItem.Name
'    Wscript.Echo "ServicePackInEffect: " & objItem.ServicePackInEffect
'    Wscript.Echo "Status: " & objItem.Status
Next