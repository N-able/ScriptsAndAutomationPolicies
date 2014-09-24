' On Error Resume Next

HKEY_LOCAL_MACHINE = &H80000002

On Error Resume Next 
Set objWMIService = GetObject("winmgmts:\\localhost\root\CIMV2") 
 

dim objNetwork 
Dim fso 
Dim CurrentDate 
Dim LogFile 
CurrentDate = Now 
Set objNetwork = WScript.CreateObject("WScript.Network") 
Set fso = CreateObject("Scripting.FileSystemObject") 
strUser = objNetwork.UserDomain

Set objReg = GetObject("winmgmts:\\localhost\root\default:StdRegProv")

strKeyPath = "SYSTEM\CurrentControlSet\Control\StorageDevicePolicies"

objReg.CreateKey HKEY_LOCAL_MACHINE, strKeyPath

ValueName = "WriteProtect"

DwordValue = "0"

objReg.SetDwordValue HKEY_LOCAL_MACHINE, strKeyPath, ValueName, DwordValue


