HKEY_LOCAL_MACHINE = &H80000002

Err.Clear 

On Error Resume Next 
Set objWMIService = GetObject("winmgmts:\\localhost\root\CIMV2") 

Set objReg = GetObject("winmgmts:\\localhost\root\default:StdRegProv")

strKeyPath = "SYSTEM\CurrentControlSet\Control\StorageDevicePolicies"

objReg.CreateKey HKEY_LOCAL_MACHINE, strKeyPath

ValueName = "WriteProtect"

DwordValue = "1"

objReg.SetDwordValue HKEY_LOCAL_MACHINE, strKeyPath, ValueName, DwordValue

