' VB Script Document
option explicit
dim RegPath, RegKey, RegValue, RegType, Shell


'Configure what registry value you want to write
RegPath = "HKLM\Software\N-able Technologies\" 'Make sure that this ends with a \
RegKey = "Registry Test Value"
RegValue = "1"
RegType = "REG_DWORD"

'Set the registry to allow an uninstall
Set Shell = CreateObject("Wscript.Shell")
Shell.RegWrite"" & RegPath & RegKey,"" & RegValue, "" & RegType
Set Shell = Nothing