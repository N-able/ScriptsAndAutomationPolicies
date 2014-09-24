' VB Script Document
option explicit
dim oReg, strComputer, strKeyPath, strKeyName, strKeyValue
const HKEY_LOCAL_MACHINE = &H80000002



'Define the registry path and value to be grabbed
strComputer = "." 'This means 'query the machine the script is being run on
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strKeyName = "ProductName"

'Run the command
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strKeyPath,strKeyName,strKeyValue


'Echo the command to the screen
wscript.echo "Your registry value (in this case, the version of Windows you're running) is: " & strKeyValue
