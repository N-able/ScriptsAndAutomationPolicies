' VB Script Document
option explicit
dim Shell, strKeyPath, strKeyName, strKeyValue




'Define the registry path and value to be created
strKeyPath = "HKEY_CURRENT_USER\SOFTWARE\N-able Technologies\"
strKeyName = "My First Registry Key"
strKeyValue = "Look! A registry value!"

'Populate the new key with the appropriate value
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.RegWrite strKeyPath & strKeyName, strKeyValue

' For a comprehensive breakdown of all the fun things you can do with the RegWrite function, check out this Microsoft KB article:
' http://msdn.microsoft.com/en-us/library/yfdfhz1b%28VS.85%29.aspx
 



