' VB Script Document
option explicit
dim Shell, strKeyPath




'Define the registry path and value to be created
strKeyPath = "HKEY_CURRENT_USER\SOFTWARE\N-able Technologies\"

'Populate the new key with the appropriate value
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.RegDelete strKeyPath 

' For a comprehensive breakdown of all the fun things you can do with the RegDelete function, check out this Microsoft KB article:
' http://msdn.microsoft.com/en-us/library/293bt9hh%28VS.85%29.aspx
 



