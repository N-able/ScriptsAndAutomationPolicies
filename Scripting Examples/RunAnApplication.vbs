' VB Script Document
option explicit
dim Shell, Application


'Configure what application should be run
Application = "%SystemRoot%\system32\calc.exe"



'Create a shell, and run the command
Set Shell = CreateObject("Wscript.Shell")
Shell.Run "%COMSPEC% /c " & Application,7,TRUE
Set Shell = Nothing


' Hint: If you're curious about what the "7,TRUE" options at the end of the SHell.Run command mean, check out http://msdn.microsoft.com/en-us/library/d5fk67ky%28VS.85%29.aspx