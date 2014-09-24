' VB Script Document
option explicit
dim Shell

'Create a shell, and run the 'eventcreate' command
Set Shell = CreateObject("Wscript.Shell")
Shell.Run "%COMSPEC% /c eventcreate /t ERROR /id 999 /l APPLICATION /so EVENTSOURCE /d " & chr(34) & "This is my first Event Log entry."& chr(34),7,TRUE
Set Shell = Nothing


' Hint: If you're curious about what the "7,TRUE" options at the end of the SHell.Run command mean, check out http://msdn.microsoft.com/en-us/library/d5fk67ky%28VS.85%29.aspx
