'*************************************************************************************
'******* Script: Reboot
'******* Description: Creates a scheduled task to reboot the target device
'******* Notes:  As per http://msdn.microsoft.com/en-us/library/bb736357(VS.85).aspx,
'                 this script will only run on XP, Vista, Windows 2003 and Windows 2008. 
'******* Author: Chris Reid
'******* Date: Oct. 21st, 2008
'*************************************************************************************


' Let's declare some variables
Option Explicit
dim timenow, Shell, cmd1, cmd2, cmd3 

' Let's format the timestring so that the 'schtasks' command can understand it
timenow = FormatDateTime(Time(), vbShortTime)


'Let's make sure that there's not already a scheduled task called 'RebootThisPC'
Set Shell = CreateObject("Wscript.Shell")
cmd1 = "%COMSPEC% /c schtasks /delete /tn RebootThisPC /f"
Shell.Run cmd1,7,TRUE
Set Shell = Nothing


' Let's create the scheduled task to reboot the machine 'NOW'
Set Shell = CreateObject("Wscript.Shell")
cmd2 = "%COMSPEC% /c schtasks /create /RU " & chr(34) & "NT AUTHORITY\SYSTEM" & chr(34) & " /SC ONCE /TN RebootThisPC /TR " & chr(34) & "shutdown -r -f -t 60" & chr(34) & " /st " & timenow & ":59"
Shell.Run cmd2,7,TRUE
Set Shell = Nothing

' Finally - let's run the scheduled task 
Set Shell = CreateObject("Wscript.Shell")
cmd3 = "%COMSPEC% /c schtasks /run /TN RebootThisPC"
Shell.Run cmd3,7,TRUE
Set Shell = Nothing
