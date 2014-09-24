'##########################################
'# Copyright 2008 by N-able Technologies  #
'# All Rights Reserved                    #
'# May not be reproduced or redistributed #
'# Without written consent from N-able    #
'# Technologies			          #
'# www.n-able.com			  #
'##########################################


' Script Name: CCCleaner.vbs
' Description: This Visual Basic script will run the CCCleaner application in silent mode.


'Run the command
Set Shell = CreateObject("Wscript.Shell")
cmdline = "" & chr(34) & "C:\Program Files\CCleaner\CCleaner.exe" & chr(34) & " /AUTO"
Shell.Run cmdline,7,TRUE
Set Shell = Nothing