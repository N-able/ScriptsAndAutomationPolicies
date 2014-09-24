'##########################################
'# Copyright 2008 by N-able Technologies  #
'# All Rights Reserved                    #
'# May not be reproduced or redistributed #
'# Without written consent from N-able    #
'# Technologies			          #
'# www.n-able.com			  #
'##########################################


' Script Name: DeleteTempFiles.vbs
' Description: This Visual Basic script will delete all of the Windows Temp files (found in the %temp% directory) on a device


'Run the command
Set Shell = CreateObject("Wscript.Shell")
cmdline = "del /q /f /s %temp%\*"
Shell.Run cmdline,7,TRUE
Set Shell = Nothing