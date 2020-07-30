'##########################################
'# Copyright 2008 by N-able Technologies  #
'# All Rights Reserved                    #
'# May not be reproduced or redistributed #
'# Without written consent from N-able    #
'# Technologies			                      #
'# www.n-able.com			                    #
'##########################################


' Script Name: UninstallTrend.vbs
' Description: This Visual Basic script will uninstall Trend Micro Internet Security 2008.


'Run the command
Set Shell = CreateObject("Wscript.Shell")
cmdline = "msiexec.exe /quiet /X {A621B45A-D138-4A95-BE10-7CABA05EF94E}"
Shell.Run cmdline,7,TRUE
Set Shell = Nothing