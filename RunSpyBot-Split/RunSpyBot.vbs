'##########################################
'# Copyright 2006 by N-able Technologies  #
'# All Rights Reserved                    #
'# May not be reproduced or redistributed #
'# Without written consent from N-able    #
'# Technologies			          #
'# www.n-able.com			  #
'##########################################



' Script Name: RunSpyBot.vbs
' Description: This VBS file will run the spybotsd.exe program with the following switches:
'
'		/taskbarhide - this runs the program invisibly
'		/autoupdate - runs autoupdate as soon as the program starts
'		/autocheck - starts scanning immediately	
'		/autofix - automatically fixes problems that were detected
' 		/autoclose - closes the program once it's finished	



'Run the command
Set Shell = CreateObject("Wscript.Shell")
cmdline = "" & Chr(34) & "C:\Program Files\Spybot - Search & Destroy\spybotsd.exe" & Chr(34) & " /taskbarhide /autocheck /autofix /autoclose"
Shell.Run cmdline,7,TRUE
Set Shell = Nothing