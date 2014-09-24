' *******************************************************************************************************
' *********************************** Stop Agent/Run Tasks/Start Agent **********************************
' *************************************** Written By: Will Smith ****************************************
' *********************************** Copyright: 10-100 Partnership Ltd *********************************
' ********** 10-100 Partnership Ltd presents this script as-is - use at your own discretion *************
' *******************************************************************************************************
'
' This script is designed to create a batch file in your temp folder, add some commands into that batch
' file to stop the Windows Agent & Maintenance service, add any commands you want (anything you can do at
' a command prompt), add lines to start the Windows Agent & Maintenance service. When the batch file has 
' been created, it will be scheduled to run in 2 minutes from now.
' 
' Scheduling can be hard with different date formats, so Ive made allocation for USA and UK date formats
' however, if you know scripting and dont use USA/UK date formats, then you should be able to figure it
' out easy enough.
'
' This could be altered to run and create VBSCript files.. in theory.. but the encapsulation for creating
' a VBScript file, from within a VBScript file would get pretty crazy, however if you really needed to,
' Im sure you could figure it out if you know VBScriping language.
'
' For most people, you will only want to touch the first lines of code in this (CMDLine commands). The 
' rest of the script (unless doing something special) should remain un-altered.
'
' *******************************************************************************************************

Dim CMDLine(30)

' *******************************************************************************************************
' Ive allowed for 30 lines of CMDLines to be created. If you need more then change the value of the 
' "DIM CMDLline(X) number to a larger number. e.g. 
'
' 					Dim CMDLine(100)
'
' Doing this WILL make your script 100 lines long (even if you only use 10 lines) and your start command
' for the agent will be right at the bottom of that file!
'
'			YOU DONT NEED TO ADD A START OR STOP FOR THE AGENT - THIS IS AUTOMATICALLY ADDED.
' *******************************************************************************************************

CMDLine(0) = "Net Stop wauau"
CMDLine(1) = "Ping -n 30 localhost > NUL"
CMDLine(2) = "del %windir\windowsupdate.log"
CMDLine(3) = "rd %windir\SoftwareDistribution\DataStore /s /q"
CMDLine(4) = "md %windir\SoftwareDistribution\DataStore"
CMDLine(5) = "rd %windir\SoftwareDistribution\Download /s /q"
CMDLine(6) = "md %windir\SoftwareDistribution\Download" 
CMDline(7) = "Net Start wauau"

' *******************************************************************************************************
' 			YOU DONT NEED TO ADD A START OR STOP FOR THE AGENT - THIS IS AUTOMATICALLY ADDED.
'
' List out things you want to do inbetween the agent being started and stopped. Please note, these lines 
' are running in a standard DOS batch file, and like any batch file, if you have a long path name (or 
' anything with spaces) you need to encapsulate it in quotes. The thing to remember however, is you are 
' working in VBScript, so you cant just type a set of quotes, it WONT work. To get around this, you need 
' to replace your quotes with:
' 
' 					& Chr(34) &
'
' This programmitcally creates a set of quotes (You dont need an & at the end of your line). To give a 
' real world example, lets take this example:
'
'					CD "C:\Program Files\Software"
'
' To write that you would do so in the following format:
'
'		 	CMDLine(0) = "CD " & Chr(34) & "C:\Program Files\Software" & Chr(34)
'
' This would creates CD (with a space after it) AND appends a double quotation mark AND then adds in 
' C:\Program Files\Software AND then adds another double quotation mark. Annoying, but thats the way it 
' works in VBScript.
'
' 		YOU DONT NEED TO ADD A START OR STOP FOR THE AGENT - THIS IS AUTOMATICALLY ADDED.
'
' If you need to put a pause in your batch file, say for file locks being released, there is nice way to
' do this is to add: 
'
'					"Ping -n 30 localhost > NUL"
'
' This will perform 30 pings (1 per second) to null (wont show on the screen). 
'
' If you dont want to run anything other than stopping/starting the agent, just leave "" in the CMDLine.
' 
'					    	CMDLine(X) = ""
'
' *******************************************************************************************************

' *******************************************************************************************************
' Set the Commands for stopping the agent, pausing to release file locks etc and starting the agent up 
' again. 
' *******************************************************************************************************

AgentStop = "Net Stop " & Chr(34) & "Windows Agent Service" & Chr(34)
AgentMaintStop = "Net Stop " & Chr(34) & "Windows Agent Maintenance Service" & Chr(34)
PauseToReleaseLocks = "Ping -n 30 localhost > NUL"
AgentStart = "Net Start " & Chr(34) & "Windows Agent Service" & Chr(34)
AgentMaintStart = "Net Start " & Chr(34) & "Windows Agent Maintenance Service" & Chr(34)

' *******************************************************************************************************
' Get the %WINDIR%\ and create path variables
' *******************************************************************************************************
Set winsh = CreateObject("WScript.Shell")
Set winenv = winsh.Environment("Process")
WinDir = winenv("WINDIR")
WinTemp = winenv("WINDIR") & "\temp\"

' *******************************************************************************************************
' Set up standard Variables
' *******************************************************************************************************
Dim objFSO, objFolder, objShell, objFile
Dim strDirectory, strFile
Dim computer, delayMinutes, startupScript 
Dim newTime, timeString, dateString
computer = "." 
strDirectory = WinTemp
strFile = "NableJobs.bat"

' *******************************************************************************************************
' Create NableJobs.bat
' *******************************************************************************************************
' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Make sure the temp folder exists or quit with an error if not.
If objFSO.FolderExists(strDirectory) Then
	Set objFolder = objFSO.GetFolder(strDirectory)
Else
	WScript.echo " - \%WINDIR%\TEMP\ Does NOT exist on this machine"
	WScript.quit(1)
End If

'Delete any existing file in there 
If objFSO.FileExists(strDirectory & strFile) Then
	'Set objFolder = objFSO.GetFolder(strDirectory)
	objFso.DeleteFile (strDirectory & strFile),True
	WScript.echo vbcrlf & " - Deleted an old file: " & (strDirectory & strFile) & vbcrlf
End If

'Create a new file ready for our commands
Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
WScript.Echo " - Created a new file: " & strDirectory & strFile & vbcrlf
Set objFolder = Nothing
Set objFile = Nothing

' *******************************************************************************************************
' Add the start/stop and our CMDLines to the NableJobs.bat
' *******************************************************************************************************
' Open the file
Set objFile = objFSO.OpenTextFile(strDirectory & strFile, 8, True)
' Add the lines to stop the agent and pause for 30 seconds
objFile.WriteLine AgentMaintStop
objFile.Writeline AgentStop
objFile.WriteLine PauseToReleaseLocks

' Add in our CMDLines to the file.
Dim i, j
For i=0 To UBound(CMDLine,1)
	objFile.Writeline CMDLine(i)
Next

' Add the lines to start the agent up again.
objFile.WriteLine AgentStart
objFile.WriteLine AgentMaintStart
WScript.Echo " - Added commands to the file: " & strDirectory & strFile & vbcrlf
objFile.Close

Set objFile = Nothing
Set objFSO = Nothing

' *******************************************************************************************************
' Schedule the Job to run in two Minutes from now (UK and USA date formats only)
' *******************************************************************************************************
delayMinutes = 2
startupScript = (strDirectory & strFile)
newTime = DateAdd("n", delayMinutes, Now())
timeString = Right("00" & DatePart("h", newTime), 2) & ":" & Right("00" & DatePart("n", newTime), 2) & ":00"
' Try and figure the right date format to use.. UK or USA
Location = GetLocale()
'UK Date Format
If Location = 2057 Then
	dateString = Right("00" & DatePart("d", newTime), 2) & "/" & Right("00" & DatePart("m", newTime), 2) & "/" & DatePart("YYYY", newTime)
Else
	'USA Date Format
	If Location = 1033 Then
		dateString = Right("00" & DatePart("m", newTime), 2) & "/" & Right("00" & DatePart("d", newTime), 2) & "/" & DatePart("YYYY", newTime)
	Else
		'Default to USA Date format and hope it works
		dateString = Right("00" & DatePart("m", newTime), 2) & "/" & Right("00" & DatePart("d", newTime), 2) & "/" & DatePart("YYYY", newTime)
	End If
End If 
' Use Schtasks to delete the old job (if it exists) and then create our new job to run 2 minutes from now.
Dim CMD
Set obJShell = WScript.CreateObject("WScript.Shell")
objShell.Run WinDir & "\system32\schtasks.exe /Delete /TN NableJobs /F" 
WScript.Sleep 5000 ' Avoids the delete/create running too quickly and overlapping.
oBJShell.Run WinDir & "\system32\schtasks.exe /create /tn NableJobs /tr " & startupScript & " /sc ONCE /sd " & dateString & " /st " & timeString & " /ru SYSTEM " 
WScript.Echo " - Scheduled the task for: " & timeString & " - " & dateString
