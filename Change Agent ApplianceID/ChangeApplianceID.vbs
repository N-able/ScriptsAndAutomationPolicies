
' Change appliance ID of a Windows agent
' by Tim Wiser, Orchid IT (August 2012)

' This script MUST be launched using the Probe, not the Agent


strComputer = "."

' Create the objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objShell = CreateObject("WScript.Shell")
Set output = WScript.StdOut

' Get the architecture of the device and adjust the location of the agent folder accordingly
strEnvironment = objShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
If strEnvironment = "AMD64" Then
	strProgramsPath = "C:\Program Files (x86)"
	output.writeline "This is a 64 bit device" & vbCrLf
Else
	strProgramsPath = "C:\Program Files"
	output.writeline "This is a 32 bit device" & vbCrLf
End If



' Get the parameter which should be the new appliance ID number
Set objArgs = WScript.Arguments
Set objArguments = objArgs.Named
strNewApplianceID = objArguments("new")
If strNewApplianceID = "" Then
	output.writeline "No new appliance ID specified!  Use /new:<New ID>"
	WScript.Quit 1
End If
output.writeline "Changing the appliance ID of this device to " & strNewApplianceID



' Check to see if the config file is actually where we expect it to be
If objFSO.FileExists(strProgramsPath & "\N-able Technologies\Windows Agent\config\ApplianceConfig.xml")=False Then
	output.writeline "XML configuration file could not be found!" & vbCrLf
	WScript.Quit 1
Else
	output.writeline "XML configuration file was found" & vbCrLf
End If



Set objNewConfigFile = objFSO.CreateTextFile("C:\Windows\Temp\ApplianceConfig.xml")
Set objConfigFile = objFSO.OpenTextFile(strProgramsPath & "\N-able Technologies\Windows Agent\config\ApplianceConfig.xml")
strLine = ""

Do Until objConfigFile.AtEndOfStream
	strLine = objConfigFile.ReadLine
	If Left(strLine, 14)="  <ApplianceID" Then
		objNewConfigFile.WriteLine "  <ApplianceID>" & strNewApplianceID & "</ApplianceID>"
	Else
		objNewConfigFile.Writeline strLine
	End If
Loop


' Close the files
objConfigFile.Close
objNewConfigFile.Close

output.writeline "Stopping the N-able agent..."
Set objServices = objWMI.ExecQuery("SELECT Name FROM Win32_Service WHERE Name LIKE 'Windows Agent%'")
For Each service in objServices
	output.writeline "Stopping service: " & service.Name & vbCrLf
	service.StopService
Next

' Wait for the agent service to finish stopping
AgentCount=0
AgentStopped = 0
Do Until AgentStopped = 1
	Set objService = objWMI.ExecQuery("SELECT Name,State FROM Win32_Service WHERE Name='Windows Agent Service'")
	
	For Each service in objService
		' Is the agent stopped, if so then bug out of the loop
		If service.State = "Stopped" Then
			output.writeline "The agent has finished stopping" & vbCrLf
			AgentStopped = 1
		End If

		' This bit prevents the script going into an endless loop if the agent cannot be detected as being stopped
		AgentCount = AgentCount + 1
		If AgentCount > 10000 Then
			output.writeline "Could not detect the agent stopping!" & vbCrLf
			WScript.Quit 1
		End If
	Next
Loop


' Copy the config file across into the agent folder
objFSO.CopyFile "C:\Windows\Temp\ApplianceConfig.xml", strProgramsPath & "\N-able Technologies\Windows Agent\config\ApplianceConfig.xml"


' Start the agent up
output.writeline "Starting the N-able agent..."
Set objServices = objWMI.ExecQuery("SELECT Name, State FROM Win32_Service WHERE Name LIKE 'Windows Agent%'")
For Each service in objServices
	output.writeline "Starting service: " & service.Name & vbCrLf
	service.StartService
Next


output.writeline "This device is now appliance ID " & strNewApplianceID & ""
WScript.Quit 0
