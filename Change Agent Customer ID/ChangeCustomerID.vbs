' Change Customer ID of a Windows agent
' By Chris Feldhaus, Systeem Medical Information Systems  (October 2013)
'
' Derived from: Change Appliance ID of a Windows agent
' By Tim Wiser, Orchid IT (August 2012)
'
' Modified by James Weakley, Diamond Technology Group (January 2014)
' Functionally behaves the same, but now launches itself into a separate 
' process which allows it to be run from the agent rather than the probe.
' Because it can all happen so fast, this script will typically time out
' on the N-Central agent it is ran against (according to N-Central), so 
' check for a new device in the target Customer to determine that it's 
' finished.
'
strComputer = "."


' Create the objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objShell = CreateObject("WScript.Shell")
Set output = WScript.StdOut

Dim objWMIService, objProcess
Dim strShell, objProgram, strExe 

' Get the parameter which should be the new customer ID number
Set objArgs = WScript.Arguments
Set objArguments = objArgs.Named
strNewCustomerID = objArguments("new")
If strNewCustomerID = "" Then
	output.writeline "No new customer ID specified!  Use /new:<New ID>"
	WScript.Quit 1
End If
output.writeline "Changing the customer ID of this device to " & strNewCustomerID


' Get the process handle of the whatever is launching this script. Initially this will
' be the N-Central agent process launcher, the second time will be the WMI service

Set launcher = GetObject("winmgmts:root\cimv2:Win32_Process.Handle='" _  
    & GetObject("winmgmts:root\cimv2:Win32_Process.Handle='" _  
    & GetObject("winmgmts:root\cimv2:Win32_Process.Handle='" _ 
    & CreateObject( "WScript.Shell").Exec("cmd.exe").ProcessId _  
    & "'").ParentProcessId & "'").ParentProcessId & "'")


output.writeline("Launched by: " & launcher.Name)
output.writeline("Script path: " & WScript.ScriptFullName)

' If the launcher wasn't WMI, launch this script again in a separate process using WMI and end
If StrComp(launcher.Name, "WmiPrvSE.exe") <> 0 Then
    ' Obtain the Win32_Process class of object.
    Set objProcess = objWMI.Get("Win32_Process")

    ' Spawn another separate cscript process to run this script
    Set objProgram = objProcess.Methods_("Create").InParameters.SpawnInstance_
    objProgram.CommandLine = "cscript.exe //I """ & WScript.ScriptFullName & """ /new:" &strNewCustomerID

    'Execute
    Set strShell = objWMI.ExecMethod("Win32_Process", "Create", objProgram) 
    ' Quit, as the new process will skip over this section and complete the job
    WScript.Quit 0

End If

strEnvironment = objShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
If strEnvironment = "AMD64" Then
	strProgramsPath = "C:\Program Files (x86)"
	output.writeline "This is a 64 bit device" & vbCrLf
Else
	strProgramsPath = "C:\Program Files"
	output.writeline "This is a 32 bit device" & vbCrLf
End If

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
		objNewConfigFile.WriteLine "  <ApplianceID>-1</ApplianceID>"
	ElseIf Left(strLine, 17)="  <CheckerLogSent" Then
		objNewConfigFile.WriteLine "  <CheckerLogSent>False</CheckerLogSent>"
	ElseIf Left(strLine, 13)="  <CustomerID" Then
		objNewConfigFile.WriteLine "  <CustomerID>" & strNewCustomerID & "</CustomerID>"
	ElseIf Left(strLine, 24)="  <CompletedRegistration" Then
		objNewConfigFile.WriteLine "  <CompletedRegistration>False</CompletedRegistration>"
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

' Terminate the agent for good measure.  Requesting it to stop often does not actually work and causes problems later on.
Set objProcess = objWMI.ExecQuery("Select Name from Win32_Process Where Name = 'agent.exe'")
For Each process in objProcess
	process.Terminate()
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


output.writeline "This device is now customer ID " & strNewCustomerID & ""
WScript.Quit 0

