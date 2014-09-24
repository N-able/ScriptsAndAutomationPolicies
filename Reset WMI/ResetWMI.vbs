'//////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////
'
' Clear out & Reset WMI & restart all related Services.
'
' Alex Woolsey - 11/04/2008
'
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////
On Error Resume Next

Set WshShell = CreateObject("Wscript.Shell")
If (right(Ucase(WScript.FullName),11)="WSCRIPT.EXE") Then 
 WshShell.Run WshShell.ExpandEnvironmentStrings("%COMSPEC%") & " /C cscript.exe """ & Wscript.ScriptFullName & """" 
 Wscript.Quit
End If 

With WshShell
oReport "Running WMI Kill & Re-Register Commands.... Please Wait..."
.Run("winmgmt /clearadap"), 0, True
.Run("winmgmt /kill"), 0, True
.Run("winmgmt /unregserver"), 0, True
.Run("winmgmt /regserver"), 0, True
.Run("winmgmt /resyncperf"), 0, True
.Run("wmiadap /c"), 0, True
.Run("wmiadap /f"), 0, True
oReport "WMI Commands Completed...." & VbCrLf
End With

oReport "Stopping Windows Management Dependent Services...."
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServiceList = objWMIService.ExecQuery("Associators of " _
    & "{Win32_Service.Name='winmgmt'} Where " _
        & "AssocClass=Win32_DependentService " & "Role=Antecedent" )

For Each objService in colServiceList
oReport "Stopping " & objService.DisplayName
    objService.StopService()
Next

oReport "Sleeping for 20 Seconds to Allow System Catchup..."
Wscript.Sleep 20000

oReport "Stopping Windows Management Services...."
Set colServiceList = objWMIService.ExecQuery _
        ("Select * from Win32_Service where Name='winmgmt'")
For Each objService in colServiceList
oReport "Stopping " & objService.DisplayName
    errReturn = objService.StopService()
Next

oReport "Sleeping for 20 Seconds to Allow System Catchup..."
Wscript.Sleep 20000

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colServiceList = objWMIService.ExecQuery _
        ("Select * from Win32_Service where Name='winmgmt'")
For Each objService in colServiceList
oReport "Restarting " & objService.DisplayName
    errReturn = objService.StartService()
Next

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServiceList = objWMIService.ExecQuery("Associators of " _
    & "{Win32_Service.Name='winmgmt'} Where " _
        & "AssocClass=Win32_DependentService " & "Role=Antecedent" )

For Each objService in colServiceList
oReport "Restarting " & objService.DisplayName
    objService.StartService()
Next

strMessage = "Press the ENTER key to continue. "
Wscript.StdOut.Write strMessage

Do While Not WScript.StdIn.AtEndOfLine
   Input = WScript.StdIn.Read(1)
Loop

Function oReport(strMessage)
 Wscript.Echo VbTab & strMessage
End Function
