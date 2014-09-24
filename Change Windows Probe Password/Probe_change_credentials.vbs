'###############################
'# Copyright 2009 by Cougar Ridge Computer Systems
'# All Rights Reserved 
'###############################
'#
'# Required Parameters:
'#
'# /u - domain\username
'# /p - Password
'# 
'# Sample
'# Probe_change_credentials.vbs /u:domain\administrator /p:Password

'# Create Objects
set WshShell        = CreateObject("WScript.Shell")
Set objShell        = CreateObject("Shell.Application" ) 
Set objFSO          = CreateObject("Scripting.FileSystemObject")

'# Set Variables
strUser         = WScript.Arguments.Named.Item("u")
strPassword     = WScript.Arguments.Named.Item("p")
strComputer     = "."
const HKEY_LOCAL_MACHINE = &H80000002

'# The following code elevates the script if we're using an operating system requiring elevation.
If WScript.Arguments.Named.Item("uac") = "" Then
    '#Check for Vista or XP.
    dim Vistaver, dwvalue, UACStatus, UACValue
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * From Win32_OperatingSystem")
    For Each objItem in colItems
    If Instr(objItem.Caption, "Vista") Then
        OperatingSystem = "UAC"
    ElseIf Instr(objItem.Caption, "Windows 7") Then
        OperatingSystem = "UAC"
    End If    
    Next
    if (OperatingSystem = "UAC") Then
        answer = MsgBox("This script is being run to change the Windows Software Probe password on your system.  If you have questions or are concerned that this may be a virus please contact us at 000-000-0000.",0,"IT Support")   
        objShell.ShellExecute "wscript.exe", Chr(34) & _
        WScript.ScriptFullName & Chr(34) & " /uac:1 /u:" & strUser & " /p:" & strPassword, "", "runas", 1
    Else
        change_credentials    
    End If
Else
    change_credentials 
End if

sub change_credentials()
'# Stop Services
cmd = "cmd.exe /c net stop ""Windows Software Probe Maintenance Service"""
Return = WshShell.Run(cmd,0,true) 
cmd = "cmd.exe /c net stop ""Windows Software Probe Service"""
Return = WshShell.Run(cmd,0,true) 
cmd = "cmd.exe /c net stop ""Windows Software Probe Syslog Service"""
Return = WshShell.Run(cmd,0,true) 

'Set working directory

if objFSO.FolderExists("c:\Program Files (x86)\N-Able Technologies\Windows Software Probe\bin") then
	strNableDirectory = "c:\Program Files (x86)\N-Able Technologies\Windows Software Probe\bin"
else
	strNableDirectory = "c:\Program Files\N-Able Technologies\Windows Software Probe\bin"
end if

'Create startup.ini File
set Nablefilestream = objFSO.CreateTextFile(strNableDirectory & "\startup.ini", True, 0) 
With Nablefilestream
    .WriteLine "username=" & strUser
    .WriteLine "password=" & strPassword
    .WriteLine ""
    .Close
End With

' # Adjust service login information
Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set objService = objWMI.Get("Win32_Service.Name='Windows Software Probe Maintenance Service'")
ChangeServ = objservice.Change(Null, Null, Null, Null, Null, FALSE, struser, strPassword)
set objService = objWMI.Get("Win32_Service.Name='Windows Software Probe Service'")
ChangeServ = objservice.Change(Null, Null, Null, Null, Null, FALSE, struser, strPassword)  
set objService = objWMI.Get("Win32_Service.Name='Windows Software Probe Syslog Service'")
ChangeServ = objservice.Change(Null, Null, Null, Null, Null, FALSE, struser, strPassword)

'# Start Services
cmd = "cmd.exe /c net start ""Windows Software Probe Maintenance Service"""
Return = WshShell.Run(cmd,0,true) 
cmd = "cmd.exe /c net start ""Windows Software Probe Service"""
Return = WshShell.Run(cmd,0,true) 
cmd = "cmd.exe /c net start ""Windows Software Probe Syslog Service"""
Return = WshShell.Run(cmd,0,true) 

End Sub