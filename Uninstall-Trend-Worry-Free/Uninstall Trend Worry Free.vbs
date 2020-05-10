'******************************************************************
' Script: Uninstall Trend.vbs
' Description: This script will uninstall Trend Worry-Free A/V
' Author: Chris Reid
' Version: 1.1
' Date: Jan. 5th, 2011
'******************************************************************



' Version History

' 1.1 - Added code to delete any remnants left behind in WMI (Jan 5th, 2011)

' 1.0 - Initial Release (Dec 10th, 2009)

' VB Script Document
option explicit
dim Shell, UninstallPath, strComputer, oReg, strKeyPath, strValueName,strValue, objSWbemServices, colSWbemObjectSet, objSWbemObject
const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."

'Set the registry to allow an uninstall
Set Shell = CreateObject("Wscript.Shell")
Shell. RegWrite"HKLM\Software\TrendMicro\PC-cillinNTCorp\CurrentVersion\Misc.\Allow Uninstall", 1, "REG_DWORD"
Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so "& chr(34) & "N-ABLE (Trend Uninstall)"& chr(34) & " /d "& chr(34) & "The registry value that allows Trend Worry-Free Edition 6.x to be uninstalled has been set."& chr(34),7,TRUE
Set Shell = Nothing


'Grab the uninstall path from the registry
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OfficeScanNT"
strValueName = "UninstallString"
oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue


'Run the uninstaller
Set Shell = CreateObject("Wscript.Shell")
Shell.Run "%COMSPEC% /c " & strValue,7,TRUE 
Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so "& chr(34) & "N-ABLE (Trend Uninstall)"& chr(34) & " /d "& chr(34) & "The Trend Worry-Free Edition 6.x uninstaller has been started."& chr(34),7,TRUE
Set Shell = Nothing


'Delete any entries in root\SecurityCenter\AntiVirusProduct - failing to do this will prevent Endpoint Security from uninstalling.
Set objSWbemServices = GetObject("winmgmts:\\" & strComputer & "\root\SecurityCenter")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("AntiVirusProduct")
For Each objSWbemObject In colSWbemObjectSet
    objSWbemObject.Delete_
Next


