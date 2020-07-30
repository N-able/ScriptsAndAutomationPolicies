'******************************************************************************
' Script: UninstallSymantecES.vbs
' Version: 1.1
' Description: This script uninstalls Symantec Endpoint Security 11.0.5
' Author: Chris Reid and Jonathan Filson
' Date: Dec. 11th, 2009
'******************************************************************************

' Declare the variables that we're going to use
option explicit
dim Shell, ThirtyTwoBitUninstall, SixtyFourBitUninstall


'Specify the uninstall command for the 32-bit and 64-bit versions of Symantec Endpoint Security
ThirtyTwoBitUninstall = "MsiExec.exe /norestart /qn /x{2EFCC193-D915-4CCB-9201-31773A27BC06}"
SixtyFourBitUninstall = "MsiExec.exe /norestart /qn /x{530992D4-DDBA-4F68-8B0D-FF50AC57531B}"




'Create a shell, and run the 32-bit and 64-bit uninstall commands
Set Shell = CreateObject("Wscript.Shell")
Shell.Run "%COMSPEC% /c " & ThirtyTwoBitUninstall,7,TRUE
Shell.Run "%COMSPEC% /c " & SixtyFourBitUninstall,7,TRUE
Set Shell = Nothing




'****************************
' Version History
'****************************

'Version 1.1
'   Added support for uninstalling the 32-bit version of Symantec Endpoint Security


'Version 1.0
'  Initial Release