' ***********************************************************************************************
' Script: UninstallMSSecurityEssentials.vbs
' Description: This script uninstalls the 32-bit version of Microsoft Security Essentials
' Author: Jonathan Filson
' Date: Dec. 15th, 2009
' ***********************************************************************************************


' VB Script Document
option explicit
dim Shell, Application


'Configure what application should be run
Application = "MsiExec.exe /norestart /qn /X{A0A77CDC-2419-4D5C-AD2C-E09E5926B806}"



'Create a shell, and run the command
Set Shell = CreateObject("Wscript.Shell")
Shell.Run "%COMSPEC% /c " & Application,7,TRUE
Set Shell = Nothing




