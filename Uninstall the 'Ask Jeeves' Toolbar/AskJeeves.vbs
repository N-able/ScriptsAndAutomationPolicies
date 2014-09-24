'##########################################
'# Copyright 2009 by N-able Technologies  #
'# All Rights Reserved                    #
'# May not be reproduced or redistributed #
'# Without written consent from N-able    #
'# Technologies			                      #
'# www.n-able.com			                    #
'##########################################


' Script Name: AskJeeves.vbs
' Description: This VBS script will uninstall the IE version of the Ask Jeeves Toolbar.


'Let's declare some variables
option explicit
dim strComputer, strKeyPath, strValueName, objReg, dwValue, Shell, cmdline
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."

'Grab the uninstall string for the IE version from the registry 
'(this script will also use the uninstall string to check whether or not the Ask Jeeves Toolbar is installed)
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
        & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall\Ask Toolbar for Internet Explorer_is1"
strValueName = "QuietUninstallString"
objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue

'Check to see if the Ask Jeeves Toolbar is actually installed
If IsNull(dwValue) Then
    
Else
    'Uninstall the IE version of the Ask Jeeves Toolbar
    Set Shell = CreateObject("Wscript.Shell")
    cmdline = "" & dwValue
    Shell.Run cmdline,7,TRUE
    Set Shell = Nothing
End If



