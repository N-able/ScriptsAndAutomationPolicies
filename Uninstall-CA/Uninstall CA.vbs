on error resume next
Set WshShell = CreateObject("wscript.Shell")
                'result=WshShell.Run(<StrCommand>, vbHidden)
Const HKEY_LOCAL_MACHINE = &H80000002



'##########################################################################################
'#################################### Set the defaults ####################################

'DefaultRemoveApp = "CA eTrust Antivirus"       ' Replace DefaultRemoveApp with your app name, copy the name from the key value: 
                                                                                                ' SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*appname*\DispalyName

Dim DefaultRemoveApps(1)                                                       'Set Equal to the Appname or as a search criteria
  DefaultRemoveApps(0) = "CA eTrust"                                  '"eTrust Catchall
  DefaultRemoveApps(1) = "CA iTechnology iGateway"  '"CA iTechnology iGateway" is Ver.8

  'DefaultRemoveApps(0) = "CA eTrust Antivirus"                              '"CA eTrust Antivirus" is Ver.7
  'DefaultRemoveApps(1) = "CA eTrustITM Agent"                           '"CA eTrustITM Agent" is Ver.8


Local_param = " /Quiet /norestart"         ' MSIExec parameter added when running locally                             Recommend: " /Quiet /norestart"
                                                                                ' run "msiexec /?" in CMD for Paramaters 
Last_param = " /Quiet /norestart"           ' MSIExec parameter added to the last uninstall                 Recommend: " /Quiet /promptrestart"               

'##########################################################################################
'##########################################################################################
strComputer = "."                                                            'for the local machine use "."
StrLastCMD=""

err.clear
For j = 0 to UBound(DefaultRemoveApps)
                'wscript.echo "Start " & DefaultRemoveApps(j)
                StrLastCMD=FindAndUninstall(DefaultRemoveApps(j),StrLastCMD)
Next
if StrLastCMD<>"" then
                StrLastCMD = Replace(StrLastCMD,Local_param,Last_param)
                RunCMD(StrLastCMD)
end if

wscript.echo "Done With Script!"
wscript.quit



'##########################################################################################
'#################################### SUBS ################################################

'################### Chech if error occured ###################
Sub ErrorCheck(StrMessage,StrErr)
                If err.number <> 0 Then 
                                Wscript.echo(StrMessage & " " & err.number)
                                wscript.quit
                end if
                err.clear
end sub

Sub RunCMD(StrCommand)
                err.clear
                wscript.echo "Running " & StrCommand
                result=WshShell.Run(StrCommand,3,TRUE)
                ErrorCheck "Error Running " & StrCommand,err.number
End Sub

'################### Find if Software is installed and write the uninstall command to batch file ###################
Function FindAndUninstall(appDisplayName,rmValueGiven)
                on error resume next
                rmValueFixed=""
                rmValueLast=rmValueGiven
                err.clear
                Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
                ErrorCheck "Error connecting to Registry",err.number

                err.clear
                strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

                '****************** Go through Registry of installed apps  ****************
                For Each subkey In arrSubKeys
                                appKeyPath = strKeyPath & "\" & subkey
                                strValueName = "DisplayName"
                                oReg.GetStringValue HKEY_LOCAL_MACHINE,appKeyPath,strValueName,strValue

                                '****************** find the application to uninstall  ****************
                                                If Instr(lcase(Trim(strValue)),lcase(appDisplayName)) Then

                                                                                
                                                                                rmValueName = "UninstallString"

                                                                                '****************** extract the uninstall parameter  ****************
                                                                                oReg.GetStringValue HKEY_LOCAL_MACHINE,appKeyPath,rmValueName,rmValue

                                                                                if instr(LCase(rmValue),"msiexec.exe") then
                                                                                                '****************** replace any install switches with uninstall ones  ****************
                                                                                                rmValueFixed = Replace(UCase(rmValue),"/I","/X")
                                                                                                rmValueFixed = Replace(UCase(rmValueFixed),"/PACKAGE","/X")
                                                                                                rmValueFixed = rmValueFixed & Local_param
                                                                                else
                                                                                                rmValueFixed = rmValue
                                                                                end if

                                                                if rmValueLast<>rmValueFixed then
                                                                                if rmValueLast<>"" then
                                                                                                RunCMD(rmValueLast)
                                                                                end if

                                                                                wscript.echo "Uninstalling  " & Trim(strValue) &"  Because it matches  " & appDisplayName '& "===" & rmValueFixed
                                                                                rmValueLast=rmValueFixed

                                                                End If
                                                End If
                Next

                FindAndUninstall=rmValueLast
                set arrSubKeys = nothing
                set rmValue = nothing
                oReg.close
End Function
