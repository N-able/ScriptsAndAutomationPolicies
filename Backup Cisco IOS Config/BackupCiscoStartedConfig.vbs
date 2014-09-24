'****************************************************************************************************************************************************************
' Script: BackupCiscoStartedConfig.vbs
' Description: This script backs up the Started config on a Cisco device. You must put the IP addresses of the Cisco devices in a text file, and specify
'              the name of the text file when running this script. The IP addresses must be comma-separated. 
' Author: Chris Reid from N-able Technologies, Phil Ozak from Datacore Consulting and Bill Heinrich from WTC Networks
' Version: 1.1
' Running Commands: cscript BackupCiscoStartedConfig.vbs IPADDRESSES.txt
' ***************************************************************************************************************************************************************


' *******************
' Version History
' *******************

' v1.1 - Added the ability to specify an FTP username and password, and altered the commands that were being run (Jan. 4th, 2011)

' v1.0 - Initial release (April 8th, 2010)

' Let's declare some variables
Option Explicit
dim TEXTFILE, objFSO, objTextFile, strNextLine, arrServiceList, i, WshShell, FTPIPADDRESS, FTPUSERNAME, FTPPASSWORD, CISCOIPADDRESS, CUSTOMERNAME, CISCOUSERNAME, CISCOPASSWORD


' The person running this script needs to fill out these values
FTPIPADDRESS = "x.x.x.x" 'Note that this *MUST* be an IP address - FQDNs are not permitted
FTPUSERNAME = "username"
FTPPASSWORD = "password"
CISCOUSERNAME = "username"
CISCOPASSWORD = "password"
CUSTOMERNAME = "customername" ' Make sure that this does not contain any spaces!


'Lets make sure that the user passed us the text file containing the IP addresses as a command-line parameter
If WScript.Arguments.Count = 1 Then
  TEXTFILE = WScript.Arguments.Item(0) 
  ReadIP
Else
  'Let's log an event to the Windows Event Log that explains what went wrong
  Set Shell = CreateObject("Wscript.Shell")   
  Shell.Run "%COMSPEC% /c eventcreate /t ERROR /id 999 /l APPLICATION /so CiscoConfigBackup /d "& chr(34) & "You did not specify the text file containing the IP addresses as a command line parameter."& chr(34),7,TRUE
  Set Shell = Nothing
  wscript.quit(1)                
End If


' This sub-routine will read in the IP addresses from the specified text file
' ************************
Sub ReadIP
' ************************
      Const ForReading = 1
      Set objFSO = CreateObject("Scripting.FileSystemObject")
      Set objTextFile = objFSO.OpenTextFile _
          (TEXTFILE, ForReading)
      Do Until objTextFile.AtEndOfStream
          strNextLine = objTextFile.Readline
          arrServiceList = Split(strNextLine , ",")
          For i = 0 to Ubound(arrServiceList)
              CISCOIPADDRESS = arrServiceList(i)
              SConfig 'this calls the SConfig sub-routine, which will configure the Cisco device to do a backup of it's startup config
              RConfig 'this calls the RConfig sub-routine, which will configure the Cisco device to do a backup of it's Running config 
          Next
      Loop
End Sub


' Via telnet, this sub-routine configures the Cisco device to send its Startup Config
' ************************
Sub SConfig
' ************************
                Set WshShell = WScript.CreateObject("WScript.Shell") 
                Run "cmd.exe" ' This line calls the 'Run' sub-routine to launch cmd.exe
                SendKeys "telnet " & CISCOIPADDRESS & "{ENTER}"
                SendKeys CISCOUSERNAME
                SendKeys "{ENTER}"
                SendKeys CISCOPASSWORD
                SendKeys "{ENTER}"
                SendKeys "en"
                SendKeys "{ENTER}"
                SendKeys CISCOPASSWORD 
                SendKeys "{ENTER}" 
                SendKeys "copy startup-config ftp://" & FTPUSERNAME & ":" & FTPPASSWORD & "@" & FTPIPADDRESS & "/" & CUSTOMERNAME &"-sconfig-"& CISCOIPADDRESS &"-"& Replace(Replace(Replace(Now," ","-"),"/","-"),":","-") &".txt{ENTER}"
                SendKeys "{ENTER}"
                                                                SendKeys "{ENTER}"
                                                                SendKeys "{ENTER}"
                                                                SendKeys "{ENTER}"
                                                                Wscript.Sleep 5000
                                                                SendKeys "exit{ENTER}" 'close telnet session' 
                SendKeys "{ENTER}"
                WScript.Sleep 1000
                SendKeys "{ENTER}"
                SendKeys "exit{ENTER}" 'close cmd.exe
      Set WshShell = Nothing           
End Sub







' Via telnet, this sub-routine configures the Cisco device to send its Running Config
' ************************
Sub RConfig
' ************************
                Set WshShell = WScript.CreateObject("WScript.Shell") 
                Run "cmd.exe" ' This line calls the 'Run' sub-routine to launch cmd.exe
                SendKeys "telnet " & CISCOIPADDRESS & "{ENTER}"
                SendKeys CISCOUSERNAME
                SendKeys "{ENTER}"
                SendKeys CISCOPASSWORD
                SendKeys "{ENTER}"
                SendKeys "en"
                SendKeys "{ENTER}"
                SendKeys CISCOPASSWORD 
                SendKeys "{ENTER}" 
                SendKeys "copy running-config ftp://" & FTPUSERNAME & ":" & FTPPASSWORD & "@" & FTPIPADDRESS & "/" & CUSTOMERNAME &"-rconfig-"& CISCOIPADDRESS &"-"& Replace(Replace(Replace(Now," ","-"),"/","-"),":","-") &".txt{ENTER}"
                SendKeys "{ENTER}"
                                                                SendKeys "{ENTER}"
                                                                SendKeys "{ENTER}"
                                                                SendKeys "{ENTER}"
                                                                SendKeys "{ENTER}"
                                                                Wscript.Sleep 5000
                                                                SendKeys "exit{ENTER}" 'close telnet session' 
                SendKeys "{ENTER}"
                WScript.Sleep 1000
                SendKeys "{ENTER}"
                ' SendKeys "exit{ENTER}" 'close cmd.exe
      Set WshShell = Nothing           
End Sub






' This sub-routine sends the specified command (denoted by the variable 's' in this case)
' ************************
Sub SendKeys(s)
' ************************
        WshShell.SendKeys s
        WScript.Sleep 1000
End Sub


' This sub-routine launchs the specified command, and then makes it the active window
' ************************
Sub Run(command)
' ************************
        WshShell.Run command
        WScript.Sleep 100 
        WshShell.AppActivate command 
        WScript.Sleep 300 
End Sub
