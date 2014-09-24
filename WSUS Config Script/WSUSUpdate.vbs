'********************************************************************************
'*********** Script: WSUSUpdate.vbs
'*********** Description: Configures WSUS settings via the registry. This script  
'***********              is intended for devices not in a domain.
'*********** Author: Chris Reid
'*********** Date: March 7th, 2008
'*********** Usage: At the DOS prompt: cscript WSUSUpdate.vbs TargetGroup WSUSAddress
'***********        (i.e. cscript WSUSUpdate.vbs AcmeCo http://1.2.3.4)
'***********
'***********        In N-central: WSUSUpdate.vbs TargetGroup WSUSAddress
'********************************************************************************

'NOTE: For information on the variables being set by this script, please refer to the following Microsoft KB article:
' http://technet2.microsoft.com/windowsserver/en/library/75ee9da8-0ffd-400c-b722-aeafdb68ceb31033.mspx?mfr=true

'NOTE: Please consult with a Microsoft Certified Engineer before running this script. This script is provided as-is, with no warranty
'      or guarantee.


' Check to make sure that two variables were passed to the script, and give the variables names.
If WScript.Arguments.Count = 2 Then
  TargetGroup = WScript.Arguments.Item(0)
  WSUSAddress = WScript.Arguments.Item(1)
Else
 wscript.Quit(1) 	
End If


'Declare Windows registry variables
RegPath1 = "HKLM\Software\Policies\Microsoft\Windows\WindowsUpdate"
RegPath2 = "HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"


'Modify the settings of the "HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\" registry keys

        Set Shell = CreateObject("Wscript.Shell")
          Shell. RegWrite "" & RegPath1 & "\ElevateNonAdmins", 0, "REG_DWORD"
          Shell. RegWrite "" & RegPath1 & "\TargetGroup", "" & TargetGroup & "", "REG_SZ"
          Shell. RegWrite "" & RegPath1 & "\TargetGroupEnabled",1, "REG_DWORD"
          Shell. RegWrite "" & RegPath1 & "\WUServer", "" & WSUSAddress & "", "REG_SZ"
          Shell. RegWrite "" & RegPath1 & "\WUStatusServer","" & WSUSAddress & "", "REG_SZ"
        Set Shell = Nothing
        
        
        
        
'Modify the settings of the "HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\AU\" registry keys

        Set Shell = CreateObject("Wscript.Shell")
          Shell. RegWrite "" & RegPath2 & "\AUOptions", 4, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\AutoInstallMinorUpdates", 1, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\DetectionFrequency", 1, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\DetectionFrequencyEnabled", 1, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\NoAutoUpdate", 0, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\RebootRelaunchTimeout", 60, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\RebootRelaunchTimeoutEnabled", 1, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\RebootWarningTimeout", 30, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\RebootWarningTimeoutEnabled", 1, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\RescheduleWaitTime", 0, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\RescheduleWaitTimeEnabled", 0, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\ScheduledInstallDay", 0, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\ScheduledInstallTime", 15, "REG_DWORD"
          Shell. RegWrite "" & RegPath2 & "\UseWUServer", 1, "REG_DWORD"
        Set Shell = Nothing
        
        
        
'Stop the "Automatic Updates" service
        Set Shell = CreateObject("Wscript.Shell")
          Shell.Run "NET STOP wuauserv",7,TRUE
        Set Shell = Nothing
        
'Delete the log file for the "Automatic Updates" service
        Set FSO = CreateObject("Scripting.FileSystemObject")
          FSO.DeleteFile "\\127.0.0.1\ADMIN$\windowsupdate.log"
        Set FSO = Nothing
        
'Start the "Automatic Updates" service
        Set Shell = CreateObject("Wscript.Shell")
          Shell.Run "NET START wuauserv",7,TRUE
        Set Shell = Nothing
          
'Force the device to check in with the WSUS server
        Dim FSO
        Set Shell = CreateObject("Wscript.Shell")
          Shell.Run "wuauclt.exe /resetauthorization /detectnow",7,TRUE
        Set Shell = Nothing
