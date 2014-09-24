'******************************************************************************************
'************************************ DiskCleanUp.vbs v1.0  *******************************
'****************************      Written By: Chris Reid      ****************************
'****************************       Date: July 21st, 2009       ***************************
'****************************   Copyright N-able Technologies  ****************************
'***  WARNING: This script should only be run by a Microsoft certified Administrator   ****
'***   N-able Technologies presents this script AS-IS - use at your own discretion      ***
'******************************************************************************************

'Description
'The purpose of this script is to clear out temporary files by running Microsoft's Windows Disk Cleanup (cleanmgr.exe) application.
'By default, this script will run against all drives (local and network) attached to the device
'Information on how to run Disk Cleanup from the command line was handily provided by the following Microsoft KB article: http://support.microsoft.com/kb/315246


'Let's declare some variables
option explicit
Dim Shell


'Let's decide what options in Windows Disk Cleanup we want to run. We're going to create an overall job called 1122 (represented by StateFlags1122 in the registry).
'A value of 2 turns that option on, and a value of 0 turns it off.
'By default, all options have been enabled
Set Shell = CreateObject("Wscript.Shell")
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Active Setup Temp Folders\StateFlags1122", 2, "REG_DWORD" 
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Content Indexer Cleaner\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Downloaded Program Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Hibernation File\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Internet Cache Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Memory Dump Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Microsoft Office Temp Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Offline Pages Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Old ChkDsk Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Previous Installations\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Recycle Bin\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Setup Log Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\System error memory dump files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\System error minidump files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Temporary Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Temporary Setup Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Temporary Sync Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Thumbnail Cache\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Upgrade Discarded Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Windows Error Reporting Archive Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Windows Error Reporting Queue Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Windows Error Reporting System Archive Files\StateFlags1122", 2, "REG_DWORD"
Shell. RegWrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Windows Error Reporting System Queue Files\StateFlags1122", 2, "REG_DWORD"
Set Shell = Nothing


'Let's run the 1122 job by using the /sagerun:1122 flag
'If you want to do it on a specific drive, run cleanmgr.exe /d C:\ /sagerun:1122
Set Shell = CreateObject("Wscript.Shell")
Shell.Run "cleanmgr.exe /sagerun:1122",7,TRUE
Set Shell = Nothing



'Having a tidy registry is a good thing. Let's delete the registry entries created by this script
Set Shell = CreateObject("Wscript.Shell")
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Active Setup Temp Folders\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Content Indexer Cleaner\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Downloaded Program Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Hibernation File\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Internet Cache Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Memory Dump Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Microsoft Office Temp Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Offline Pages Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Old ChkDsk Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Previous Installations\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Recycle Bin\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Setup Log Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\System error memory dump files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\System error minidump files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Temporary Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Temporary Setup Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Temporary Sync Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Thumbnail Cache\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Upgrade Discarded Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Windows Error Reporting Archive Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Windows Error Reporting Queue Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Windows Error Reporting System Archive Files\StateFlags1122"
Shell. RegDelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Windows Error Reporting System Queue Files\StateFlags1122"
Set Shell = Nothing

