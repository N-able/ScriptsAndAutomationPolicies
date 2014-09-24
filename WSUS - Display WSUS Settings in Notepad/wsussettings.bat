@echo off
::
Echo This script will query Automatic Update Client / Windows Update Agent settings. 
REM as configured via Group Policy to detect/download/install updates from WSUS / SUS.

REM Result will be written to %temp%\WUASettings.txt and then launched in Notepad.
REM Disable Word Wrap and Save the batch file "WUASettings.cmd".

Echo More information on PatchAholic Blog http://msmvps.com/athif
Echo For Windows 2000 machines - 
Echo You can download REG.EXE from http://www.dynawell.com/reskit/microsoft/win2000/reg.zip 

REM Author:- Mohammed Athif Khaleel :- Date April 25, 2006
Pause
::
@echo on
reg query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /s > %temp%\WUASettings.txt
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" >> 

%temp%\WUASettings.txt

notepad %temp%\WUASettings.txt

