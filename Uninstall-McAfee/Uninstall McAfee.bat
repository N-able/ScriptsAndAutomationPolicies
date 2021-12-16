@ECHO OFF
REM McAfee Removal Script
REM Last Update: 11/01/2010
REM

ECHO Removing AntiSpyware
"C:\Program Files\McAfee\VirusScan Enterprise\scan32.exe" /UninstallMAS
"C:\Program Files (x86)\McAfee\VirusScan Enterprise\scan32.exe" /UninstallMAS

REM Kill McTray & Trusted Validation
ECHO Killing processes
taskkill.exe /f /t /im mctray.exe
taskkill.exe /f /t /im mfevtps.exe

ECHO Removing VirusScan 8.0
msiexec.exe /x {5DF3D1BB-894E-4DCD-8275-159AC9829B43} REMOVE=ALL REBOOT=R /q

ECHO Removing VirusScan 8.5
msiexec.exe /x {35C03C04-3F1F-42C2-A989-A757EE691F65} REMOVE=ALL REBOOT=R /q

ECHO Removing VirusScan 8.7
msiexec.exe /x {147BCE03-C0F1-4C9F-8157-6A89B6D2D973} REMOVE=ALL REBOOT=R /q

ECHO Remove McAfee Agent
"C:\Program Files\McAfee\Common Framework\frminst.exe" /forceuninstall /silent
"C:\Program Files (x86)\McAfee\Common Framework\frminst.exe" /forceuninstall /silent

REM Remove McAfee Registry Keys
ECHO Removing Registry Keys
REG DELETE HKLM\SYSTEM\CurrentControlSet\services\McShield /f
REG DELETE HKLM\SYSTEM\CurrentControlSet\services\McTaskManager /f
REG DELETE HKLM\SYSTEM\CurrentControlSet\Services\mfeapfk /f
REG DELETE HKLM\SYSTEM\CurrentControlSet\Services\mfeavfk /f
REG DELETE HKLM\SYSTEM\CurrentControlSet\Services\mfebopk /f
REG DELETE HKLM\SYSTEM\CurrentControlSet\Services\mfehidk /f
REG DELETE HKLM\SYSTEM\CurrentControlSet\Services\mferkdet /f
REG DELETE HKLM\SYSTEM\CurrentControlSet\Services\mfetdik /f
REG DELETE HKLM\SYSTEM\CurrentControlSet\Services\mfevtp /f
REG DELETE HKLM\SOFTWARE\McAfee /f