REM 
REM Stopping Sophos Services
net stop "Sophos AutoUpdate Service"
net stop "Sophos Agent"
net stop "SAVService"
net stop "SAVAdminService"
net stop "Sophos Message Router"
net stop "Sophos Web Control Service"
net stop "swi_service"
net stop "swi_update"

REM 
REM Removing Sophos AutoUpdater
MsiExec.exe /X{D924231F-D02D-4E0B-B511-CC4A0E3ED547} /qn REBOOT=SUPPRESS /PASSIVE

REM
REM Removing Sophos Remote Management System
MsiExec.exe /X{FED1005D-CBC8-45D5-A288-FFC7BB304121} /qn REBOOT=SUPPRESS /PASSIVE

REM
REM Removing Sophos Anti-Virus
MsiExec.exe /X{D929B3B5-56C6-46CC-B3A3-A1A784CBB8E4} /qn REBOOT=SUPPRESS /PASSIVE

REM
REM Removing Sophos Patch
MsiExec.exe /X{29006785-9EF7-4E84-ABE8-6244D12E7909} /qn REBOOT=SUPPRESS /PASSIVE