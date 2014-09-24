net stop wuauserv
reg delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v AccountDomainSid /f 
reg delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v PingID /f 
reg delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientId /f 
del %SystemRoot%\SoftwareDistribution\*.* /S /Q 
reg delete HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /f
net start wuauserv