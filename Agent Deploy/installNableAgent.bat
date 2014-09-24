@echo off
cls

title Installing N-able Remote Monitoring Software

REM SET	"server=%1"
REM SET	"customerID=%2"
SET	"installerLocation=%1"
REM SET	"minVersion=%4"
CALL \\%installerLocation%\agentparam.bat

SET	"alreadyInstalled=The N-able Agent is installed"
SET	"notInstalled=The N-able Agent is not yet installed, installing it now..."
SET	"programFiles=c:\program files"
SET	"cleanUp=Removing old Agent..."
SET	"missingArgs=One or more parameters were not specified"
SET	"counter=0"



REM       Check to see if its x86 or x64
IF %PROCESSOR_ARCHITECTURE% EQU  AMD64 ( SET "programFiles=%programFiles% (x86)" )

REM Debug Information
echo %server%
echo %customerID%
echo %installerLocation%
echo %programFiles%
echo %minVersion%

IF %server% == "" GOTO ERROR
IF %customerID% == "" GOTO ERROR
IF %installerLocation% == "" GOTO ERROR
IF %minVersion% == "" GOTO ERROR

\\%installerLocation%\AgentCleanup.exe %minVersion%
GOTO CONTINUE

:CONTINUE
IF %counter% == 2 GOTO END
IF NOT EXIST "%programFiles%\N-Able Technologies\Windows Agent\bin\agent.exe" ( GOTO INSTALL ) else ( GOTO AlreadyInstalled )
GOTO END

:INSTALL
echo %notInstalled%
\\%installerLocation%\WindowsAgentSetup.exe /s /v" /qn CUSTOMERID=%customerID% CUSTOMERSPECIFIC=1 SERVERPROTOCOL=HTTPS SERVERADDRESS=%server% SERVERPORT=443"
set /a counter=counter+1
GOTO CONTINUE

:AlreadyInstalled
echo %AlreadyInstalled%
GOTO END

:ERROR
echo %missingArgs%

:END
