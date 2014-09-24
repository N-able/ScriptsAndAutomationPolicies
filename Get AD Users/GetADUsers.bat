@echo off
rem Grabs a list of user accounts in the current Active Directory domain
rem and displays which ones are disabled.
rem by Tim wiser, Orchid IT (November 2011)


set DC=FALSE
set OUTPUTFILE=userlist2.txt
net share > %windir%\temp\orchid_sharelist.txt
for /f "tokens=1" %%d in ('type %windir%\temp\orchid_sharelist.txt') do IF %%d EQU SYSVOL set DC=TRUE
if %DC% EQU FALSE echo This script must be run on a domain controller. && exit

rem If we get this far then we're running on a DC

rem prepare files
set OUTPUTFILE=AD_User_List.txt
if EXIST %OUTPUTFILE% del %OUTPUTFILE%
echo. > %OUTPUTFILE%
echo. > %OUTPUTFILE%

rem get a list of user accounts and disabled user accounts
dsquery user -name * > allusers.txt
dsquery user -name * -disabled > disabledusers.txt

rem read the user list line by line and compare it to the disabled users list.  Where a match
rem is found, prefix the user line with an X.  Write the user line to an output file.
for /f "tokens=*" %%d in (allusers.txt) do CALL :CHECK_DISABLED %%d

rem we now have a list of users with X or A preceeding them so we can format it nicely
for /f "tokens=*" %%d in (users.txt) do call :PARSE_USER_LIST %%d
goto END


:CHECK_DISABLED
set LINE=%1%
set DISABLED=FALSE
for /f "tokens=*" %%d in (disabledusers.txt) do IF %LINE% EQU %%d set DISABLED=TRUE
if %DISABLED% EQU TRUE ( echo X_%LINE% >> users.txt
                                                ) else ( echo A_%LINE% >> users.txt )
goto :EOF


:PARSE_USER_LIST
set LINE=%1%
rem strip off the right hand side of the line into a file
for /f "tokens=1 delims=," %%d in ('echo %LINE%') do echo %%d > truncatedline.txt
rem strip off the start of the line and dump the remainder - the name - into a variable
for /f "tokens=* delims==" %%d in (truncatedline.txt) do call :GET_STATUS %%d
goto :EOF


:GET_STATUS
set LINE=%1%
set STATUS=%LINE:~0,1%
set NAME=%LINE:~6,255%
if %STATUS% EQU A echo %NAME% >> %OUTPUTFILE%
if %STATUS% EQU X echo %NAME% (disabled) >> %OUTPUTFILE%
goto :EOF

:END
type %OUTPUTFILE%
del allusers.txt
del disabledusers.txt
del truncatedline.txt
del users.txt
exit
