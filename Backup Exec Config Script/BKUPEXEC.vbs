'******************************************************************************************
'************************************ BKUPEXEC.vbs v1.15 **********************************
'****************************      Written By: Chris Reid      ****************************
'****************************       Date: July 15th, 2010      ****************************
'****************************   Copyright N-able Technologies  ****************************
'***  WARNING: This script should only be run by a Microsoft certified Administrator   ****
'***   N-able Technologies presents this script AS-IS - use at your own discretion      ***
'******************************************************************************************
 
' ***********************
' Version History
' ***********************

' Version 1.15
'   - Modified how the script detects a 64-bit machine vs. a 32-bit machine.


' Version 1.14
'   - Updated the registry commands - they weren't enabling the IPs in every scenario


'Version 1.13
'   - Added code to enable 'SQL Server and Windows Authentication' on the BKUPEXEC database instance


' Version 1.12
'   - Added support for x64 machines


' Version 1.11
'   - Removed Support for Backup Exec 10.x and older
'   - Rearranged the script into sub calls
'   - Added Windows Event Log logging
'   - Added support for running this script on devices that don't have a C:\ drive
'   - Altered the code so that the script doesn't always run *both* the MSDE and SQL Server commands 


 
 
 
 
' Declare the variables that are going to be used
Option Explicit
dim Shell, Username, Password, FSO, TextStream, strMSDEcmdline, strSQLcmdline, strComputer, objReg, strTemp, strKeyPath, strValueName, ScriptExec 
dim objStdOut, strLogFile, strValue, MSDEOutput, strTestKeyPath, Registry, arrValueNames, strInstancePath, Value, objWMIService, wbemFlagReturnImmediately
dim wbemFlagForwardOnly, colItems, objItem, AddressWidth


' Set known values for a few of the variables
Const FOR_APPENDING = 8
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
strValueName = "BKUPEXEC"
wbemFlagReturnImmediately = &h10
wbemFlagForwardOnly = &h20 




'We want to place the text file that contains the OSQL commands on the root drive, but we need to confirm whether that's C: or D: or something else
Set Shell = CreateObject("WScript.Shell")
strTemp = Shell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
Set Shell = Nothing
'Based on that %SYSTEMDRIVE% variable, let's set some of the other variables that this script will use
strLogFile = "" & strTemp & "\NABLEBKUPEXECqueries.sql"
strMSDEcmdline = "%COMSPEC% /c osql -E -S .\BKUPEXEC -i " & strLogFile
strSQLcmdline = "%COMSPEC% /c osql -E -S . -i " & strLogFile
 






' First - let's check to make sure that two command-line parameters were specified (we need a username and a password)
ParameterCheck

' Next - Check to see if the file we're going to create already exists. If it does exist, let's delete it
FileCheck

' Next - let's create the file
FileCreate

' Next - let's figure out what type of OS we have - 64-bit or 32-bit?
OSType

' Next - let's modify the registy values of the SQL database, so that the Windows Agent/Windows Probe can make ODBC queries to it
ChangeRegistry

' Next - let's run the OSQL commands that create the SQL user account we want
RunOSQL


' Final step - we need to stop/start the Backup Exec SQL service so that it will accept the new SQL account that we've created
RestartServices




'****************************************************************************************************************
' Sub: ParameterCheck
' Description: This sub checks to see whether or not two command line variables were passed to the script
' ***************************************************************************************************************
Sub ParameterCheck
  If WScript.Arguments.Count = 2 Then
    Set Shell = CreateObject("Wscript.Shell")   
    Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "You specified the username/password command line parameter."& chr(34),7,TRUE
    Set Shell = Nothing
   Username = WScript.Arguments.Item(0)
   Password = WScript.Arguments.Item(1)
  Else
   Set Shell = CreateObject("Wscript.Shell")   
    Shell.Run "%COMSPEC% /c eventcreate /t ERROR /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "You did not specify the username/password command line parameter."& chr(34),7,TRUE
   Set Shell = Nothing
   wscript.Quit(1)                 
  End If
End Sub
     

'******************************************************************************************************************************************
' Sub: FileCheck
' Description: This sub checks to see whether the text file has already been created. If it has been created, this sub will delete it
' ***************************************************************************************************************************************** 
Sub FileCheck
  Set FSO = CreateObject("Scripting.FileSystemObject")
  If FSO.FileExists(strLogFile) Then
     Set Shell = CreateObject("Wscript.Shell")   
     Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "The file already exists - we'll delete it now"& chr(34),7,TRUE
     Set Shell = Nothing
     FSO.DeleteFile(strLogFile)
  End If
  Set FSO = Nothing
End Sub
 
 
'******************************************************************************************************************************************
' Sub: FileCreate
' Description: This sub creates the text file that contains the OSQL commands
' *****************************************************************************************************************************************  
Sub FileCreate
  Set Shell = CreateObject("Wscript.Shell")   
  Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "Creating the log file that contains the OSQL commands"& chr(34),7,TRUE
  Set Shell = Nothing
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set TextStream = FSO.OpenTextFile(strLogFile, FOR_APPENDING, True)
  TextStream.WriteLine "use master"
  TextStream.WriteLine "EXEC sp_addlogin '"& Username &"','"& Password &"','BEDB'"
  TextStream.WriteLine "use BEDB"
  TextStream.WriteLine "EXEC sp_grantdbaccess '"& Username &"'"
  TextStream.WriteLine "EXEC sp_addrolemember 'db_datareader','"& Username &"'"
  TextStream.WriteLine "go"
  TextStream.WriteLine "exit"
  TextStream.Close
  Set FSO = Nothing
End Sub
 
 
 
'******************************************************************************************************************************************
' Sub: ChangeRegistry
' Description: This sub modifies the registry so that the SQL database Backup Exec is using will listen to external ODBC queries
' ***************************************************************************************************************************************** 
Sub ChangeRegistry 
                         

  ' 1. Figure out the internal name (usually MSSQL.1 or MSSQL.2) of the BKUPEXEC database instance
  Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
  & strComputer & "\root\default:StdRegProv")
  strInstancePath = "" & Registry & "Microsoft\Microsoft SQL Server\Instance Names\SQL"
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strInstancePath,strValueName,Value

 
  ' 2. Now that we know the instance under which Backup Exec is running, let's enable the NIC interfaces for that instance
  Set Shell = CreateObject("Wscript.Shell")   
  Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "Setting the registry settings for Backup Exec 11.x and 12.x"& chr(34),7,TRUE
  Shell. RegWrite"HKLM\" & Registry & "Microsoft\Microsoft SQL Server\" & Value & "\MSSQLServer\SuperSocketNetLib\TCP\IP1\Enabled", 1, "REG_DWORD"  'This enables the loopback IP (127.0.0.1)
  Shell. RegWrite"HKLM\" & Registry & "Microsoft\Microsoft SQL Server\" & Value & "\MSSQLServer\SuperSocketNetLib\TCP\IP2\Enabled", 1, "REG_DWORD"  'This enables the LAN IP (for example, 192.168.1.1)
  Shell. RegWrite"HKLM\" & Registry & "Microsoft\Microsoft SQL Server\" & Value & "\MSSQLServer\LoginMode", 2, "REG_DWORD"  'This changes the login mode to 'Windows Authentication and SQL Authentication' (we *need* SQL Authentication)
  Set Shell = Nothing
End Sub



 
 
'******************************************************************************************************************************************
' Sub: RunOSQL
' Description: This sub runs the appropriate OSQL commands to create a SQL DB account, using the information supplied in the text file.
' *****************************************************************************************************************************************  
Sub RunOSQL 
  'Run the OSQL commands for MSDE DB instance
  Set Shell = CreateObject("Wscript.Shell")   
  Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "Now running the OSQL commands for SQL Express"& chr(34),7,TRUE
  Set ScriptExec = Shell.Exec(strMSDEcmdline)
  Set objStdOut = ScriptExec.stdout
  While not objStdOut.AtEndOfStream
  MSDEOutput = WScript.StdOut.WriteLine(objStdOut.ReadAll())
  If InStrRev(MSDEOutput,">") = 0 Then 'This function says "if you find the word 'error' then run the SQL Server commands"
    Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "Now running the OSQL commands for SQL server."& chr(34),7,TRUE       
    Shell.Run strSQLcmdline,7,TRUE
  End If
  Set Shell = Nothing
  Wend
 
   
  'Cleanup Time - delete the SQL file that was created
  Set Shell = CreateObject("Wscript.Shell")   
  Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "Deleting the OSQL command file"& chr(34),7,TRUE
  Set Shell = Nothing
  Set FSO = CreateObject("Scripting.FileSystemObject")
  If FSO.FileExists(strLogFile) Then
    FSO.DeleteFile(strLogFile)
  End If
  Set FSO = Nothing

End Sub
 
 
'******************************************************************************************************************************************
' Sub: RestartServices
' Description: This sub stops the "SQL Server(BKUPEXEC)" Windows Services and it's child services, and then starts them.
' *****************************************************************************************************************************************  
Sub RestartServices
  Set Shell = CreateObject("Wscript.Shell")   
  Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "Stopping the Backup Exec Windows Services"& chr(34),7,TRUE
  
  Shell.Run "%COMSPEC% /c NET STOP MSSQL$BKUPEXEC /y",7,TRUE                 'This is the "SQL Server(BKUPEXEC)" service. The /y switch means stop all dependant services as well
  
  Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "Starting the Backup Exec Windows Services"& chr(34),7,TRUE
  
  Shell.Run "%COMSPEC% /c NET START BackupExecAgentBrowser",7,TRUE          ' This starts the 'Backup Exec Agent Browser' service

  Shell.Run "%COMSPEC% /c NET START BackupExecJobEngine",7,TRUE             ' This starts the 'Backup Exec Job Engine' service

  Shell.Run "%COMSPEC% /c NET START BackupExecRPCService",7,TRUE            ' This starts the 'Backup Exec Server' service

  Shell.Run "%COMSPEC% /c NET START BackupExecDeviceMediaService",7,TRUE    ' This starts the 'Backup Exec Device and Media Service' service

  Shell.Run "%COMSPEC% /c NET START MSSQL$BKUPEXEC",7,TRUE                  ' This starts the 'SQL Server (BKUPEXEC)' service
                
  Set Shell = Nothing
End Sub





' *****************************  
' Sub: OSType
' *****************************
Sub OSType
                       
  ' 1. Determine if this is a 32-bit machine or a 64-bit machine (as this will determine what registry values we modify)
  Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2")
  Set colItems = objWMIService.ExecQuery("SELECT AddressWidth FROM Win32_Processor where DeviceID='CPU0'", "WQL", _
  wbemFlagReturnImmediately + wbemFlagForwardOnly)

  For each objItem in colItems 
                AddressWidth = objItem.AddressWidth
  Next


  If AddressWidth = 64 Then
    'This is a 64-bit machine
    Registry = "SOFTWARE\Wow6432Node\"
    Set Shell = CreateObject("Wscript.Shell")   
    Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "This is a 64-bit machine."& chr(34),7,TRUE
    Set Shell = Nothing
  ElseIf AddressWidth = 32 Then
    'This is a 32-bit machine
    Registry = "SOFTWARE\"
    Set Shell = CreateObject("Wscript.Shell")   
    Shell.Run "%COMSPEC% /c eventcreate /t INFORMATION /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "This is a 32-bit machine."& chr(34),7,TRUE
    Set Shell = Nothing
  Else
    'Windows doesn't know what OS Type it's running
    Set Shell = CreateObject("Wscript.Shell")   
    Shell.Run "%COMSPEC% /c eventcreate /t WARNING /id 999 /l APPLICATION /so BKUPEXEC /d "& chr(34) & "Unable to determine if this is a 64-bit OS or a 32-bit OS."& chr(34),7,TRUE
    Set Shell = Nothing
  End If
End Sub 