' *************************************************************************************************************************************************
' Script: AVStatus.vbs
' Version: 1.86
' Author: Chris Reid / Marc-Andre Tanguay
' Description: This script checks the status of the A/V software installed on 
'              the machine, and writes data about the A/V software to the 
'              AntiVirusProduct WMI Class in the root\SecurityCenter WMI namespace.
' Date: August 18th, 2014
' Compatibility : It is tested on the current versions of windows but it also should work on all desktop and server versions, from Windows XP to Windows Server 2012.
' Usage in N-Central : avstatus.vbs WRITE     OR     avstatus.vbs DONOTWRITE   (the second option will write no data into WMI if an A/V product cannot be found)
' Usage in Windows Commant Prompt : CSCRIPT avstatus.vbs WRITE     OR     CSCRIPT avstatus.vbs DONOTWRITE   (the second option will write no data into WMI if an A/V product cannot be found)
' *************************************************************************************************************************************************


' Supported Anti-Virus Products:
'   - AVG 2012 (for Windows Vista/7/8 only - Server-class OS' are not supported)
'   - AVG 2013 (for Windows Vista/7/8 only - Server-class OS' are not supported)
'   - AVG 2014 (desktop and server
'   - Avira AntiVirus 12.x
'   - Avira AntiVirus 10.x (Server)
'   - ESET NOD32 Antivirus 4.x
'   - Kaspersky 6.0
'   - Kaspersky 8.0 
'   - Kaspersky 6.0 Enterprise
'   - Kaspersky 8.0 Enterprise
'   - Kaspersky 10.0
'   - Kaspersky 10.2
'   - Kaspersky Anti-Virus 2012
'   - Kaspersky Small Office Security 2
'   - McAfee AntiVirus 8.7 thru 8.8
'   - McAfee Security-As-A-Service 5.x
'   - Microsoft Defender
'   - Microsoft Security Essentials (MSE)
'   - Microsoft System Center Endpoint Protection (SCEP)
'   - Microsoft Forefront
'   - Sophos AntiVirus 9.x
'   - Symantec AntiVirus
'   - Symantec Endpoint Protection 11.x and 12.x
'   - Symantec Endpoint Security
'   - Symantec Endpoint.Cloud 20.x
'   - Symantec Endpoint Protection - SBE2013
'   - Total Defence r12
'   - Trend Micro Worry-Free Business Security 16
'   - Trend Micro Worry-Free Business Security 6.x
'   - Trend Micro Worry-Free Business Security 7.x
'   - Trend Micro Worry-Free Business Security 8.x
'   - Trend Micro Worry-Free Business Security Services
'   - Trend Micro OfficeScan
'   - Vipre Antivirus 4.x
'   - Vipre Enterprise Agent 4.x
'   - Vipre Antivirus Business 5.x
'   - Vipre Antivirus 2012
'   - F-Secure Client Security 8.x, 9.x
'   - F-Secure Protection Suite Business (PSB) 4.x
'   - Windows Defender
'   - Webroot SecureAnywhere
'   - Panda Cloud Endpoint Protection 6.11
'   - Avast 9.0


' Version History

' 1.87 - Added support for Avast

' 1.86 - Fixed issue with systems without AV and discovery of securitycenter2 (misconfigured state) and fixed new date formatting for Trend AV

' 1.85 - Updated AVG monitoring and fixed a few issues with AV

' 1.84 - Fixed issue with obtaining last update date for ESET NOD32 Antivirus

' 1.83 - Fixed date formatting issue with Kaspersky 10.x

' 1.82 - Added Windows Defender registry entry for av version date' - NOT RELEASED TO PUBLIC

' 1.81 - updated av definition file setting for Symantec Endpoint.Cloud (SEP Cloud) to report updated date. - NOT RELEASED TO PUBLIC

' 1.80 - added new registry entry for MS Forefront AV - May 7th 2014

' 1.79 - Added support for Kaspersky 10.2 - April 17th - 2014

' 1.78 - fixed issue with Microsoft Security Essentials on Windows 7 due to change in date formatting , and added MS SCEP - April 17th - 2014

' 1.77 - Added support for updated version of SEP Cloud 20.4 - February 26, 2014

' 1.76 - Added support for AVG 2014 - February 20, 2014

' 1.75 - Added support for Vipre AV date formatting change (added 22 instead of 21 on line 1687), and - October 15, 2013

' 1.72 - Added suport for Panda Cloud Endpoint Protection 6.11 - September 30th, 2013

' 1.71 - Adding support for Webroot SecureAnywhere & Fixed issue with Ms Forefront Date Formating - September 17th 2013

' 1.70 - fixing regional settings with Kaspersky - August 12, 2013

' 1.69 - Fixed date formatting issue with Kaspersky (lines 1950-1960), and fixed an issue with date formatting in windows defender AV - July 12, 2013

' 1.68 - Fixed issue with TrendMicro AV (Thanks Jean-Pierre) - July 11th, 2013

' 1.66 - Changed the registry key that is used to detect the AV Definition date of Trend WFBS (June 6th, 2013)

' 1.65 - Fixed another issue with detecting Trend WFBS 7 (April 29th, 2013)

' 1.64 - Added partial support for Windows Defender (April 8th, 2013)
'      - Fixed an issue with detecting Trend WFBS 7 on 64-bit machines. Thanks to Greg Michael for finding the issue!

' 1.63 - Added support for Kaspersky Endpoint Security 10 (March 4th, 2013)

' 1.62 - Fixed an issue where the version of Trend OfficeScan wasn't being correctly detected if the machine was switched from Conventional Scans to Smart Scans (March 4th, 2013)
'      - Added support for AVG 2013
'      - Added support for Trend Micro Worry-Free Business Security 8.x
'      - Added a new property in WMI that lists when the script was last ran. The AV Status service (in N-central) will need to be updated to support displaying this value.

' 1.61 - Updated the code for Trend Micro to go to the correct spot in the registry for both 32-bit and 64-bit machines. (Feb 4th, 2013)

' 1.60 - Fixed an issue where the fix in build 1.59 was missing a \ character in the path (Feb 4th, 2013)

' 1.59 - Fixed an issue where the script couldn't find the XML file for Kaspersky 8.0 Enterprise on Windows 2003 Servers (Jan 28th, 2013)

' 1.58 - Added support for Symantec Endpoint.Cloud (Symantec Endpoint Protection - SBE2013)

' 1.572 - fixed a date formatting issue with regional settings that was incorectly reporting wrong AV date

' 1.571 - for F-Secure, swapped FS@AQUA.INI to be first and FS@HYDRA.INI to be second for definition dating
' 1.57 - Added more descriptive name details for F-Secure AV (October 9 2012)

' 1.56 - Added support to return Version as well as AV Date in version field (October 2 2012)

' 1.55 - Added support for F-Secure Endpoint and Antivirus (October 1 2012)


' 1.54b - Changed how the script detects the A/V Definition Date for Kaspersky 8.0 for Servers - in some cases the value is not present in the registry (Sept. 20th, 2012)

' 1.53 - Fixed an issue where the script was returning a NULL 'VersionNumber' value on some installations of Trend Micro Worry Free  (July 17th, 2012)

' 1.52 - Modified the script so that if the DONOTWRITE parameter is specified, the script will clear out any data that exists in the AntiVirusProduct WMI class. (July 4th, 2012)

' 1.51 - Added support for Avira 10.x on Windows Servers, and Kaspersky Small Office Security 2 (June 27th, 2012)

' 1.50 - Fixed an issue where the script wasn't correctly determining the version of Symantec Endpoint Protection 12.x on x64 machines. (June 27th, 2012)
'      - Removed the AVG-specific code, and replaced it with code that checks the root\SecurityCenter2 WMI namespace for any installed A/V. While this only works on Vista, Windows 7 and Windows 8, it should make the script able to more reliabily detect A/V products

' 1.49 - Added a command line flag (possible values are WRITE and DONOTWRITE) that lets users choose whether or not the script should write data to WMI if no known A/V product is detected. (June 11th, 2012)
'      - Added support for AVG 2012
'      - Confirmed that the script supports Symantec Endpoint Protection 12.x (it previously had only been tested against Symantec Endpoint Protection 11.x)

' 1.48 - Added support for Avira AntiVirus 12.x (June 6th, 2012)

' 1.47 - Added support for Total Defense r12, and cleaned up some messy, unecessarily repetitive code that was writing text to 'standard out'. (May 29th, 2012)

' 1.46 - Fixed an issue where the Microsoft Security Essentials portion of the script reported an error with the RawAVDefDate variable. This issue only affected 64-bit machines. (May 24th, 2012)

' 1.45 - Added support for Kaspersky Small Office Security (May 18th, 2012)

' 1.44 - Fixed an issue in the Kaspersky SubRoutine where an incorrect registry key was being called. (May 15th, 2012)

' 1.43 - Forced the script to launch in the 32-bit CMD prompt so that it will properly detect Sophos. Thanks Jason Berg! (May 8th, 2012)

' 1.42 - Added support for Kaspersky Enterprise 6.0 for Windows Servers, and fixed how the script calculates the A/V Definition date for Kaspersky Enterprise 6.0 and 8.0 (May 7th, 2012)

' 1.41 - Updated how the script detects the server version of Kaspersky 8 and all versions of ESET, and how it detects the N-central Endpoint Security product. (March 30th, 2012)

' 1.40 - Fixed an issue that was preventing McAfee from being correct detected. Thansk Leon Boers! (March 8th, 2012)

' 1.39 - Fixed an issue that was preventing Symantec AntiVirus from being detected. Thanks Kyler Wolf! (Feb 20th, 2012)

' 1.38 - Added a 'Set' command that was preventing the Sophos portion of the script from launching successfully. Thanks Joe Sheehan! (Feb. 16th, 2012)

' 1.37 - Fixed a typo that prevented the script from launching successfully. (Feb 14th, 2012)

' 1.36 - Fixed an issue where an extra, unnecessary comma was preventing the Sophos portion of the script from running correctly. (Feb. 13th, 2012)
'      - Added support for McAfee Security-As-A-Service v5.x. Thanks Khaled Antar! (Feb 14th, 2012)
'      - Added support for Microsoft Forefront. Thanks Pat Albert! (Feb 10th, 2012)

' 1.35 - Added support for Kaspersky 8.0 Server Edition. Thanks Pat Albert! (Feb 10th, 2012)

' 1.34 - Added support for ESET NOD32 Antivirus 4.x. Thanks Leon Boers! (Feb 6th, 2012)

' 1.33 - Added support for GFI Vipre Antivirus Business 5.x and Vipre Antivirus 2012. Thanks Herb Meyerowitz! (Feb 2nd, 2012)

' 1.32 - Added support for Kaspersky 8.0. Thanks Pat Albert! (Jan 20th, 2011)
'      - The script now appends the version of Symantec Endpoint Protection to the 'AntiVirus Product Name' value

' 1.31 - Fixed an issue where the script failed to detect really old versions of Trend Micro OfficeScan. Thanks to David Lynnwood for the help! (Jan 17th, 2012)
'      - Fixed an issue where the script wasn't correctly detecting N-able's Endpoint Security product on Windows 7 machines. Thanks to James Clay for the help! (Jan 17th, 2012)

' 1.30 - Fixed an issue where the script reported the month of the A/V Definition Date for Symantec as 00 instead of 01 if the month was January. Thanks to Jonathan Baker for the help! (Jan 9th, 2012)

' 1.29 - Added a check for N-able's Endpoint Security product. If ES is found, the script exits immediately and doesn't write any values to WMI. This check was added because ES scans don't run properly when A/V data is stored in WMI. (Dec. 14th, 2011)
'      - Fixed bug in Trend Micro Worry-Free Business Security 7 when checking for definition version

' 1.28 - Added support for Microsoft Security Essentials (Nov. 30th, 2011)

' 1.27 - Added support for Kaspersky 6.0 (Nov. 7th, 2011)

' 1.26 - Added support for Kaspersky Anti-Virus 2012 (Oct. 18th, 2011)

' 1.25 - Fixed an issue where Symantec ES SBS wasn't being detected. Also, the A/V Security Center service will now report a Failed status if no A/V has been found - previously the service just went Misconfigured. (Oct. 18th, 2011)

' 1.24 - Updated the way that the script obtains Trend Worry Free 7 data (Oct. 7th, 2011)

' 1.23 - Changed how Vipre Enterprise and Vipre AV works - the script now grabs the install location from the registry (Oct. 6th, 2011)

' 1.22 - Added support for Vipre's Enterprise A/V product. Cleaned up the code by making a new 'CalculateAVAge' function that removes some unneeded duplicate lines of code (Oct 6th, 2011)

' 1.21 - Added a variable called OutOfDateDays that lets users configure how old the Definitions must be before the service will throw a Failure. (Oct 6th, 2011)

' 1.20 - Added more screen output when Trend is detected (Oct. 5th, 2011)

' 1.19 - Added support for Trend Worry-Free Business Security where the 'ProductName' registry value doesn't contain a version number (Oct. 5th, 2011)

' 1.18 - Fixed an issue where the script reported the version of McAffee being run instead of the version of the A/V DAT file (Sept. 23rd, 2011)

' 1.17 - Added support for McAfee AntiVirus version 8.5.x, and streamlined how some of the McAffee code was written (Sept. 22nd, 2011)

' 1.16 - Changed the command that was outputting text to the screen from wscript.echo to output.writeline - this will make the output show up when running the script in N-central. (Sept. 22nd, 2011)

' 1.15 - Added support for Mcafee (thanks Leon Boers!) and Vipre A/V (thanks Chris Jonas!) and fixed an issue with monitoring Sophos (Sept. 12th, 2011)

' 1.14 - Fixed a few issues are detecting and monitoring Trend WFBS 7 (May 27th, 2011)

' 1.13 - Added code sourced from Sophos to better support their product line (March 31st, 2011)

' 1.12 - Added support for 'Trend Micro Worry-Free Business Security Services', and made detection of Symantec more accurate (Mar. 30th, 2011)

' 1.11 - Added support for Trend Micro Worry Free Business 7 (Jan. 22nd, 2011)

' 1.10 - Added support for Sophos AntiVirus (Nov. 2nd, 2010)

' 1.9 - Fixed an issue where Trend wasn't getting detected on a Windows 2003 server (Oct. 5th, 2010)

' 1.8 - Fixed an issue where Symantec Endpoint Protection wasn't being detected (Sept. 7th, 2010)

' 1.7 - Added support for checking whether or not Real-Time Scanning is enabled for Symantec Anti-Virus (Aug. 11th, 2010)


' 1.6 - Added support for Trend Office Scan (July 8th, 2010)
'     - Changed how the script detects a 32-bit OS vs a 64-bit OS
'     - The script now logs to the 'AntiVirusProduct' WMI class in the root\SecurityCenter WMI namespace

' 1.5 - Added support for Symantec Endpoint Security (June 27th, 2010)


' 1.4 - Added checking for 32-bit vs. 64-bit operating systems (this affects which
'        registry key needs to be queried.)   (June 7th, 2010)


' 1.3 - Added checks to see if the InternalPatternVerKey for Trend is populated.


' 1.2 - Cast the PatternAge variable as a UINT32, instead of a string value.
'       This will allow users to threshold on that value. (May 20th, 2010)


' 1.1 - Added a check to make sure that the WMI Namespace exists before checking
'       for the presence of the WMI class. (May 20th, 2010)

' 1.0 - Initial Release (May 18th, 2010)





' Define the  variables used in the script
Option Explicit
dim HKEY_LOCAL_MACHINE, strComputer, objReg, strTrendKeyPath, InputRegistryKey1, InputRegistryKey2
dim oReg,arrSubKeys,SubKey
dim InputRegistryKey3, InstalledAV, NamespacePresense, objClassCreator, objGetClass, objWMIObject
dim RegstrValue, ReturneddwValue1, objWMIService, objItem, objNameSpace, ReturneddwValue2
dim wbemCimtypeString, wbemCimtypeUint32, colNamespaces, objNewNamespace
dim RawAVVersion, FormattedAVVersion, objNewInstance, strWMINamespace, ParentWMINamespace
dim ReturnedstrValue1, strValue, FormattedPatternAge, CalculatedPatternAge, CurrentDate
dim WMINamespace, strWMIClassWithQuotes, strWMIClassNoQuotes, colClasses, objClass 
dim strTestKeyPath, Registry, Registry32, Registry64, arrValueNames, strSymantecESKeyPath, WshShell
dim ReturnedBinaryArray1(7), bytevalue, i, f, SymantecAvPatternDate
dim SymantecAVMonth, SymantecAVYear, SymantecAVDate, ProductUpToDate, wbemCimtypeBoolean, OnAccessScanningEnabled, ReturneddwValue3, strTrendRealTimeKeyPath, InputRegistryKey4
dim AddressWidth, colItems, RawOnAccessScanningEnabled, Revision, strSymantecAVKeyPath, strSophosAVKeyPath
dim RawAVDate, strSymantecAVRealTimePath, strTrendPatternAgeKeyPath, strTrend7KeyPath, objComponentMgr, objConfigMgr
Dim strTrendVersionKeyPath, VIPREAvPatternDate, VIPREAVMonth, VIPREAVYear, VIPREAVDate, strVIPREAVRealTimePath, strVIPREESFolderPath
Dim McAfeeDatVersion, strMcAfeeVersionPath, McAfeeDatOAS, strMcAfeePath, InputRegistryKey5, strMcAfeeOASPath, OASEnabled, RAWOASEnabled
Dim strVIPREAVKeyPath, objFSO, ProgramFiles, Path, objFile, strLine, objNode, output, ProductVersion, AVDatDate, AVDatVersion, OutOfDateDays
Dim strVIPREEnterpriseKeyPath, InstallLocation, strTrendProductVersionKeyPath, strTrendAVVersionKeyPath, ReturneddwValue4, strKasperskyAV2012KeyPath, ReturneddwValue5
Dim strKaspersky2012AVDatePath, strKasperskyAV2012Path, strKasperskyAV60KeyPath, strKasperskyAV60DatePath, strKasperskyAV60Path, strKasperskyKES8KeyPath, strKasperskyKES8AVDatePath
Dim strSecurityEssentialsKeyPath, RawAVDefDate, strEndpointSecurityKeyPath   
Dim InputRegistryKey6, ProductName, strVIPREBusiness5KeyPath, strVIPREAV2012KeyPath, strESETKeyPath, DateStart
Dim strKasperskyKESServerKeyPath, strKasperskyKESServerAVDatePath, InstalledApp, ProductVersionKey, strForefrontKeyPath, strKasperskyKES8ServerKeyPath, strKasperskyKES6ServerKeyPath 
Dim Version, WSHStdOut, filename, cscriptExec, strKasperskySOS2KeyPath, Return, strTotalDefenseKeyPath, strWMIQuery
Dim strAviraKeyPath, NoAVBehavior, strWMINamespace2, HexProductState, HexScannerState, HexAVDefState, strAviraServerKeyPath, DeleteWMINamespace
Dim strFSecureRegPath0, strFSecureRegPath00, strFSecureRegPath1, strFSecureRegPath2, strFSecureRegPath3, strFSecureRegPath4, strFSecureRegPath5, strFSecureRegPath6, strFSecureRegPathLoc,strFSecureInstallPath
Dim strSEPCloudRegPath0, strSEPCloudRegpath1,strSEPCloudRegpath2,strSEPCloudRegpath3,strSEPCloudRegpath4,strSEPCloudDefPath, strSEPCloudDefPath2, OSName, strKES10KeyPath
Dim strPandaCloudEPPath64, strPandaCloudEPPath32, strPandaAVDefinitionPath
Dim strWindowsDefenderPath, DisableRealtimeMonitoring, strTrendDefVersionPath
Dim stravg2014regpath32, stravg2014regpath64, stravg2014defpath
dim lngBias, dtmDate,lngHigh,lngLow
Dim strWebRootStatusPath
Dim strTrendVerLen
Dim strAvastRegPath32, strAvastInstallPath, strAvastRegPath64

' Specify values for some of the variables
HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
wbemCimtypeString = 8
wbemCimtypeUint32 = 19
wbemCimtypeBoolean = 11
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20 
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strWMIClassWithQuotes = chr(34) & "AntiVirusProduct" & chr(34)
strWMIClassNoQuotes = "AntiVirusProduct"
strWMINamespace = "SecurityCenter"
strWMINamespace2 = "SecurityCenter2"
Set ParentWMINamespace = GetObject("winmgmts:\\" & strComputer & "\root")
Set WshShell = WScript.CreateObject("WScript.Shell")
Set output = Wscript.stdout
OutOfDateDays = 5
Version = "1.87"
InstallLocation = WshShell.ExpandEnvironmentStrings("%AllUsersProfile%")



'Lets see if the user specified whether or not the script should enter data into WMI if no A/V is found.
If WScript.Arguments.Count = 1 Then
  NoAVBehavior = WScript.Arguments.Item(0)
  If NoAVBehavior = "WRITE" Then
    output.writeline "- The " & NoAVBehavior & " flag has been specified as a command-line parameter."
  ElseIf NoAVBehavior = "DONOTWRITE" Then
    output.writeline "- The " & NoAVBehavior & " flag has been specified as a command-line parameter."   
  Else
    output.writeline "- An invalid command-line parameter has been specified. Please specify either WRITE or DONOTWRITE, or do not specify a command-line parameter at all."
    Wscript.Quit(0)
  End If  
Else
  output.writeline "- The command-line parameter (either WRITE or DONOTWRITE) for choosing whether or not to write data to WMI if an A/V product isn't found was not specified. This script will write data to WMI regardless of whether or not an A/V product is discovered."                
End If







	
'This is a meat of the script - where all of the functions are called.

  output.writeline "- This is version " & Version & " of the script." 

  OSType 'This function determines whether this is a 32-bit or 64-bit OS
  OSVersion 'This function figures out what OS the machine is running
  
  DetectInstalledAV   'This function will detect what AV software is installed
    
  If InstalledAV="Trend Micro Worry-Free Business Security 6" Then
    ObtainTrendMicroData 'Call the function we created to grab info about Trend Micro from the registry
    
  ElseIf InstalledAV="Trend Micro Worry-Free Business Security 7" Then
    ObtainTrend7AVData 'Call the function we created to grab info about Trend WFBS 7 from the registry

  ElseIf InstalledAV="Trend Micro Worry-Free Business Security" Then
    ObtainTrendMicroData 'Call the function we created to grab info about Trend WFBS 7 from the registry 
  
  ElseIf InstalledAV="Symantec Endpoint Protection" Then
   ObtainSymantecESData 'Call the function we created to grab info about Symantec Endpoint Security from the registry
  

  ElseIf InstalledAV="Trend Micro OfficeScan" Then
    ObtainTrendMicroData 'Call the function we created to grab info about Trend Micro from the registry  
  
  
  ElseIf InstalledAV="Symantec AntiVirus" Then
    ObtainSymantecAVData 'Call the function we created to grab info about Symantec AntiVirus from the registry

  
  ElseIf InstalledAV="Sophos AntiVirus" Then
    ' If the script is launcehed on a 64-bit machine, let's re-launch it in the 32-bit command prompt. This will allow the script to properly detect Sophos.
    ' Thanks to Jason Berg for this code snippet!
    Set WSHShell=CreateObject("Wscript.Shell")
    Set WSHStdOut=WScript.StdOut
    If WshShell.ExpandEnvironmentStrings("%processor_architecture%")="AMD64" then
      filename=Wscript.ScriptFullName
      output.writeline "- This script is being run on a 64-bit machine, so it'll be re-launched using the 32-bit version of cscript. This will ensure the proper discovery of Sophos."
      Set cscriptExec=WSHShell.Exec(WshShell.ExpandEnvironmentStrings("%windir%") & "\Syswow64\cscript.exe /nologo " & Chr(34) & filename & Chr(34))
      Do While cscriptExec.Status=0
        WScript.Sleep 100
        WSHStdout.WriteLine(cscriptExec.StdOut.ReadAll())
        WScript.StdErr.Write(cscriptExec.StdErr.ReadAll())
      Loop
      Wscript.Quit(cscriptExec.ExitCode)
    End If
    ObtainSophosAVData 'Call the function we created to grab info about Sophos AntiVirus from the registry

  ElseIf InstalledAV="Trend Micro Worry-Free Business Security Services" Then
    ObtainTrendMicroData 'Call the function we created to grab info about Trend Micro from the registry
    
  Elseif InstalledAV="McAfee AntiVirus" Then
    ObtainMcafeeAVData  
    
  ElseIf InstalledAV="VIPRE AntiVirus" Then
    ObtainVIPREAVData 'Call the function we created to grab info about VIPRE AntiVirus from the registry

  ElseIf InstalledAV="Sunbelt Enterprise Agent" Then
    ObtainVIPREEnterpriseData 'Call the function we created to grab info about VIPRE Enterprise from the registry

  ElseIf InstalledAV="Kaspersky Anti-Virus 2012" Then
    ObtainKaspersky2012AVData 'Call the function we created to grab info about Kaspersky from the registry


  ElseIf InstalledAV="Kaspersky Anti-Virus 6.0" Then
    ObtainKaspersky60AVData 'Call the function we created to grab info about Kaspersky from the registry



  ElseIf InstalledAV="Microsoft SCEP (or Security Essentials)" Then
    ObtainSecurityEssentialsAVData 'Call the function we created to grab info about MS Essentials from the registry
    
    
  ElseIf InstalledAV="Kaspersky Endpoint Security 8" Then
    ObtainKES8Data 'Call the function we created to grab info about Kaspersky Endpoint Security 8 from the registry

    
  ElseIf InstalledAV="VIPRE Business Antivirus" Then
    ObtainVIPREEnterpriseData 'Call the function we created to grab info about Vipre Business from the registry


  ElseIf InstalledAV="VIPRE Antivirus 2012" Then
    ObtainVIPREAVData 'Call the function we created to grab info about Vipre Antivirus 2012 from the registry
    
    
  ElseIf InstalledAV="ESET NOD32 Antivirus" Then
    ObtainESETAVData 'Call the function we created to grab info about Vipre Antivirus 2012 from the registry  
    
    
  ElseIf InstalledAV="Kaspersky Anti-Virus 8.0 For Windows Servers Enterprise Edition" Then
    strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KAVFSEE\8.0.0.0\"
    If InStr(OSName,2003) Then
      Path = InstallLocation & "\Application Data\Kaspersky Lab\KAV for Windows Servers Enterprise Edition\8.0\Update\u0607g.xml"
      output.writeline "- Windows 2003 has been detected. Using the following path to the Kaspersky XML file: " & Path
    Else
      Path = InstallLocation & "\Kaspersky Lab\KAV for Windows Servers Enterprise Edition\8.0\Update\u0607g.xml"
      output.writeline "- Using the following path to the Kaspersky XML file: " & Path
    End If     
    ObtainKESServerData 'Call the function we created to grab info about Kaspersky Endpoint Security 8 from the registry


  ElseIf InstalledAV="Kaspersky Anti-Virus 6.0 For Windows Servers Enterprise Edition" Then
    strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KAVFSEE\6.0.0.0\" 
    Path = InstallLocation & "\Kaspersky Lab\KAV for Windows Servers Enterprise Edition\6.0\Update\u0607g.xml"
    ObtainKESServerData 'Call the function we created to grab info about Kaspersky Endpoint Security 6 from the registry
 
 

  ElseIf InstalledAV="Kaspersky Small Office Security" Then
    ObtainKasperskySOSata 'Call the function we created to grab info about Kaspersky Endpoint Security 6 from the registry    
    

  ElseIf InstalledAV="Microsoft Forefront" Then
    ObtainSecurityEssentialsAVData 'Call the function we created to grab info about Microsoft Forefront from the registry    
          

  ElseIf InstalledAV="Total Defense R12 Client" Then
    ObtainTotalDefenseAVData 'Call the function we created to grab info about Total Defense from the registry

    
    
  ElseIf InstalledAV="Avira AntiVirus" Then
    ObtainAviraAVData 'Call the function we created to grab info about Avira from the registry
          
  ElseIf InStr(1,InstalledAV,"F-Secure")>0 Then
                ObtainFSecureAVData 'Call the function we created to grab info about F-Secure from registry and folder

  ElseIf InstalledAV="Symantec Endpoint.Cloud" then
                ObtainSEPCloudData
  
  ElseIf InstalledAV="Avast!" Then
				ObtainAvastData
                
'  ElseIf InStr(1,InstalledAV,"Symantec")>0 Then
'               ObtainFSecureAVData 'Call the function we created to grab info about F-Secure from registry and folder


  ElseIf InstalledAV="Kaspersky Endpoint Security 10" Then
    strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\10.1.0.0\" 

        Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
        oReg.EnumKey HKEY_LOCAL_MACHINE, Registry & "KasperskyLab\Components\34\Connectors\KES", arrSubKeys
        For Each subkey In arrSubKeys
            strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\" & subkey & "\"
	   	Next







    Path = InstallLocation & "\Kaspersky Lab\KES10\Data\u0607g.xml"
    ObtainKESServerData 'Call the function we created to grab info about Kaspersky from the registry

  ElseIf InstalledAV="Panda Endpoint Protection 10 32 Bit" Then
    strPandaAVDefinitionPath = "C:\Program Files\Panda Security\WaAgent\WalUpd\Data\Catalog"
    ObtainPandaCloudOfficeData 'Call the function we created to grab info about Kaspersky Endpoint Security 6 from the registry

  ElseIf InstalledAV="Panda Endpoint Protection 10 64 Bit" Then
    strPandaAVDefinitionPath = "C:\Program Files (x86)\Panda Security\WaAgent\WalUpd\Data\Catalog"
    ObtainPandaCloudOfficeData 'Call the function we created to grab info about Kaspersky Endpoint Security 6 from the registry
	
  ElseIf InstalledAV="Windows Defender" then
                ObtainWindowsDefenderData
                
  ElseIf InstalledAV="AVG 2014" Then
  				ObtainAVG2014Data
                
  ElseIf InstalledAV="Webroot SecureAnywhere" Then
	ObtainWebrootAnywhereAVData 'Call the function to grab info about webroot from the registry
  
  End If

  
      
  'Check to see if an instance of the WMI namespace exists; if it does, 
  'check to see if the WMI class exists. If the class exists, delete it, recreate it, and populate it
  If WMINamespaceExists(ParentWMINamespace,strWMINamespace) Then
      output.writeline "- The Namespace already exists."
      If WMIClassExists(strComputer,strWMIClassWithQuotes) Then
          output.writeline "- The WMI Class exists"
          WMINamespace.Delete strWMIClassNoQuotes
          CreateWMIClass
          PopulateWMIClass
      Else
          output.writeline "- The Namespace exists, but the WMI class does not. Curious." 
          CreateWMIClass
          PopulateWMIClass      
      End If
  Else
      'Create the WMI Namespace (if it doesn't already exist), the WMI Class, and populate the class with data.              
      output.writeline "- The WMI Namespace and Class does not exist"
      CreateWMINamespace
      CreateWMIClass
      PopulateWMIClass 
  End If    
                
                

  
' *****************************  
' Sub: OSType
' *****************************
Sub OSType
                       
  ' 1. Determine if this is a 32-bit machine or a 64-bit machine (as this will determine what registry values we modify)
Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT AddressWidth FROM Win32_Processor where DeviceID='CPU0'", "WQL", _
wbemFlagReturnImmediately + wbemFlagForwardOnly)

Registry32 = "SOFTWARE\"
Registry64 = "SOFTWARE\Wow6432Node\"
For each objItem in colItems 
                AddressWidth = objItem.AddressWidth
Next


  If AddressWidth = 64 Then
    'This is a 64-bit machine
    Registry = Registry64
    ProgramFiles = WshShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%") 'It's useful to know if we need to access C:Program Files or C:\Program Files(x86) - especially for Vipre A/V
    output.writeline "- This is a 64-bit machine."
  ElseIf AddressWidth = 32 Then
    'This is a 32-bit machine
    Registry = Registry32
    ProgramFiles = WshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")
    output.writeline "- This is a 32-bit machine."
  Else
    'Windows doesn't know what OS Type it's running
    output.writeline "- The type of OS is unknown - the script can't detect if it's 32-bit or 64-bit."
  End If
End Sub   



  
' *****************************  
' Sub: DetectInstalledAV
' *****************************
Sub DetectInstalledAV
    strMcAfeePath = Registry & "McAfee\AVEngine\DAT"
    strTrendKeyPath = Registry & "TrendMicro\PC-cillinNTCorp\CurrentVersion\Misc.\"
    strTrend7KeyPath = "SOFTWARE\TrendMicro\UniClient\1600\Update\PatternOutOfDateDays"
    strSymantecESKeyPath = Registry & "Symantec\Symantec Endpoint Protection\AV\"
    strSymantecAVKeyPath = Registry & "Symantec\Symantec AntiVirus\"
    strSophosAVKeyPath = Registry & "Sophos\"
    strVIPREAVKeyPath = Registry & "Sunbelt Software\VIPRE Antivirus\"
    strVIPREEnterpriseKeyPath = Registry & "Sunbelt Software\Sunbelt Enterprise Agent\"
    strKasperskyAV2012KeyPath = Registry & "KasperskyLab\protected\AVP12\settings\"
    strKasperskyAV60KeyPath = Registry & "KasperskyLab\protected\AVP80\settings\"
    strSecurityEssentialsKeyPath = "SOFTWARE\Microsoft\Microsoft Antimalware\"
    strEndpointSecurityKeyPath = Registry & "Microsoft\Windows\CurrentVersion\Uninstall\" 
    strKasperskyKES8KeyPath = Registry & "KasperskyLab\protected\KES8\settings\"
    strVIPREBusiness5KeyPath = Registry & "GFI Software\GFI Business Agent\"
    strVIPREAV2012KeyPath = Registry & "GFI Software\VIPRE Antivirus\" 
    strESETKeyPath = "SOFTWARE\ESET\ESET Security\CurrentVersion\"  
    strKasperskyKES8ServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KAVFSEE\8.0.0.0\"
    strKasperskyKES6ServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KAVFSEE\6.0.0.0\"
    strForefrontKeyPath = "SOFTWARE\Microsoft\Microsoft Forefront\Client Security\1.0\AM\"
    strKasperskySOS2KeyPath =  Registry & "KasperskyLab\protected\AVP9\settings\"
    strTotalDefenseKeyPath =  "SOFTWARE\CA\TDClient\"
    strAviraKeyPath = Registry & "Avira\AntiVir Desktop\"
    strAviraServerKeyPath = Registry & "Avira\AntiVir Server\"
                 strFSecureRegPath0= "SOFTWARE\Wow6432Node\Data Fellows\F-Secure\F-Secure GUI\PUB\" 
                  strFSecureRegPath00= "SOFTWARE\Data Fellows\F-Secure\F-Secure GUI\PUB\" 
                  strFSecureRegPath1= "SOFTWARE\Wow6432Node\Data Fellows\F-Secure\Anti-Virus\" 
                  strFSecureRegPath2= "SOFTWARE\Data Fellows\F-Secure\Anti-Virus\" 
                  strFSecureRegPath4= "SOFTWARE\Wow6432Node\F-Secure\Anti-Virus\" 
                  strFSecureRegPath3= "SOFTWARE\F-Secure\Anti-Virus\" 
                  strFSecureRegPath5= "SOFTWARE\Wow6432Node\Data Fellows\F-Secure\Anti-Virus Definition Databases\" 
                  strFSecureRegPath6= "SOFTWARE\F-Secure\Anti-Virus Definition Databases\" 
    strSEPCloudRegPath0= "SOFTWARE\Wow6432Node\Norton\"
                  strSEPCloudRegpath1= "SOFTWARE\Norton\"
                  strKES10KeyPath = Registry & "KasperskyLab\protected\KES10\settings\"
                  strWindowsDefenderPath = "SOFTWARE\Microsoft\Windows Defender\"
    	strPandaCloudEPPath64 = "Software\wow6432node\panda software\setup"
	strPandaCloudEPPath32 = "Software\panda software\setup"
    stravg2014regpath32 = "Software\Avg\Avg2014"
    stravg2014regpath64 = "Software\Wow6432Node\Avg\Avg2014"
    stravg2014defpath = "C:\ProgramData\AVG2014\avi"
	strAvastRegPath32 = "SOFTWARE\AVAST SOFTWARE\Avast"
    strAvastRegPath64 = "SOFTWARE\Wow6432Node\AVAST SOFTWARE\Avast"
    
	'-- Define Webroot Variables ---'
	strWebRootStatusPath = Registry & "WRData\Status\"
    

    'Check if N-able's Endpoint Security is installed - if it is, we should exit the script immediately (dumping data into WMI negatively affects ES' ability to run scans)
    'We need to check two different registry values, as it changes depending on what OS is installed.
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strEndpointSecurityKeyPath & "AVTC64","DisplayName",ReturnedstrValue1
    If ReturnedstrValue1 = "Endpoint Security Manager" Then
      output.writeline "- N-able's Endpoint Security product has been detected on this machine. This script will now exit."
      wscript.quit(1)
    End If 
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strEndpointSecurityKeyPath & "AVNT64","DisplayName",ReturnedstrValue1
    If ReturnedstrValue1 = "Endpoint Security Manager" Then
      output.writeline "- N-able's Endpoint Security product has been detected on this machine. This script will now exit."
      wscript.quit(1)
    End If  
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strEndpointSecurityKeyPath & "AVTC","DisplayName",ReturnedstrValue1
    If ReturnedstrValue1 = "Endpoint Security Manager" Then
      output.writeline "- N-able's Endpoint Security product has been detected on this machine. This script will now exit."
      wscript.quit(1)
    End If 
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strEndpointSecurityKeyPath & "AVNT","DisplayName",ReturnedstrValue1
    If ReturnedstrValue1 = "Endpoint Security Manager" Then
      output.writeline "- N-able's Endpoint Security product has been detected on this machine. This script will now exit."
      wscript.quit(1)
    End If  
  


    'Check to see what A/V product is installed
    If RegKeyExists("HKLM\" & strTrendKeyPath & "ProductName") Then
      strValue = "ProductName"
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strTrendKeyPath,strValue,RegstrValue
      InstalledAV = RegstrValue
      output.writeline "- " & InstalledAV & " has been detected."


    ElseIf RegKeyExists("HKLM\" & strTrendKeyPath & "ProgramVer") Then
      InstalledAV = "Trend Micro OfficeScan"
      output.writeline "- " & InstalledAV & " has been detected."

                                                                              

    Elseif RegKeyExists("HKLM\" & strSymantecESKeyPath & "ScanEngineVendor") Then
      InstalledAV = "Symantec Endpoint Protection"
      output.writeline "- " & InstalledAV & " has been detected."
      

    ElseIf RegKeyExists ("HKLM\" & strSymantecAVKeyPath & "CorporateFeatures") Then
       InstalledAV = "Symantec AntiVirus"
       output.writeline "- " & InstalledAV & " has been detected."
          
    ElseIf RegKeyExists ("HKLM\" & strSophosAVKeyPath) Then
       InstalledAV = "Sophos AntiVirus"
       output.writeline "- " & InstalledAV & " has been detected."

       
    ElseIf RegKeyExists ("HKLM\" & strTrend7KeyPath) Then 
       InstalledAV = "Trend Micro Worry-Free Business Security 7"
       output.writeline "- " & InstalledAV & " has been detected."

       
       
    ElseIf RegKeyExists ("HKEY_LOCAL_MACHINE\" & strMcAfeePath) Then
       InstalledAV = "McAfee AntiVirus"
       output.writeline "- " & InstalledAV & " has been detected."


    ElseIf RegKeyExists ("HKLM\" & strVIPREAVKeyPath & "ProductCode") Then
       InstalledAV = "VIPRE AntiVirus"
       output.writeline "- " & InstalledAV & " has been detected."



    ElseIf RegKeyExists ("HKLM\" & strVIPREEnterpriseKeyPath & "ProductCode") Then
       InstalledAV = "Sunbelt Enterprise Agent"
       output.writeline "- " & InstalledAV & " has been detected."

       
       
    ElseIf RegKeyExists ("HKLM\" & strKasperskyAV2012KeyPath & "SettingsVersion") Then
       InstalledAV = "Kaspersky Anti-Virus 2012"
       output.writeline "- " & InstalledAV & " has been detected."


       
    ElseIf RegKeyExists ("HKLM\" & strKasperskyAV60KeyPath & "SettingsVersion") Then
       InstalledAV = "Kaspersky Anti-Virus 6.0"
       output.writeline "- " & InstalledAV & " has been detected."

       
       
    ElseIf RegKeyExists ("HKLM\" & strSecurityEssentialsKeyPath & "InstallLocation") Then
       InstalledAV = "Microsoft SCEP (or Security Essentials)"
       output.writeline "- " & InstalledAV & " has been detected."

       
       
    ElseIf RegKeyExists ("HKLM\" & strKasperskyKES8KeyPath & "SettingsVersion") Then
       InstalledAV = "Kaspersky Endpoint Security 8"
       output.writeline "- " & InstalledAV & " has been detected."

       
    ElseIf RegKeyExists ("HKLM\" & strVIPREBusiness5KeyPath & "ProductCode") Then
       InstalledAV = "VIPRE Business Antivirus"
       strVIPREEnterpriseKeyPath = Registry & "GFI Software\GFI Business Agent\"
       output.writeline "- " & InstalledAV & " has been detected."
       
       
    ElseIf RegKeyExists ("HKLM\" & strVIPREAV2012KeyPath & "ProductCode") Then
       InstalledAV = "VIPRE Antivirus 2012"
       strVIPREAVKeyPath = Registry & "GFI Software\VIPRE Antivirus\"
       output.writeline "- " & InstalledAV & " has been detected."
       
            
    ElseIf RegKeyExists ("HKLM\" & strESETKeyPath & "\info\ProductName") Then
       InstalledAV = "ESET NOD32 Antivirus"
       output.writeline "- " & InstalledAV & " has been detected."
       
       
    ElseIf RegKeyExists ("HKLM\" & strKasperskyKES8ServerKeyPath & "ProdDisplayName") Then
       InstalledAV = "Kaspersky Anti-Virus 8.0 For Windows Servers Enterprise Edition"
       output.writeline "- " & InstalledAV & " has been detected."
       

    ElseIf RegKeyExists ("HKLM\" & strKasperskyKES6ServerKeyPath & "ProdDisplayName") Then
       InstalledAV = "Kaspersky Anti-Virus 6.0 For Windows Servers Enterprise Edition"
       output.writeline "- " & InstalledAV & " has been detected."
       
       
    ElseIf RegKeyExists ("HKLM\" & strForefrontKeyPath & "InstallLocation") Then
       InstalledAV = "Microsoft Forefront"
       output.writeline "- " & InstalledAV & " has been detected."
       
       
    ElseIf RegKeyExists ("HKLM\" & strKasperskySOS2KeyPath & "Ins_DisplayName") Then
       InstalledAV = "Kaspersky Small Office Security"
       output.writeline "- " & InstalledAV & " has been detected."


    ElseIf RegKeyExists ("HKLM\" & strTotalDefenseKeyPath & "ProductName") Then
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strTotalDefenseKeyPath,ProductName,InstalledAV
      InstalledAV = "Total Defense R12 Client"
      output.writeline "- " & InstalledAV & " has been detected."
       
       
    ElseIf RegKeyExists ("HKLM\" & strAviraKeyPath & "EngineVersion") Then
       InstalledAV = "Avira AntiVirus"
       output.writeline "- " & InstalledAV & " has been detected."
       
       
    ElseIf RegKeyExists ("HKLM\" & strAviraServerKeyPath & "EngineVersion") Then
       InstalledAV = "Avira AntiVirus"
       output.writeline "- " & InstalledAV & " has been detected."           
      
                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath0 & "ProductName") Then
       objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPath0,"ProductName",InstalledAV
                    'strFSecureRegPathLoc=strFSecureRegPath1 
       output.writeline "- " & InstalledAV & " has been detected."           

                                                If RegKeyExists ("HKLM\" & strFSecureRegPath1 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath1 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath2 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath2 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath3 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath3 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath4 & "InstallationDirectory") Then
                                                   strFSecureRegPathLoc=strFSecureRegPath4 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath5 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath5 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath6 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath6 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath1 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath1 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath2 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath2 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath3 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath3 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath4 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath4 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath5 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath5 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath6 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath6 
                                                end if
                   ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath00 & "ProductName") Then
       objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPath00,"ProductName",InstalledAV
                    'strFSecureRegPathLoc=strFSecureRegPath1
       output.writeline "- " & InstalledAV & " has been detected."           
                                                                                output.writeline "00"

                                                If RegKeyExists ("HKLM\" & strFSecureRegPath1 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath1 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath2 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath2 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath3 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath3 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath4 & "InstallationDirectory") Then
                                                   strFSecureRegPathLoc=strFSecureRegPath4 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath5 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath5 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath6 & "InstallationDirectory") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath6 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath1 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath1 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath2 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath2 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath3 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath3 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath4 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath4 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath5 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath5 
                                                ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath6 & "Path") Then
                                                                strFSecureRegPathLoc=strFSecureRegPath6 
                                                end if


'NEWSEP
                
    ElseIf RegKeyExists ("HKLM\" & strSEPCloudRegPath0 ) Then
    
                                                Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

                                                oReg.EnumKey HKEY_LOCAL_MACHINE, strSEPCloudRegPath0, arrSubKeys
                                                For Each subkey In arrSubKeys
								                   If RegKeyExists("HKLM\Software\Norton\" & subkey & "\PRODUCTVERSION") Then
                                                	output.WriteLine SubKey 
                                                    strSEPCloudRegPath2 = subkey 
											   end If
											   	Next
											   	
                                                oReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Wow6432Node\Norton\", arrSubKeys
                                                For Each subkey In arrSubKeys
											   	output.WriteLine SubKey
								                   If RegKeyExists("HKLM\Software\Wow6432Node\Norton\" & subkey & "\SharedDefs\AVDEFMGR") Then
                                                	output.WriteLine SubKey 
                                                    strSEPCloudRegPath2 = subkey 
											   end If
                                                                
                                                Next
                
                                                strSEPCloudRegPath2 = strSEPCloudRegPath0 & strSEPCloudRegPath2 
                                                InstalledAV="Symantec Endpoint.Cloud"

                ElseIf RegKeyExists ("HKLM\" & strSEPCloudRegPath1 ) Then
                                                Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
                                                oReg.EnumKey HKEY_LOCAL_MACHINE, strSEPCloudRegPath1, arrSubKeys
                                                For Each subkey In arrSubKeys
								                   If RegKeyExists("HKLM\Software\Norton\" & subkey & "\PRODUCTVERSION") Then
                                                	output.WriteLine SubKey  & "32"
                                                    strSEPCloudRegpath2 = subkey 
												   end if
                                                Next
                
                                                oReg.EnumKey HKEY_LOCAL_MACHINE, strSEPCloudRegPath1, arrSubKeys
                                                For Each subkey In arrSubKeys
								                   If RegKeyExists("HKLM\Software\Norton\" & subkey & "\SharedDefs\AVDEFMGR") Then
                                                	output.WriteLine SubKey  & "32"
                                                    strSEPCloudRegpath2 = subkey 
												   end if
                                                Next
                
                                                strSEPCloudRegPath2 = strSEPCloudRegPath1 & strSEPCloudRegPath2 
                                                InstalledAV="Symantec Endpoint.Cloud"
                                                

         
                                                
  ElseIf RegKeyExists ("HKLM\" & strKES10KeyPath & "SettingsVersion") Then
      InstalledAV = "Kaspersky Endpoint Security 10"
      output.writeline "- " & InstalledAV & " has been detected."  
  

  ElseIf RegKeyExists ( "HKLM\" & strPandaCloudEPPath32 & "\InitialProductName") Then
      InstalledAV = "Panda Endpoint Protection 10 32 Bit"
      output.writeline "- " & InstalledAV & " has been detected."			

  ElseIf RegKeyExists ( "HKLM\" & strPandaCloudEPPath64 & "\InitialProductName") Then
      InstalledAV = "Panda Endpoint Protection 10 64 Bit"
      output.writeline "- " & InstalledAV & " has been detected."			

  ElseIf RegKeyExists ( "HKLM\" & stravg2014regpath32 & "\ProdType") Then
      InstalledAV = "AVG 2014"
      output.writeline "- " & InstalledAV & " has been detected."			

  ElseIf RegKeyExists ( "HKLM\" & stravg2014regpath64 & "\ProdType") Then
      InstalledAV = "AVG 2014"
      output.writeline "- " & InstalledAV & " has been detected."			

  '--- Check for Webroot Anywhere ---	  
  ElseIf RegKeyExists ("HKLM\" & strWebRootStatusPath) Then
	  InstalledAV = "Webroot SecureAnywhere"
      output.writeline "- " & InstalledAV & " has been detected."

  '--- Check for AVAST! ---	  
  ElseIf RegKeyExists ("HKLM\" & strAvastRegPath32 & "\Version") Then
	  InstalledAV = "Avast!"
	  objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strAvastRegPath32,"ProgramFolder",strAvastInstallPath  
	  output.writeline "- " & strAvastInstallPath
      output.writeline "- " & InstalledAV & " has been detected."

  '--- Check for AVAST! ---	  
  ElseIf RegKeyExists ("HKLM\" & strAvastRegPath64 & "\Version") Then
	  InstalledAV = "Avast!"
	  objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strAvastRegPath64,"ProgramFolder",strAvastInstallPath  
	  output.writeline "- " & strAvastInstallPath
      output.writeline "- " & InstalledAV & " has been detected."

      
  ElseIf WMINamespaceExists(ParentWMINamespace,strWMINamespace2) Then
       output.writeline "- Unable to find a known A/V product on this device. As a last-ditch effort, let's look in the root\SecurityCenter2 WMI Namespace."
       ' This next line calls the 'ObtainSecurityCenter2Data sub-routine, which queries the root\SecurityCenter2 WMI Namespace for info.
       ObtainSecurityCenter2Data
                           
  ElseIf RegKeyExists ("HKLM\" & strWindowsDefenderPath & "ProductStatus") Then
      InstalledAV = "Windows Defender"
      output.writeline "- " & InstalledAV & " has been detected."                  
	  

  Else
    If NoAVBehavior = "WRITE" Then
      output.writeline "- Unable to determine installed AV."
      InstalledAV = "No AV installed"
      onAccessScanningEnabled = "FALSE"
      ProductUpToDate = "FALSE"
      FormattedAVVersion = "Unknown"
    ElseIf NoAVBehavior = "DONOTWRITE" Then
      output.writeline "- The script could not detect an A/V product on this device. As the DONOTWRITE command-line parameter was specified, no data will be written to WMI, and this script will now exit."
      ' We need to clear out the existing data from the AntiVirusProduct WMI Class
      If WMIClassExists(strComputer,strWMIClassWithQuotes) Then
        output.writeline "- The WMI Class exists. Deleting it with great prejudice."
        WMINamespace.Delete strWMIClassNoQuotes
        output.writeline "- Now recreating the AntiVirusProduct WMI class. It's a thing of beauty."
        CreateWMIClass
      End If 
      wscript.quit(0)
    ElseIf NoAVBehavior = "" Then
      output.writeline "- Unable to determine installed AV."
      InstalledAV = "No AV installed"
      onAccessScanningEnabled = FALSE
      ProductUpToDate = FALSE
      FormattedAVVersion = "Unknown"    
    End If                         
                                  
  End If
End Sub  


' *****************************  
' Function: ObtainTrendMicroData
' *****************************
Sub ObtainTrendMicroData
    'Grab the A/V Pattern Version from the Registry
    strTrendRealTimeKeyPath = Registry & "TrendMicro\PC-cillinNTCorp\CurrentVersion\Real Time Scan Configuration\"
    strTrendDefVersionPath = Registry & "TrendMicro\PC-cillinNTCorp\CurrentVersion\iCRC Scan\"
    InputRegistryKey1 =  "InternalPatternVer"
    InputRegistryKey2 =  "PatternDate"
    InputRegistryKey3 = "InternalNonCrcPatternVer"
    InputRegistryKey4 = "Enable"
    InputRegistryKey5 =  "NonCrcPatternDate"
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strTrendKeyPath,InputRegistryKey1,ReturneddwValue1
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strTrendKeyPath,InputRegistryKey3,ReturneddwValue2
     objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strTrendRealTimeKeyPath,InputRegistryKey4,ReturneddwValue3

   If ReturneddwValue3 = 1 Then
                OnAccessScanningEnabled = TRUE
   ElseIf ReturneddwValue3 = 0 Then
                OnAccessScanningEnabled = FALSE
   End If
    
    output.writeline "- The unformatted A/V Definition Version is: " & ReturneddwValue1
    output.writeline "- The unformatted alternate A/V Version Date is: " & ReturneddwValue2
    output.writeline "- The unformatted date of the A/V Definition File is: " & ReturnedstrValue1
    output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled   

    
    
    ' If Smart Scan is installed, the version will be reported in the InternalNonCrcPatternVer key, 
    ' and the InternalPatternVer key will be 0.
    ' If Smart Scan isn't installed, the version will be reported in the InternalPatternVer key, 
    ' and the InternalNonCrcPatternVer key will be 0.
    ' This IF statements checks to see if the InternalPatternVer key is equal to 0; if it is, then the
   ' script will use the InternalNonCrcVer key instead.
	

   
    If ReturneddwValue1 = 0 Then
        RawAVVersion =  cstr(ReturneddwValue2)
        objReg.GetStringValue HKEY_LOCAL_MACHINE,strTrendKeyPath,InputRegistryKey5,ReturnedstrValue1
    Else
        output.writeline "- Using the InternalNonCrcPatternVer registry value. This is the Smart Scan definition version."
        RawAVVersion =  ReturneddwValue1
        objReg.GetStringValue HKEY_LOCAL_MACHINE,strTrendKeyPath,InputRegistryKey2,ReturnedstrValue1
    End If 
    
    output.writeline "- The unformatted A/V Definition Version is: " & ReturneddwValue1
    output.writeline "- The unformatted alternate A/V Version Date is: " & ReturneddwValue2
    output.writeline "- The unformatted date of the A/V Definition File is: " & ReturnedstrValue1   
    
	
	'Convert the dwValue variable from hex to decimal, and pretty it up
	strTrendVerLen = Len(ReturneddwValue1)
	If strTrendVerLen = 6 Then
		FormattedAVVersion =  Left(RawAVVersion,1) & "." & Mid(RawAVVersion,2,3) & "." & Right(RawAVVersion,2)
    ElseIf strTrendVerLen = 7 Then
		FormattedAVVersion =  Left(RawAVVersion,2) & "." & Mid(RawAVVersion,3,3) & "." & Right(RawAVVersion,2)
	End If
	output.writeline "- The version string length is: " &strTrendVerLen
	   output.writeline "- The formatted A/V Definition Version is: " & FormattedAVVersion

    'Format the Pattern Age that was in the registry to make it more human-readable
    FormattedPatternAge = Left(ReturnedstrValue1,4) & "/" & Mid(ReturnedstrValue1,5,2) & "/" & Right(ReturnedstrValue1,2)
    'output.writeline "" & FormattedAVVersion
    
    'Calculate how old the A/V Pattern really is
    CurrentDate = Now
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate) 
    FormattedAVVersion = FormattedAVVersion & " - " & FormattedPatternAge
    
    
    output.writeline "- The formatted A/V Version is: " & FormattedAVVersion
    output.writeline "- The formatted date of the A/V Definition File is: " & FormattedPatternAge
    output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
    
    CalculateAVAge 'Call the function to determine how old the A/V Definitions are   

End Sub




' *****************************  
' Function: ObtainSymantecESData
' *****************************
Sub ObtainSymantecESData
    'Grab the Symantec data from the Registry
    strSymantecESKeyPath =  Registry & "Symantec\Symantec Endpoint Protection"
    InputRegistryKey1 =  "PatternFileDate"
    InputRegistryKey2 =  "ScanEngineVersion"
    InputRegistryKey3 = "OnOff"
    InputRegistryKey4 = "PatternFileRevision"
    InputRegistryKey5 = "PRODUCTVERSION"
    InputRegistryKey6 = "PRODUCTNAME"
    objReg.GetBinaryValue HKEY_LOCAL_MACHINE,strSymantecESKeyPath & "\AV",InputRegistryKey1,ReturnedBinaryArray1
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strSymantecESKeyPath & "\AV",InputRegistryKey2,RawAVVersion
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strSymantecESKeyPath & "\AV\Storages\FileSystem\RealTimeScan",InputRegistryKey3,RawOnAccessScanningEnabled
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strSymantecESKeyPath & "\AV",InputRegistryKey4,Revision
    
    'If it's Symantec ES 12, the ProductVersion value is stored in one location. If it's Symantec ES 11, it's stored in a different location.
    'Let's check to see if it's in the location for ES 12 first.
     If RegKeyExists ("HKLM\" & Registry & "\Symantec\Symantec Endpoint Protection\CurrentVersion\PRODUCTVERSION") Then  'It's version 12 or newer!
      objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSymantecESKeyPath & "\CurrentVersion",InputRegistryKey5,ProductVersion
      objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSymantecESKeyPath & "\CurrentVersion",InputRegistryKey6,ProductName
     Else 
      objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSymantecESKeyPath & "\SMC",InputRegistryKey5,ProductVersion
      ProductName = "Symantec Endpoint Protection"
     End If
      
    
    
    'Set the right 'InstalledAV' value (it should include the product name and version)
    InstalledAV = ProductName & " " & ProductVersion
    output.writeline InstalledAV & " has been detected on this machine."
    
    
    'Let's figure out if Real Time Scanning is enabled or not
    If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
    ElseIf RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
    End If
    
    
    'The A/V Definition File Date is a binary value - let's format things appropriately.
    SymantecAVYear = ReturnedBinaryArray1(0) + 1970 ' Need to add 1970 to the value seen in the registry
    
    'Need to convert the month from it's decimal value to a human-readable value
    SymantecAVMonth = ReturnedBinaryArray1(1)
            output.writeline "The month is: " & SymantecAVMonth
    If SymantecAvMonth = "0" Then
        SymantecAvMonth = "01"
    ElseIf SymantecAvMonth = "1" Then 
        SymantecAvMonth = "02" 
    ElseIf SymantecAvMonth = "2" Then 
        SymantecAvMonth = "03"    
    ElseIf SymantecAvMonth = "3" Then 
        SymantecAvMonth = "04"    
    ElseIf SymantecAvMonth = "4" Then 
        SymantecAvMonth = "05"    
    ElseIf SymantecAvMonth = "5" Then 
        SymantecAvMonth = "06"    
    ElseIf SymantecAvMonth = "6" Then 
        SymantecAvMonth = "07"    
    ElseIf SymantecAvMonth = "7" Then 
        SymantecAvMonth = "08"    
    ElseIf SymantecAvMonth = "8" Then 
        SymantecAvMonth = "09"   
    ElseIf SymantecAvMonth = "9" Then 
        SymantecAvMonth = "10"    
    ElseIf SymantecAvMonth = "10" Then 
        SymantecAvMonth = "11"    
    ElseIf SymantecAvMonth = "11" Then
        SymantecAvMonth = "12"
    End If
         
    SymantecAVDate = ReturnedBinaryArray1(2)
    ' Let's add a zero in front of the date value if it's less than 10. We like pretty things.
    If SymantecAvDate = "1" Then 
        SymantecAvDate = "01" 
    ElseIf SymantecAvDate = "2" Then 
        SymantecAvDate = "02"    
    ElseIf SymantecAvDate = "3" Then 
        SymantecAvDate = "03"    
    ElseIf SymantecAvDate = "4" Then 
        SymantecAvDate = "04"    
    ElseIf SymantecAvDate = "5" Then 
        SymantecAvDate = "05"    
    ElseIf SymantecAvDate = "6" Then 
        SymantecAvDate = "06"    
    ElseIf SymantecAvDate = "7" Then 
        SymantecAvDate = "07"    
    ElseIf SymantecAvDate = "8" Then 
        SymantecAvDate = "08"   
    ElseIf SymantecAvDate = "9" Then
            SymantecAvDate = "09"
    End If

    FormattedPatternAge =  SymantecAVYear & "/" & SymantecAvMonth & "/" & SymantecAvDate
    output.writeline "- The formatted date of the A/V Definition File is: " & FormattedPatternAge
    
    FormattedAVVersion = FormattedPatternAge & " r" & Revision
    
    output.writeline "- The formatted A/V Version is: " & FormattedAVVersion
    
    'Calculate how old the A/V Pattern really is
    CurrentDate = Now
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
    output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
    
    
    CalculateAVAge 'Call the function to determine how old the A/V Definitions are 
    
End Sub




' *****************************  
' Function: ObtainSymantecAVData
' *****************************
Sub ObtainSymantecAVData
    'Grab the Symantec data from the Registry
    strSymantecAVKeyPath =  Registry & "Symantec"
    strSymantecAVRealTimePath = Registry & "INTEL\LANDesk\VirusProtect6\CurrentVersion\Storages\Filesystem\Realtimescan"
    InputRegistryKey1 =  "DEFWATCH_10"
    objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSymantecAVKeyPath & "\SharedDefs",InputRegistryKey1,RawAVDate

    RawAVDate = Right(RawAVDate,12)
    FormattedPatternAge = Left(RawAVDate,4) & "/" & Mid(RawAVDate,5,2) & "/" & Mid(RawAVDate,7,2)
    FormattedAVVersion = FormattedPatternAge & " r" & Right(RawAVDate,1) 
    output.writeline FormattedAVVersion
    
    
    CurrentDate = Now
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
    output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
    
    CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  
    
    
    'Let's figure out if Real Time Scanning is enabled or not
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strSymantecAVRealTimePath,"OnOff",RawOnAccessScanningEnabled
    If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
    ElseIf RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
    End If
    output.writeline OnAccessScanningEnabled
    


End Sub



 
 
 
 
 
 
 ' *****************************  
' Function: ObtainSophosAVData
' *****************************
Sub ObtainSophosAVData
' Let's call the Sophos script/function so that the registry gets populated
                SophosScript

' Let's figure out when Sophos was last updated
                Set objComponentMgr = CreateObject("Infrastructure.ComponentManager.1")
                Set objConfigMgr = objComponentMgr.FindComponent("ConfigurationManager")
                Set objNode = objConfigMgr.GetNode(2, "ProductInfo/updateDate")      
                FormattedPatternAge = objNode.GetAttributeValue("year") & "/" & objNode.GetAttributeValue("month") & "/" & objNode.GetAttributeValue("day")   
                output.writeline   "Update time:" & FormattedPatternAge
    CurrentDate = Now
        CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
        output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
    
        CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub



 
' *****************************  
' Function: ObtainTrend7AVData
' *****************************
Sub ObtainTrend7AVData
    'Grab the A/V Pattern Version from the Registry
    strTrendRealTimeKeyPath = "SOFTWARE\TrendMicro\UniClient\1600\Scan\Real Time\"
    strTrendPatternAgeKeyPath = "SOFTWARE\TrendMicro\UniClient\1600\Update\"
    strTrendProductVersionKeyPath = "SOFTWARE\TrendMicro\UniClient\1600\"
    strTrendAVVersionKeyPath = "TrendMicro\UniClient\1600\Component\"
    InputRegistryKey1 =  "LastUpdateTime"
    InputRegistryKey2 = "Enable"
    InputRegistryKey3 = "ProgramVer"
    InputRegistryKey4 = "c3t4"
    InputRegistryKey5 = "c3t1208090624"
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strTrendPatternAgeKeyPath,InputRegistryKey1,ReturneddwValue1
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strTrendRealTimeKeyPath,InputRegistryKey2,ReturneddwValue2
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strTrendProductVersionKeyPath,InputRegistryKey3,ReturneddwValue3
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strTrendAVVersionKeyPath,InputRegistryKey5,FormattedAVVersion
    
    ' In some cases the 'Component' Registry Key doesn't exist, and depending on the version it might be on two different locations in the registry. Let's check if it is exists; if it does, then cool, and if not we'll need to fabricate some information.
    If RegKeyExists ("HKLM\" & Registry32 & strTrendAVVersionKeyPath) Then
	  objReg.GetStringValue HKEY_LOCAL_MACHINE,Registry32 & strTrendAVVersionKeyPath,InputRegistryKey4,ReturneddwValue4
      objReg.GetStringValue HKEY_LOCAL_MACHINE,Registry32 & strTrendAVVersionKeyPath,InputRegistryKey5,ReturneddwValue5
      If Len(ReturneddwValue4) <> 0 Then
        FormattedAVVersion = ReturneddwValue4
      Else
        FormattedAVVersion = ReturneddwValue5
      End If
    ElseIf RegKeyExists ("HKLM\" & Registry64 & strTrendAVVersionKeyPath) Then
	  objReg.GetStringValue HKEY_LOCAL_MACHINE,Registry64 & strTrendAVVersionKeyPath,InputRegistryKey4,ReturneddwValue4
      objReg.GetStringValue HKEY_LOCAL_MACHINE,Registry64 & strTrendAVVersionKeyPath,InputRegistryKey5,ReturneddwValue5
      If Len(ReturneddwValue4) <> 0 Then
        FormattedAVVersion = ReturneddwValue4
      Else
        FormattedAVVersion = ReturneddwValue5
      End If
    Else	
      output.writeline "- The appropriate registry key for determining the A/V Definition version was not found. As a result, the 'VersionNumber' property returned by this script will be set to 'Unknown.'"
      FormattedAVVersion = "Unable to determine the version of A/V definitions being run on this device." 
    End If


   If ReturneddwValue2 = 1 Then
                OnAccessScanningEnabled = TRUE
   ElseIf ReturneddwValue2 = 0 Then
                OnAccessScanningEnabled = FALSE
   End If
   
   'We need to convert the 'LastUpdateTime' value from Unix/Epoch date-time to the current date (and remove the time - we don't care about it)
   CalculatedPatternAge = DateAdd("s", ReturneddwValue1, "01/01/1970 00:00:00")
   CalculatedPatternAge = DateDiff ("d", CDate(CalculatedPatternAge), Date)
     
   CalculateAVAge 'Call the function to determine how old the A/V Definitions are
   
   InstalledAV = "Trend Micro Worry-Free Business Security 7 (Version: " & ReturneddwValue3 & ")" 

   output.writeline "- Product version: " & InstalledAV
   output.writeline "- Version: " & FormattedAVVersion
   output.writeline "- Is Real-Time Scanning Enabled? " & OnAccessScanningEnabled
   output.writeline "- Is the Product up-to-date? " & ProductUpToDate  


End Sub








' *****************************  
' Sub: CreateWMINamespace
' *****************************
Sub CreateWMINamespace
    Set objItem = ParentWMINamespace.Get("__Namespace")
    Set objNewNamespace = objItem.SpawnInstance_    
    objNewNamespace.Name = strWMINamespace
    objNewNamespace.Put_
End Sub





' *****************************  
' Sub: CreateWMIClass
' *****************************
Sub CreateWMIClass
    Set objWMIService = GetObject("Winmgmts:root\" & strWMINamespace)
    Set objClassCreator = objWMIService.Get() 'Load the Namespace           
    'Define the Properties of the WMI Class
    objClassCreator.Path_.Class = "" & strWMIClassNoQuotes
    
    objClassCreator.Properties_.add "Displayname", wbemCimtypeString
    objClassCreator.Properties_.add "onAccessScanningEnabled", wbemCimtypeBoolean
    objClassCreator.Properties_.add "ProductUpToDate", wbemCimtypeBoolean
    objClassCreator.Properties_.add "VersionNumber", wbemCimtypeString
    objClassCreator.Properties_.add "ScriptExecutionTime", wbemCimtypeString
                
                
    ' Make the 'InstalledAV' property a 'key' (or index) property
    objClassCreator.Properties_("Displayname").Qualifiers_.add "key", true
                
    ' Write the new class to the 'root\SecurityCenter' namespace in the repository
    objClassCreator.Put_

End Sub
    
    
' *****************************  
' Function: WMINamespaceExists
' Thanks to http://www.cruto.com/resources/vbscript/vbscript-examples/misc/wmi/List-All-WMI-Namespaces.asp for this code 
' *****************************
Function WMINamespaceExists(ParentWMINamespace,WMINamespace)
                WMINamespaceExists = vbFalse
                Set colNamespaces = ParentWMINamespace.InstancesOf("__Namespace")
                For Each objNamespace In colNamespaces
                      If instr(objNamespace.Path_.Path,WMINamespace) Then
                            WMINamespaceExists = vbTrue                        
                      End if
                Next
                Set colNamespaces = Nothing
End Function


  


' *****************************  
' Function: WMIClassExists
' Thanks to http://gallery.technet.microsoft.com/ScriptCenter/en-us/a1b23364-34cb-4b2c-9629-0770c1d22ff0 for this code 
' *****************************
Function WMIClassExists(strComputer, strWMIClassWithQuotes)
                WMIClassExists = vbFalse
                Set WMINamespace = GetObject("winmgmts:\\" & strComputer & "\root\" & strWMINamespace)
                Set colClasses = WMINamespace.SubclassesOf()
                For Each objClass In colClasses
                      If instr(objClass.Path_.Path,strWMIClassNoQuotes) Then
                            WMIClassExists = vbTrue
                      End if
                Next
                Set colClasses = Nothing
End Function  
  
     

    
' *****************************  
' Sub: PopulateWMIClass
' *****************************
Sub PopulateWMIClass    
    'Create an instance of the WMI class using SpawnInstance_
    Set WMINamespace = GetObject("winmgmts:\\" & strComputer & "\root\" & strWMINamespace)
    Set objGetClass = WMINamespace.Get(strWMIClassNoQuotes)
    Set objNewInstance = objGetClass.SpawnInstance_
    objNewInstance.VersionNumber = FormattedAVVersion & " - " & AVDatVersion
    objNewInstance.displayName = InstalledAV
    objNewInstance.onAccessScanningEnabled = OnAccessScanningEnabled
    objNewInstance.ProductUpToDate = ProductUpToDate
    objNewInstance.ScriptExecutionTime = Date & " " & Time
    output.writeline "- Populating the WMI Class with the data."
    
    ' Write the instance into the WMI repository
    objNewInstance.Put_()
End Sub



' *****************************  
' Function: RegKeyExists
' Returns a value (true / false)
' Thanks to http://www.tek-tips.com/faqs.cfm?fid=5864 for this code
'************************************
Function RegKeyExists (RegistryKey)

    'If there isn't the key when we read it, it will return an error, so we need to resume
    On Error Resume Next

    'Try reading the key
                WshShell.RegRead(RegistryKey)

    'Catch the error
    Select Case Err
      'Error Code 0 = 'success'
      
      Case 0:
        RegKeyExists = true
      'Any other error code is a failure code
      
      Case Else:
        RegKeyExists = false
    End Select

    'Turn error reporting back on
    On Error Goto 0

End Function


' *****************************  
' Function: SophosScript
' All of the code below this comment has been written by Sophos. Thanks to Sean Richmond @ Sophos for his help!
'************************************
Sub SophosScript


on error resume next


Dim properties  ' SavControl properties

' columns in the array
const SAV_PROPERTY_NAME  = 0
const REG_KEY_NAME       = 1
const REG_KEY_VALUE_NAME = 2
const THE_VALUE          = 3
const VALUE_TYPE         = 4

' value types
const BOOL_VALUE  = 1
const DWORD_VALUE = 2
const STRING_VALUE= 3

                                                                '(SavControl property name,                      key to create,                value name,                default, type )
properties  = Array(_   
                                                                Array("Version.Major",                                 "SavService\Version",         "Major", "-1",                                              DWORD_VALUE), _
                Array("Version.Minor",                                                 "SavService\Version",         "Minor", "-1",                                              DWORD_VALUE), _
                Array("Version.AV.Data",                            "SavService\Version",         "Data",  "",                                                    STRING_VALUE), _
                Array("Version.Extra",                                   "SavService\Version",         "Extra", "",                                                    STRING_VALUE), _
                Array("Protection.OnAccess",    "SavService\Status\Policy",   "OnAccessScanningEnabled", "-1", BOOL_VALUE) _
                  )



Dim oReg
Set oReg=GetObject("winmgmts:{impersonationlevel=delegate}!\\.\root\default:StdRegProv")
If Err.number <> 0 Then
    output.writeline "- Error opening registry"
    Wscript.Quit
End If

'------------------------
' Get SAV data
'-------------------------
Dim LastUpdatedStr
Dim IDEcount
Dim SavUpdatingStatus
Dim canAccessSAV

LastUpdatedStr = ""
IDEcount = -1
canAccessSAV = false

'Check if SAV is not being updated
SavUpdatingStatus = GetSAVUpdatingStatus(oReg)

If SavUpdatingStatus <> "-1" Then
    If SavUpdatingStatus = "1" Then
        ' updating - don't try to access SAV
        output.writeline "Update in progress"
    Else
        canAccessSAV = true         'any other status means not updating
    End If
End If

If canAccessSAV Then
    SavUpdatingStatus = "0" 'reset the value for MSP
    
    'try to get values
    Dim value
    Dim sc
    Set sc = CreateObject( "SAVControl.SophosAntiVirusControl" )
    If Err.number = 0 Then
        'Values availabe directly from SAVControl
        Dim i
        i =0
        Dim p
        For Each p In properties
            value = ""
            value = sc.GetProperty(p(SAV_PROPERTY_NAME))
            If Err.number = 0 And CStr(value) <> "" Then
                properties(i)(THE_VALUE) = value
                                output.writeline "got " & p(SAV_PROPERTY_NAME) 
                                output.writeline value
            Else
                output.writeline "Failed to get " & p(SAV_PROPERTY_NAME) & Err.number
            End If
            i = i+1
        Next
    Else
        output.writeline "- Failed to create SavControl object"
    End If
    
    ' (16) = LastUpdated
    value = ""
    value = GetLastUpdatedTimeOfSAV(oReg)
    If value <> "" Then
        LastUpdatedStr= value
    End If
End If 'not updating

'set formatted AV Version and onaccess scannning state
                FormattedAVVersion = left(properties(0)(THE_VALUE),2) & "." & right(properties(1)(THE_VALUE),1) & " virus data " & properties(2)(THE_VALUE) 
                output.writeline FormattedAVVersion
                OnAccessScanningEnabled = properties(4)(THE_VALUE)
    output.writeline OnAccessScanningEnabled
    


Err.number = 0
output.writeline "- SophosScript finished"


End Sub




'*************************************************************************
' GetDwordRegValue(regObject, key, valueName)
'*************************************************************************
Function GetDwordRegValue(regObject, key, valueName)
    Dim value
    value = &H0FFFFFFFF
    regObject.GetDWORDValue HKEY_LOCAL_MACHINE, key, valueName, value
    If Err.number <> 0 Or Not(IsNumeric(value)) Then 'it seems that GetDWordValue doesn't set error number always
        value=&H0FFFFFFFF
        output.writeline "GetDWORDValue " & valueName & " in " & key & " failed"
    End If
    GetDwordRegValue = value
End Function

const UPDATE_IN_PROGRESS = "1"
const UPDATE_STATUS_UNKNOWN = "-1"

'***************************************************************************
' GetSAVUpdatingStatus(registryObject) returns the "UpdateState" value for SAV
'***************************************************************************
Function GetSAVUpdatingStatus(regObject)
    GetSAVUpdatingStatus = UPDATE_STATUS_UNKNOWN
    Dim value
    value = GetDwordRegValue(regObject, "SOFTWARE\Sophos\SAVService\UpdateStatus", "UpdateState")
    If Err.number = 0 And CStr(value) <> "" Then 
        GetSAVUpdatingStatus = value 'we have a value, call succeeded
    End If
End Function

'***************************************************************************
' GetLastUpdatedTimeOfSAV() returns last updated time as "dd.mm.yy hh:mm:ss"
'   The last updated time is not currently available through SAV control,
'   so it is read from SAV configuration using SAVService directly
'***************************************************************************
Function GetLastUpdatedTimeOfSAV(regObject)

    GetLastUpdatedTimeOfSav = ""
    Dim SavUpdatingStatus
    SavUpdatingStatus = GetSAVUpdatingStatus(regObject)
    
    If SavUpdatingStatus <> UPDATE_IN_PROGRESS And SavUpdatingStatus <> UPDATE_STATUS_UNKNOWN Then    ' if not updating
        Dim componentManager
        Err.Clear 
        Set componentManager = CreateObject("Infrastructure.ComponentManager")
        If Err.number = 0 Then
            Dim configMgr 
            Set configMgr = componentManager.FindComponent("ConfigurationManager")
            If Err.number = 0 Then
                Dim node 
                Set node = configMgr.GetNode(2, "ProductInfo/updateDate") 
                If Err.number = 0 Then
                    Dim dateString
                    dateString = ""
                    'Ignore any errors below
                    dateString = node.GetAttributeValue("day") & "." &_
                                 node.GetAttributeValue("month") & "." &_
                                 node.GetAttributeValue("year")
                      
                    Dim timeString
                    timeString = ""
                    timeString = node.GetAttributeValue("hour") & ":" &_
                                 node.GetAttributeValue("minute") & ":" &_
                                 node.GetAttributeValue("second")    
                            
                    GetLastUpdatedTimeOfSAV = dateString & " " & timeString
                End If
            End If
        End If
    End If
End Function


' *****************************
' Sub: Mcafee
' *****************************

Sub ObtainMcafeeAVData

    InputRegistryKey1 = "AVDatDate"
    InputRegistryKey2 = "OASEnabled"
    InputRegistryKey3 = "bNetshieldEnabled"
    InputRegistryKey5 = "OASState"
    
    ' Need to check if this is McAfee Security As a Service or not
    objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,Registry & "McAfee\ManagedServices\Agent","szAppName",InstalledApp
    If InstalledApp = "McAfee Security-as-a-Service" Then
      output.writeline InstalledApp & " has been detected!"
      strMcAfeePath = Registry & "McAfee\ManagedServices\VirusScan\"
      strMcAfeeVersionPath = Registry & "McAfee\ManagedServices\Agent\"
      strMcAfeeOASPath = Registry & "Mcafee\DesktopProtection\"
      ProductVersionKey = "szMyUsrSrvVersion"
      'Let's get the status of On-Access Scanning (OAS)
      objReg.GetDWORDValue HKEY_LOCAL_MACHINE,Registry & "McAfee\SystemCore\vscore\On Access Scanner\McShield\Configuration\","OASEnabled",RAWOASEnabled
      If RAWOASEnabled = 3 Then
         OnAccessScanningEnabled = TRUE
      Else
         OnAccessScanningEnabled = FALSE
      End If
      output.writeline "OAS is in a " & OnAccessScanningEnabled & " state."
      InstalledAV = InstalledApp
      InputRegistryKey4 = "AVDatVer"
      objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strMcAfeePath,InputRegistryKey4,AVDatVersion
    Else
      ProductVersionKey = "szProductVer"
      output.writeline "- " & InstalledApp & " was not found."
      strMcAfeePath = Registry & "McAfee\AVEngine\"
      strMcAfeeVersionPath = Registry & "McAfee\DesktopProtection\"
      strMcAfeeOASPath = Registry & "Mcafee\DesktopProtection\"
      'Lets find the state of the OAS  
      If RegKeyExists ("HKLM\Software\McAfee\VScore\LockDownEnabled") Then  'It's version 8.7 or older!
        objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strMcAfeeOASPath,InputRegistryKey3,RAWOASEnabled
          If RAWOASEnabled = 1 Then
            OnAccessScanningEnabled = TRUE
          Else
            OnAccessScanningEnabled = FALSE
         End If
      Else 'It's at least version 8.8!
        output.writeline "- It seems to be at least Mcafee version 8.8"
        objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strMcAfeeVersionPath,InputRegistryKey5,RAWOASEnabled
          If RAWOASEnabled = 3 Then
            OnAccessScanningEnabled = TRUE
          Else
            OnAccessScanningEnabled = FALSE
        End If
      End If
      output.writeline "- OAS is in a " & OnAccessScanningEnabled & " state."
      InstalledAV = "McAfee Anti Virus"
      InputRegistryKey4 = "AVDatVersion"
      objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strMcAfeePath,InputRegistryKey4,AVDatVersion
    End If

 

    objReg.GetStringValue HKEY_LOCAL_MACHINE,strMcAfeePath,InputRegistryKey1,AVDatDate
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strMcAfeeVersionPath,ProductVersionKey,ProductVersion
    
    
    FormattedAVVersion = AvDatDate & " SDAT: " & AVDatVersion
    
    output.writeline "- The version of the A/V Definition File is: " & FormattedAVVersion
    output.writeline "- The version of McAfee AntiVirus running on this machine is: " & ProductVersion
    InstalledAV = InstalledAV & " " & ProductVersion
    
    
    'Calculate how old the A/V Pattern really is
    CurrentDate = Now
    CalculatedPatternAge = DateDiff("d",AVDatDate,CurrentDate)
    output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
  
    CalculateAVAge 'Call the function to determine how old the A/V Definitions are
    
    

End Sub  



' *****************************  
' Sub: ObtainVIPREAVData
' *****************************
Sub ObtainVIPREAVData
  'Find out where Vipre has been installed
  objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strVIPREAVKeyPath,"InstallPath",InstallLocation

  'Grab the VIPRE data from the text file from the A/V install path
  Path = InstallLocation & "Definitions\DefVer.txt"
  Set objFSO = CreateObject("Scripting.FileSystemObject") 
  Set objFile = objFSO.OpenTextFile(Path, 1) 
  Set f = objFSO.GetFile(Path)

  Dim arrFileLines() 
  i = 0 
  Do Until objFile.AtEndOfStream 
  Redim Preserve arrFileLines(i) 
  arrFileLines(i) = objFile.ReadLine 
  i = i + 7 
  Loop 
  objFile.Close 
 
  'Then you can iterate it like this 
 
  For Each strLine in arrFileLines 
    FormattedPatternAge = Right(f.DateLastModified,21)
  Next
      
    CurrentDate = Now
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
   
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  

  'Let's figure out if Real Time Scanning is enabled or not
  strVIPREAVRealTimePath = "SYSTEM\CurrentControlSet\Services\SBAMSvc"
  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strVIPREAVRealTimePath,"Start",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
  ElseIf RawOnAccessScanningEnabled = 2 Then
      OnAccessScanningEnabled = TRUE
  End If

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  'Let's figure out the version number
  strValue = "Version"
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strVIPREAVKeyPath,strValue,RegstrValue
  FormattedAVVersion = RegstrValue
  output.writeline "- The version of Vipre A/V running is: " & FormattedAVVersion

End Sub


' *************************************  
' Function: ObtainVIPREEnterpriseData
' *************************************
Function ObtainVIPREEnterpriseData
  'Find out where Vipre has been installed
  objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strVIPREEnterpriseKeyPath,"InstallPath",InstallLocation

  'Grab the VIPRE data from the text file from the A/V install path
  Path = InstallLocation & "Definitions\DefVer.txt"

  Set objFSO = CreateObject("Scripting.FileSystemObject") 
  Set objFile = objFSO.OpenTextFile(Path, 1) 
  Set f = objFSO.GetFile(Path)

  Dim arrFileLines() 
  i = 0 
  Do Until objFile.AtEndOfStream 
  Redim Preserve arrFileLines(i) 
  arrFileLines(i) = objFile.ReadLine 
  i = i + 7 
  Loop 
  objFile.Close 
 
  'Then you can iterate it like this 
 
  For Each strLine in arrFileLines 
    FormattedPatternAge = Right(f.DateLastModified,22)
  Next
      
      CurrentDate = Now
CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
     
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  

  'Let's figure out if Real Time Scanning is enabled or not
  strVIPREAVRealTimePath = "SYSTEM\CurrentControlSet\Services\SBAMSvc"
  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strVIPREAVRealTimePath,"Start",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
  ElseIf RawOnAccessScanningEnabled = 2 Then
      OnAccessScanningEnabled = TRUE
  End If

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  'Let's figure out the version number
  strValue = "Version"
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strVIPREEnterpriseKeyPath,strValue,RegstrValue
  FormattedAVVersion = RegstrValue
  output.writeline "- The version of Vipre A/V running is: " & FormattedAVVersion

End Function




' *****************************  
' Sub: CalculateAVAge
' *****************************
Sub CalculateAVAge

  If CalculatedPatternAge < OutOfDateDays Then
      ProductUpToDate = TRUE
  Else
      ProductUpToDate = FALSE
  End If

End Sub


' *************************************  
' Sub: ObtainKaspersky2012AVData
' *************************************
Sub ObtainKaspersky2012AVData
  strKaspersky2012AVDatePath = Registry & "KasperskyLab\protected\AVP12\"
  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyAVDatePath & "Data\","LastSuccessfulUpdate",AVDatVersion
  
  ' The AVDatVersion value is a UNIX timestamp - so this script needs to convert it to a human-readable format before continuing.
    CurrentDate = Now
  FormattedPatternAge = DateAdd("s", AVDatVersion, #1/1/1970#)
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time Kasperky was updated was " & CalculatedPatternAge & " days ago."


     
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  

  'Let's figure out if Real Time Scanning is enabled or not
  ' NOTE: At this point I can't figure this out - no registry value seems to change when you enable/disable this feature in Kaspersky.
  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyAVDatePath & "settings\def\","EnableSelfProtection",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
  ElseIf RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
  End If

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  
  
  'Let's figure out the version number
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strKaspersky2012AVDatePath & "settings\def\","SettingsVersion",FormattedAVVersion
  output.writeline "- The version of Kaspersky running is: " & FormattedAVVersion

End Sub



' *************************************  
' Sub: ObtainKaspersky60AVData
' *************************************
Sub ObtainKaspersky60AVData
  strKasperskyAV60DatePath = Registry & "KasperskyLab\protected\AVP80\"
  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyAV60DatePath & "Data\","LastSuccessfulUpdate",AVDatVersion
  
  ' The AVDatVersion value is a UNIX timestamp - so this script needs to convert it to a human-readable format before continuing.
    CurrentDate = Now
  FormattedPatternAge = DateAdd("s", AVDatVersion, #1/1/1970#)
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time Kasperky was updated was " & CalculatedPatternAge & " days ago."


     
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  

  'Let's figure out if Real Time Scanning is enabled or not
  ' NOTE: At this point I can't figure this out - no registry value seems to change when you enable/disable this feature in Kaspersky.
  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyAV60DatePath & "settings\def\","EnableSelfProtection",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
  ElseIf RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
  End If

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  
  
  'Let's figure out the version number
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyAV60DatePath & "settings\","SettingsVersion",FormattedAVVersion
  output.writeline "- The version of Kaspersky running is: " & FormattedAVVersion

End Sub





' *************************************  
' Sub: ObtainSecurityEssentialsAVData
' *************************************
Sub ObtainSecurityEssentialsAVData
  ' In testing, it's been found that MS Security Essentials (or MS Forefront) *always* writes to this path - it doesn't change if the machine is 32-bit or 64-bit.
  ' So, to handle that scenario, this script will just use one path - which is different than every other A/V application.
  If InstalledAV = "Microsoft Forefront" Then
    strSecurityEssentialsKeyPath = "Software\Microsoft\Microsoft Forefront\Client Security\1.0\AM\"
  Else
    strSecurityEssentialsKeyPath = "Software\Microsoft\Microsoft Antimalware\"
  End If
  
 
  
  ' Let's grab what A/V definition version Security Essentials is using
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strSecurityEssentialsKeyPath & "Signature Updates\","AVSignatureVersion",FormattedAVVersion
  output.writeline "- The A/V version " & InstalledAV & " is running is: " & FormattedAVVersion
  
  
  ' Let's figure out if Real-Time Scanning is enabled or not
'  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strSecurityEssentialsKeyPath & "Real-Time Protection\","DisableRealTimeMonitoring",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = FALSE
  Else
      OnAccessScanningEnabled = TRUE
  End If
  
  output.writeline "- Is Real-Time Scanning enabled? " & OnAccessScanningEnabled
  
  
  ' Lets grab the date when the A/V definitions were last updated. Since it's stored as a REG_BINARY value, we'll need to convert it.
  ' Thanks to http://www.tek-tips.com/viewthread.cfm?qid=1189987 for the conversion code!
  objReg.GetBinaryValue HKEY_LOCAL_MACHINE,strSecurityEssentialsKeyPath & "Signature Updates\","AVSignatureApplied",RawAVDefDate
  
                                dim lngBias, dtmdate, lngHigh, lngLow
                                lngBias=0 'determine it per your link on your own
                                lngHigh=0
                                lngLow=0
                                For i=7 to 4 step -1
                                  lngHigh=lngHigh*256+RawAVDefDate(i)
                                Next
                                
    For i=3 to 0 step -1
                                  lngLow=lngLow*256+RawAVDefDate(i)
                                Next
                                
                                If err.number<>0 Then
                                  dtmDate = #1/1/1601#
                                  err.clear
                                Else
                                  If lngLow < 0 Then
                                    lngHigh = lngHigh + 1
                                  End If
                                  If (lngHigh = 0) And (lngLow = 0 ) Then
                                    dtmDate = #1/1/1601#
                                  Else
                                    dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
                                    + lngLow)/600000000 - lngBias)/1440
                                  End If
                                End If
                                Dim intdatespace
                                If InStr(dtmDate," ") >0 Then
	                                'this shows you what you get
	                                intdatespace=InStr(dtmDate," ")
	                                dtmDate = CDate(Left(dtmdate,intdatespace))
	                                output.writeline "- The converted date is: " & dtmDate
                                Else 
	                                'this shows you what you get
	                                dtmDate = CDate(Left(dtmdate,10))
	                                output.writeline "- The converted date is: " & dtmDate
                                
                                End If
                                
                                
' - VERSION 1.71 - changing the date formattting to fix regional issue                                
'      CurrentDate = Now
  CurrentDate = Year(Now) & "/" & Month (Now) & "/" & Day (Now)
CalculatedPatternAge = DateDiff("d",dtmDate,CurrentDate)
  output.writeline "- The last time " & InstalledAV & " was updated was " & CalculatedPatternAge & " days ago."                                


     
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are


End Sub


' *************************************  
' Sub: ObtainKES8Data
' *************************************
Sub ObtainKES8Data
  strKasperskyKES8AVDatePath = Registry & "KasperskyLab\protected\KES8\"
  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyKES8AVDatePath & "Data\","LastSuccessfulUpdate",AVDatVersion
  
  ' The AVDatVersion value is a UNIX timestamp - so this script needs to convert it to a human-readable format before continuing.
    CurrentDate = Now
  FormattedPatternAge = DateAdd("s", AVDatVersion, #1/1/1970#)
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time Kasperky was updated was " & CalculatedPatternAge & " days ago."



  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

                Dim oShell, oExec, sLine, sExecPath

                Set oShell = CreateObject("WScript.Shell")
                set oExec = oShell.Exec(ProgramFiles & "\Kaspersky Lab\Kaspersky Endpoint Security 8 for Windows\avp.com status FM")

                sLine = oExec.StdOut.ReadLine


                If InStr(sLine, "running") <> 0 Then
                                OnAccessScanningEnabled = TRUE
                Else
                                OnAccessScanningEnabled = FALSE
                End If 
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  
  
  'Let's figure out the version number
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyKES8AVDatePath & "settings\def\","SettingsVersion",FormattedAVVersion
  output.writeline "- The version of Kaspersky running is: " & FormattedAVVersion

End Sub

 
 
 
 
' *************************************  
' Sub: ObtainESETAVData
' *************************************
Sub ObtainESETAVData
  objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strESETKeyPath&"\info","ProductVersion",ProductVersion
  
  ' Let's append the version number to the name of the installed AV product.
  InstalledAV = InstalledAV & " " & ProductVersion
  output.writeline "- The installed version of ESET NOD32 is: " & InstalledAV
  
  ' Let's grab the A/V Definition version
  objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strESETKeyPath&"\info","ScannerVersion",AVDatVersion
  ' That A/V Definition version is composed of two pieces of data - the build number and the date (   i.e. 6762 (20120102)   ). We need to separate them out.
  DateStart = InStr(AVDatVersion,"(") ' This determines the position of the "(" character
  RawAVDate = Mid(AVDatVersion, DateStart+1) ' This grabs everything to the right of the "(" character
  FormattedPatternAge = Left(RawAVDate,4) & "/" & Mid(RawAVDate,5,2) & "/" & Mid(RawAVDate,7,2)
  output.writeline "- Last A/V Definition update was: " & FormattedPatternAge
  
  ' Now let's grab the '
  FormattedAVVersion = "Version: " & Left(AVDatVersion,DateStart -1) & "Date: " & FormattedPatternAge
  output.writeline FormattedAVVersion 

  objReg.GetDWordValue HKEY_LOCAL_MACHINE,strESETKeyPath&"Scanners\01010100\Profiles","Enable", OnAccessScanningEnabled

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled
  
  ' Let's figure out how long it's been since ESET was updated
    CurrentDate = Now
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time ESET NOD32 was updated was " & CalculatedPatternAge & " days ago."
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are


End Sub


 
 
 
 
' *************************************  
' Sub: ObtainKESServerData
' *************************************
Sub ObtainKESServerData
   strKasperskyKESServerAVDatePath = Registry & "KasperskyLab\Components\34\1103\1.0.0.0\Statistics\AVState\"
    CurrentDate = Now

  
  
  'Let's figure out the version number
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyKESServerKeyPath,"ProdVersion",FormattedAVVersion
  output.writeline "- The version of Kaspersky running is: " & FormattedAVVersion

     
   If RegKeyExists ("HKLM\" & strKasperskyKESServerAVDatePath & "Protection_BasesDate") Then
    
    output.writeline "- The age of the Kaspersky A/V defintions is in the registry."
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath,"Protection_BasesDate",AVDatVersion
    RawAVDate = LEFT(AVDatVersion,10)
    output.writeline "- The AV Date from the registry is: " & RawAVDate
    
    'Looks like KES 10 changes things, and store dates in a format that needs transposition. Older versions of KES still do though, so we'll check the 
    'version number and decide which action to take (transpose date values or not).

	'Version 1.70 - Removed comments for the formatting (the next 5 lines except for the formattedpartternage line
	'Verstion 1.83 - made <=10 instead of <10 to account for new setting in 10.2.1.23
    If CInt(Left(FormattedAVVersion,InStr(FormattedAVVersion,".")-1)) <= 10 Then
'      FormattedPatternAge = Mid(RawAVDate,4,2) & "/" & Left(RawAVDate,2) & "/" & Mid(RawAVDate,7,4)
	   output.WriteLine Mid(RawAVDate,7,4) & "/" & Mid(RawAVDate,4,2) & "/" & Left(RawAVDate,2) 
	   output.WriteLine RawAVDate
       FormattedPatternAge = DateValue(Mid(RawAVDate,7,4) & "/" & Mid(RawAVDate,4,2) & "/" & Left(RawAVDate,2) )
       output.writeline "- Transposing the Month and the Day, we get: " & FormattedPatternAge
    Else
                    FormattedPatternAge = RawAVDate
    End If
    output.writeline "- The current date is: " & CurrentDate 
    output.writeline "- According to Kaspersky, it was last updated on: " & FormattedPatternAge
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
    output.writeline "- The last time Kaspersky was updated was " & CalculatedPatternAge & " days ago."
   Else
   
    output.writeline "- The age of the Kaspersky A/V definitions wasn't found in the registry."
    'Grab the Kaspersky data from the text file from the A/V install path
    output.writeline Path

    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    Set objFile = objFSO.OpenTextFile(Path, 1) 
    Set f = objFSO.GetFile(Path)

	Dim s_MonthName
    Dim arrFileLines() 
    i = 0 
    Do Until objFile.AtEndOfStream 
      Redim Preserve arrFileLines(i) 
      arrFileLines(i) = objFile.ReadLine
      If InStr(arrFileLines(i),"UpdateDate") Then 
        RawAVDate = Mid(arrFileLines(i),14,8)

        FormattedPatternAge = DateValue(Mid(RawAVDate,5,4) & "/" & Mid(RawAVDate,3,2) & "/" & Left(RawAVDate,2) )
        output.writeline "- The current date is: " & CurrentDate
        output.writeline "- Transposing the Month and the Day, we get: " & FormattedPatternAge
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
    output.writeline "- The last time Kaspersky was updated was " & CalculatedPatternAge & " days ago."
      End If
      i = i + 7 
    Loop 
    objFile.Close 
    
   End If





  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

   objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath,"Protection_AvRunning",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
  Else
      OnAccessScanningEnabled = FALSE
  End If
                                OnAccessScanningEnabled = TRUE
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled



End Sub



' *************************************  
' Sub: ObtainWebrootAnywhereAVData
' *************************************
Sub ObtainWebrootAnywhereAVData
	objReg.GetStringValue HKEY_LOCAL_MACHINE, strWebRootStatusPath, "Version", FormattedAVVersion
	
	AVDatVersion = "Online Database"
	
	Dim ProtectionEnabled
	Dim ServiceActive
	Dim CalculatedUpdateAge
	
	objReg.GetDWordValue HKEY_LOCAL_MACHINE, strWebRootStatusPath, "ProtectionEnabled" , ProtectionEnabled
	ServiceActive = isProcessRunning(strComputer,"WRSA.exe")
	
	if ( ProtectionEnabled = 1 AND ServiceActive) then
		OnAccessScanningEnabled = True
	Else
		OnAccessScanningEnabled = False
	End If
	
	objReg.GetDWordValue HKEY_LOCAL_MACHINE, strWebRootStatusPath, "UpdateTime" , ProductUpToDate
	
	CalculatedPatternAge = DateAdd("s", ProductUpToDate, "01/01/1970 00:00:00")
    CalculatedPatternAge = DateDiff ("d", CDate(CalculatedPatternAge), Date)
    
	CalculateAVAge 'Call the function to determine how old the A/V Definitions are
	
   output.writeline "- Product version: " & InstalledAV
   output.writeline "- Version: " & FormattedAVVersion
   output.writeline "- Is Real-Time Scanning Enabled? " & OnAccessScanningEnabled
   output.writeline "- Is the Product up-to-date? " & ProductUpToDate  

End Sub


' *************************************  
' Sub: ObtainKasperskySOSData
' *************************************
Sub ObtainKasperskySOSata
   strKasperskyKESServerAVDatePath = Registry & "KasperskyLab\protected\AVP9\"
   objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath & "\Data\","LastSuccessfulUpdate",AVDatVersion
  




  ' The AVDatVersion value is a UNIX timestamp - so this script needs to convert it to a human-readable format before continuing.
    CurrentDate = Now
  output.writeline "The current date is: " & CurrentDate
  FormattedPatternAge = DateAdd("s", AVDatVersion, #1/1/1970#)
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time Kaspersky was updated was on " & FormattedPatternAge & ", which was " & CalculatedPatternAge & " days ago."




  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

   objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath,"Enabled",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
  Else
      OnAccessScanningEnabled = FALSE
  End If
                                OnAccessScanningEnabled = TRUE
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  
  
  'Let's figure out the version number
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath & "environment\","ProductVersion",FormattedAVVersion
  output.writeline "- The version of Kaspersky running is: " & FormattedAVVersion
  InstalledAV = InstalledAV & " " & FormattedAVVersion

End Sub





' *************************************  
' Sub: ObtainTotalDefenseAVData
' *************************************
Sub ObtainTotalDefenseAVData
   objReg.GetStringValue HKEY_LOCAL_MACHINE,strTotalDefenseKeyPath,"AMSigsVersion",FormattedAVVersion
   objReg.GetStringValue HKEY_LOCAL_MACHINE,strTotalDefenseKeyPath,"Version",Version
   
   
   InstalledAV = InstalledAV & " " & Version
   output.writeline "- " & InstalledAV & " has been found on this device."
   output.writeline "- Total Defense is running the following A/V Definition Version: " & FormattedAVVersion 
   
   
   ' Let's check to see if the service is running. If it is, then Real-Time Scanning is enabled
   strWMIQuery = "Select * from Win32_Service Where Name = 'UMXEngine' and state='Running'"
   Set objWMIService = GetObject("winmgmts:" _
                                & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

                If objWMIService.ExecQuery(strWMIQuery).Count > 0 then
    OnAccessScanningEnabled = TRUE
  Else
    OnAccessScanningEnabled = FALSE
                End If
                
                output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  
  ' We're going to look at the 'Last Modified' date of a Total Defense file to determine if the A/V Definitions are up-to-date or not.
  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  If objFSO.FolderExists("C:\Documents and Settings\All Users\Application Data\CA\TDClient\Ccube") Then
                  Path = "C:\Documents and Settings\All Users\Application Data\CA\TDClient\Ccube\ccUpdateLog.txt"
  Else
                  Path = "C:\Program Data\CA\TDClient\Ccube\ccUpdateLog.txt"
  End If
  Set objFile = objFSO.GetFile(Path) 
  AVDatVersion = objFile.DateLastModified
  output.writeline "The file was last modified on: " & AVDatVersion
    CurrentDate = Now
  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time Total Defense r12 was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub



' *************************************  
' Sub: ObtainAviraAVData
' *************************************
Sub ObtainAviraAVData
  ' Let's figure out what version of Avira has been installed.
  Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2\Applications\Avira_AntiVir")
  Set colItems = objWMIService.ExecQuery("SELECT * from Product_Info", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

  For each objItem in colItems 
    InstalledAV = objItem.Product_Name & " " & objItem.Product_Version
    output.writeline "- " & InstalledAV & " has been found on this device."
    FormattedAVVersion = objItem.VDF_Version
    output.writeline "- Avira is running version " & FormattedAVVersion & " of the AV definition file."
    AVDatVersion = Left(objItem.Last_Update_Date,8)                
  Next
  
  colItems = NULL
  objItem = NULL
  
  
  ' Let's figure out whether or not Real-Time Scanning is enabled.
  Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2\Applications\Avira_AntiVir")
  Set colItems = objWMIService.ExecQuery("SELECT * from Status_Info", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

  For each objItem in colItems 
                 If objItem.Guard_Status = 2 then
    OnAccessScanningEnabled = TRUE
   Else
    OnAccessScanningEnabled = FALSE
                End If
   output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled              
  Next  
 

  ' Let's figure out when Avira's AV Definitions were last updated.
    CurrentDate = Now
  output.writeline "- The current date is: " & CurrentDate
  RawAVDate = LEFT(AVDatVersion,8)
  output.writeline "- According to WMI, The AV definition file was last updated on: " & RawAVDate
  FormattedPatternAge = Left(RawAVDate,4) & "/" & Mid(RawAVDate,5,2) & "/" & Right(RawAVDate,2)
  output.writeline "- Prettying up what was in WMI, we get: " & FormattedPatternAge
  

  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
 output.writeline "- The last time Avira was updated was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub


' *************************************  
' Sub: ObtainFSecureAVData
' *************************************
Sub ObtainFSecureAVData
   objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPathLoc,"CurrentVersion",Version
   If RegKeyExists("HKLM\" & strFSecureRegPathLoc & "InstallationDirectory") Then
                                objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPathLoc,"InstallationDirectory",strFSecureInstallPath
   else
                                objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPathLoc,"Path",strFSecureInstallPath
   end if
   
   
  ' We're going to look at the 'Last Modified' date of a Total Defense file to determine if the A/V Definitions are up-to-date or not.
  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  If objFSO.FolderExists(strFSecureInstallPath) Then
                                if objFSO.FileExists(strFSecureInstallPath & "\FS@aqua.ini") then
                                                Path = strFSecureInstallPath & "\FS@aqua.ini"
                                                output.writeline "Using File : AQUA"
                                elseif objFSO.FileExists(strFSecureInstallPath & "\FS@hydra.ini") then
                                                Path = strFSecureInstallPath & "\FS@hydra.ini"
                                                output.writeline "Using File : HYDRA"
                                else
                                                Path = strFSecureInstallPath & "\FS@orion.ini"
                                                output.writeline "Using File : ORION"
                                end if


                                ' TWO METHODS POSSIBLE - I CAN LOOK THROUGH THE FILE TO SEE THE UPDATE DATE, OR LOOK AT THE FILE VERSION DATE. CURRENTLY SET TO UPDATE DATE

                                'METHOD 1 
                                Set objFile = objFSO.OpenTextFile(Path) 
                                Do Until objFile.AtEndOfStream 
                                strLine = objFile.ReadLine 
                                If InStr(strLine, "File_set_visible_version=") Then 
                                                output.writeline left(right(strLine,13),10) 
                                                AVDatVersion= left(right(strLine,13),10)
                                End if 
                                If InStr(strLine, "the requested operation failed") Then 
                                output.writeline "Failed" 
                                End if 
                                Loop 

                                objFile.Close 

                                'METHOD 2 - DISABLED
                                
                                ' Set objFile = objFSO.GetFile(Path) 
                                 ' AVDatVersion = objFile.DateLastModified
                                ' output.writeline "The file was last modified on: " & AVDatVersion

                                
                                
  else
                                AVDatVersion="01/01/2001"
  End If
    'FormattedAVVersion
    ' InstalledAV
     'OnAccessScanningEnabled
     'ProductUpToDate
  

if isProcessRunning(".","fsav32.exe") then
                OnAccessScanningEnabled          =TRUE
else
                OnAccessScanningEnabled          =FALSE
end if
  

   InstalledAV = InstalledAV  & " - " & Version
   FormattedAVVersion =  Version '& " - " & AVDatVersion
   output.writeline "- " & InstalledAV & " has been found on this device."
   output.writeline "- F-Secure is running the following A/V Definition Version: " & FormattedAVVersion 

   CurrentDate = Now
  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time F-Secure was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  output.writeline "product up to tdate ? " & ProductUpToDate

End Sub



'ObtainSEPCloudData
' *************************************  
' Sub: ObtainSEPCloudData
' *************************************
Sub ObtainSEPCloudData
   'objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPathLoc,"CurrentVersion",Version
   strSEPCloudRegpath3 = strSEPCloudRegpath2 & "\SharedDefs\"

   If RegKeyExists("HKLM\" & strSEPCloudRegpath3 & "AVDEFMGR") Then
   
                                objReg.GetStringValue HKEY_LOCAL_MACHINE,strSEPCloudRegpath3 ,"AVDEFMGR",strSEPCloudDefPath
                                output.writeline "- AV Definition date in file : " & strSEPCloudDefPath
   else
                                output.writeline "- INVALID PATH : HKLM\" & strSEPCloudRegpath3
   end if
                
                
                If RegKeyExists("HKLM\" & strSEPCloudRegpath2 & "\PRODUCTVERSION") Then
                                
                                objReg.GetStringValue HKEY_LOCAL_MACHINE,strSEPCloudRegpath2 ,"PRODUCTVERSION",Version
                                output.writeline "- AV Version is : " & Version
   else
                                output.writeline "- INVALID Version Path : HKLM\" & strSEPCloudRegpath2 & "\PRODUCTVERSION"
   end If
   
   
   
   
  ' We're going to look at the 'Last Modified' date of a Total Defense file to determine if the A/V Definitions are up-to-date or not.
  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  If objFSO.FolderExists(strSEPCloudDefPath) Then

                                'METHOD 1 
                                
'                               Set objFile = objFSO.OpenTextFile(strSEPCloudDefPath & "\virscan.inf") 
'                               Do Until objFile.AtEndOfStream 
'                               strLine = objFile.ReadLine 
'                               If InStr(strLine, "CurDefs=") Then 
'                                               output.writeline left(right(strLine,10),13) 
'                                               AVDatVersion= left(right(strLine,10),13)
'                               End if 
'                               If InStr(strLine, "the requested operation failed") Then 
'                               output.writeline "Failed" 
'                               End if 
'                               Loop 
'
'                               objFile.Close 

                                
                                
                                  Set objFile = objFSO.GetFile(strSEPCloudDefPath & "\versioninfo.dat") 
                                  AVDatVersion = objFile.DateLastModified
                                  output.writeline "The file was last modified on: " & AVDatVersion

                                
                                
  else
                                AVDatVersion="01/01/2001"
  End If
  

if isProcessRunning(".","AVAgent.exe") then
                OnAccessScanningEnabled          =TRUE
else
                OnAccessScanningEnabled          =FALSE
end if
  

   InstalledAV = InstalledAV  & " - " & Version
   FormattedAVVersion =  Version '& " - " & AVDatVersion
   output.writeline "- " & InstalledAV & " has been found on this device."
   output.writeline "- Symantec Endpoint.Cloud is running the following A/V Definition Version: " & FormattedAVVersion 

   CurrentDate = Now
  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time SEP Cloud was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  output.writeline "product up to tdate ? " & ProductUpToDate

End Sub




'ObtainPandaCloudOfficeData
' *************************************  
' Sub: ObtainPandaCloudOfficeData
' *************************************
Sub ObtainPandaCloudOfficeData

  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  strPandaAVDefinitionPath = "C:\Program Files\Panda Security\WaAgent\WalUpd\Data\Catalog"

  If objFSO.FolderExists(strPandaAVDefinitionPath) Then
	  Set objFile = objFSO.GetFile( strPandaAVDefinitionPath & "\Last_Catalog") 
	  AVDatVersion = objFile.DateLastModified
	  output.writeline "The file was last modified on: " & AVDatVersion
  else
		AVDatVersion="01/01/2001"

	    strPandaAVDefinitionPath = "C:\Program Files (x86)\Panda Security\WaAgent\WalUpd\Data\Catalog"
	  
	    If objFSO.FolderExists(strPandaAVDefinitionPath) Then
	  	  Set objFile = objFSO.GetFile( strPandaAVDefinitionPath & "\Last_Catalog") 
	  	  AVDatVersion = objFile.DateLastModified
	  	  output.writeline "The file was last modified on: " & AVDatVersion
	    else
	  		AVDatVersion="01/01/2001"
	    End If  
  End If  

	if isProcessRunning(".","WAHost.exe") then
		OnAccessScanningEnabled	=TRUE
	else
		OnAccessScanningEnabled	=FALSE
	end if

	'InstalledAV = InstalledAV  & " - " & Version
	
  	If RegKeyExists ( "HKLM\" & strPandaCloudEPPath32 & "\Normal") Then
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strPandaCloudEPPath32 ,"Normal",RawAVVersion 
  	ElseIf RegKeyExists ( "HKLM\" & strPandaCloudEPPath64 & "\Normal") Then
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strPandaCloudEPPath64 ,"Normal",RawAVVersion 
    Else
    	RawAVVersion = "UNKNOWNVERSION"
	End If
	FormattedAVVersion =  InstalledAV & " - " & RawAVVersion
	output.writeline "- " & InstalledAV & " has been found on this device."
	output.writeline "- Panda Cloud Endpoint Protection is running the following A/V Definition Version: " & FormattedAVVersion 

  CurrentDate = Now
  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time Panda Cloud Endpoint Protection was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  output.writeline "product up to tdate ? " & ProductUpToDate

end sub



'ObtainAVG2014Data
' *************************************  
' Sub: ObtainAVG2014Data
' *************************************
Sub ObtainAVG2014Data

  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  If objFSO.FolderExists(stravg2014defpath) Then
	  Set objFile = objFSO.GetFile( stravg2014defpath & "\incavi.avm") 
	  AVDatVersion = objFile.DateLastModified
	  output.writeline "The file was last modified on: " & AVDatVersion
  else
		AVDatVersion="01/01/2001"
	    output.writeline "the AV Definition file is not found"
  End If  

	if isProcessRunning(".","avgwdsvc.exe") Then
		OnAccessScanningEnabled	=TRUE
	else
		OnAccessScanningEnabled	=FALSE
	end if
	
	  If RegKeyExists ( "HKLM\" & stravg2014regpath32 & "\ProdType") Then
		   objReg.GetStringValue HKEY_LOCAL_MACHINE,stravg2014regpath32,"ProdType",RawAVVersion
 '	      output.writeline "- " & InstalledAV & " " & RawAVVersion			& " has been detected." 
      ElseIf RegKeyExists ( "HKLM\" & stravg2014regpath64 & "\ProdType") Then
		   objReg.GetStringValue HKEY_LOCAL_MACHINE,stravg2014regpath64,"ProdType",RawAVVersion
' 	      output.writeline "- " & InstalledAV & " " & RawAVVersion			& " has been detected." 
      End If

	FormattedAVVersion =  InstalledAV & " - " & RawAVVersion
	output.writeline "- " & FormattedAVVersion & " has been found on this device."

  CurrentDate = Now
  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time AVG 2014 was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  output.writeline "product up to tdate ? " & ProductUpToDate

end sub






Sub ObtainAvastData
	output.writeline "CHECKING AVAST"
	
  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  If objFSO.FolderExists(strAvastInstallPath & "\Defs") Then
	  Set objFile = objFSO.GetFile( StrAvastInstallPath & "\defs\aswdefs.ini") 
	  AVDatVersion = objFile.DateLastModified
	  output.writeline "The file was last modified on: " & AVDatVersion
  else
		AVDatVersion="01/01/2001"
	    output.writeline "the AV Definition file is not found"
  End If  
  
	if isProcessRunning(".","avastsvc.exe") Then
		OnAccessScanningEnabled	=TRUE
	else
		OnAccessScanningEnabled	=FALSE
	end if

	  If RegKeyExists ( "HKLM\" & strAvastRegPath32 & "\Version"	  ) Then
		   objReg.GetStringValue HKEY_LOCAL_MACHINE,strAvastRegPath32,"Version",RawAVVersion
 	      output.writeline "- " & InstalledAV & " " & RawAVVersion			& " has been detected." 
	  ElseIf RegKeyExists ( "HKLM\" & strAvastRegPath64 & "\Version") Then
		   objReg.GetStringValue HKEY_LOCAL_MACHINE,strAvastRegPath64,"Version",RawAVVersion
	      output.writeline "- " & InstalledAV & " " & RawAVVersion			& " has been detected." 
      End If



	
  CurrentDate = Now
  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time Avast! was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  output.writeline "product up to tdate ? " & ProductUpToDate
  	
  
  
End Sub








function isProcessRunning(byval strComputer,byval strProcessName)

                Dim objWMIService, strWMIQuery

                strWMIQuery = "Select * from Win32_Process where name like '" & strProcessName & "'"
                
                Set objWMIService = GetObject("winmgmts:" _
                                & "{impersonationLevel=impersonate}!\\" _ 
                                                & strComputer & "\root\cimv2") 


                if objWMIService.ExecQuery(strWMIQuery).Count > 0 then
                                isProcessRunning = true
                else
                                isProcessRunning = false
                end if

end function



' *************************************  
' Sub: ObtainSecurityCenter2Data
' *************************************
Sub ObtainSecurityCenter2Data

   ' Let's check to see what A/V product is installed, according to the Windows Security Center.
   Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\SecurityCenter2")
   Set colItems = objWMIService.ExecQuery("SELECT * from AntiVirusProduct", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

   output.writeline "- Unable to determine installed AV."
      InstalledAV = "No AV installed"
      onAccessScanningEnabled = "FALSE"
      ProductUpToDate = "FALSE"
      FormattedAVVersion = "Unknown"

   For each objItem in colItems
     InstalledAV = objItem.displayName
     FormattedAVVersion = "Unable to detect the version of A/V Definitions being used - this information is not available through the Windows Security Center."
     output.writeline "- " & InstalledAV & " has been detected."
     
     ' The 'ProductState' value needs to be converted into HEX, and the parsed into 3 different sub-values, according to the following article: 
     ' http://neophob.com/2010/03/wmi-query-windows-securitycenter2/
     ' Keep this line around for debugging purposes. output.writeline "- According to WMI, the state of " & InstalledAV & " is " & objItem.ProductState
     HexProductState = Hex(objItem.ProductState)
     ' Keep this line around for debugging purposes. output.writeline "- Converting that value to HEX, we get " & HexProductState
     
     ' The middle of the 3 HEX values equates to the scanner state
     HexScannerState =  Mid(HexProductState,2,2)
     ' Keep this line around for debugging purposes. output.writeline CLng("&H" & HexScannerState)
     If CLng("&H" & HexScannerState) = 16 then
      OnAccessScanningEnabled = TRUE
     Else
      OnAccessScanningEnabled = FALSE
	  
                   End If
    output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled  
    
     ' The last of the 3 HEX values tells us if the A/V definition files are up-to-date or not
     HexAVDefState =  Right(HexProductState,2)
     ' Keep this line around for debugging purposes. output.writeline CLng("&H" & HexAVDefState)
     If CLng("&H" & HexAVDefState) = 0 then
      ProductUpToDate = TRUE
     Else
      ProductUpToDate = FALSE
                   End If
    output.writeline "- Are the AV Definitions up-to-date? " & OnAccessScanningEnabled 
    
                    
  Next

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub




' *****************************  
' Sub: OSVersion
' *****************************
Sub OSVersion
                       
  ' 1. Determine if this is a 32-bit machine or a 64-bit machine (as this will determine what registry values we modify)
Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

For each objItem in colItems 
                OSName = objItem.Caption
                output.writeline "- This machine is running " & OSName
Next

End Sub

Function HexToDec(strHex)
    Dim i
    Dim size
    Dim ret
    size = Len(strHex) - 1
    ret = CDbl(0)
    For i = 0 To size
        ret = ret + CDbl("&H" & Mid(strHex, size - i + 1, 1)) * (CDbl(16) ^ CDbl(i))
    Next
    HexToDec = ret
End Function
Function DecToHex(dblNumber)
    Dim Q
    Dim ret

    ret = ""
    Q = CDbl(Fix(dblNumber))
    While Q > 0
        ret = Hex(Q - Fix(Q / 16) * 16) & ret
        Q = Fix(Q / CDbl(16))
    Wend
    DecToHex = ret
End Function

' *************************************  
' Sub: ObtainWindowsDefenderData
' *************************************
Sub ObtainWindowsDefenderData
    CurrentDate = Now

  
  
  'Let's figure out the version number
	If RegKeyExists ("HKLM\" & strWindowsDefenderPath & "Signature Updates\SignaturesLastUpdated") Then
		objReg.GetBinaryValue HKEY_LOCAL_MACHINE,strWindowsDefenderPath & "Signature Updates\","SignaturesLastUpdated",ReturnedBinaryArray1
	ElseIf RegKeyExists ("HKLM\" & strWindowsDefenderPath & "Signature Updates\ASSignatureApplied") Then
		objReg.GetBinaryValue HKEY_LOCAL_MACHINE,strWindowsDefenderPath & "Signature Updates\","ASSignatureApplied",ReturnedBinaryArray1
	Else
		output.WriteLine "DATE NOT FOUND"
	End If

	  objReg.GetStringValue HKEY_LOCAL_MACHINE,strWindowsDefenderPath & "Signature Updates\","ASSignatureVersion",FormattedAVVersion
	  output.writeline "- The version of Windows Defender running on this machine is: " & FormattedAVVersion

   RawAVDate=""

   'For i = 0 To UBound(ReturnedBinaryArray1)
'		if(clng(ReturnedBinaryArray1(i))) < 10 then
'			RawAVDate= RawAVDate & "0" & hex(ReturnedBinaryArray1(i))
'		else
'			RawAVDate= RawAVDate & hex(ReturnedBinaryArray1(i))'
'		end if
'  Next
'  output.writeline "Hex Value: " & RawAVDate
'  output.writeline "Decimal Value: " & VarType(Abs(cdbl("&h" & RawAVDate)))
'  output.writeline "Update Date: " & DateAdd("s",(Abs(cdbl("&h" & RawAVDate)))/10,"01-Jan-1601")
  'output.writeline dConvertWMItoVBSDate()
  ' Test function #2
  'output.writeline DateAdd("s",HexToDec(RawAVDate),"01-Jan-1601")
lngBias=0    'determine it per your link on your own

lngHigh=0
lngLow=0
for i=7 to 4 step -1
    lngHigh=lngHigh*256+ReturnedBinaryArray1(i)
next
for i=3 to 0 step -1
    lngLow=lngLow*256+ReturnedBinaryArray1(i)
next

if err.number<>0 then
    dtmDate = #1/1/1601#
    err.clear
else
    If lngLow < 0 Then
        lngHigh = lngHigh + 1
    End If
    If (lngHigh = 0) And (lngLow = 0 ) Then
        dtmDate = #1/1/1601#
    Else
        dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
            + lngLow)/600000000 - lngBias)/1440
    End If
End If
on error goto 0
'this shows you what you get
wscript.echo dtmDate

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  
  ' If the DisableRealTimeMonitoring registry key exists, then Real-Time Monitoring is disabled. If that key does not exist, then Real-Time Monitoring is enabled.
  If RegKeyExists ("HKLM\" & strWindowsDefenderPath & "Real-Time Protection\DisableRealtimeMonitoring") Then
    ' If real-Time Scanning has been disabled in the past, the DisableRealtimeMonitoring key will still be around, even if Real-Time Scanning has been re-enabled. The value will be set to zero though, in that case.
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strWindowsDefenderPath & "Real-Time Protection\","DisableRealtimeMonitoring",DisableRealtimeMonitoring
    If DisableRealtimeMonitoring = 1 Then
      OnAccessScanningEnabled = FALSE
    Else
      OnAccessScanningEnabled = TRUE
    End If    
  Else
    OnAccessScanningEnabled = TRUE
  End If    
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled



End Sub

Private Function dConvertWMItoVBSDate(sDate)
  Dim sMonth, sDay, sYear, sHour, sMinutes, sSeconds
  sMonth = Mid(sDate,5,2)
  sDay = Mid(sDate,7,2)
  sYear = Mid(sDate,1,4)
  sHour = Mid(sDate,9,2)
  sMinutes = Mid(sDate,11,2)
  sSeconds = Mid(sDate,13,2)
  dConvertWMItoVBSDate = DateSerial (sYear, sMonth, sDay) + TimeSerial (sHour, sMinutes, sSeconds)
End Function


Function HexToDec(strHex)
    Dim i
    Dim size
    Dim ret
    size = Len(strHex) - 1
    ret = CDbl(0)
    For i = 0 To size
        ret = ret + CDbl("&H" & Mid(strHex, size - i + 1, 1)) * (CDbl(16) ^ CDbl(i))
    Next
    HexToDec = ret
End Function

