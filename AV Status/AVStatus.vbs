  ' *************************************************************************************************************************************************
' Script: AVStatus.vbs
' Version: 2.37
' Maintained by : Chris Reid @SolarWinds MSP
' Description: This script checks the status of the A/V software installed on 
'              the machine, and writes data about the A/V software to the 
'              AntiVirusProduct WMI Class in the root\SecurityCenter WMI namespace.
' Date: Nov 6th, 2020
' Compatibility : It is tested on the current versions of Windows but it also should work on all desktop and server versions, from Windows XP to Windows Server 2019.
' Usage in N-Central : avstatus.vbs WRITE     OR     avstatus.vbs DONOTWRITE   (the second option will write no data into WMI if an A/V product cannot be found)
' Usage in Windows Command Prompt : CSCRIPT avstatus.vbs WRITE     OR     CSCRIPT avstatus.vbs DONOTWRITE   (the second option will write no data into WMI if an A/V product cannot be found)
' *************************************************************************************************************************************************

                                             

' Define the  variables used in the script
Option Explicit

dim HKEY_LOCAL_MACHINE, strComputer, objReg, strTrendKeyPath, InputRegistryKey1, InputRegistryKey2
dim oReg,arrSubKeys,SubKey
dim InputRegistryKey3, InstalledAV, NamespacePresense, objClassCreator, objGetClass, objWMIObject
dim RegstrValue, ReturneddwValue1, objWMIService, objItem, objItem2, objNameSpace, ReturneddwValue2
dim wbemCimtypeString, wbemCimtypeUint32, colNamespaces, objNewNamespace
dim RawAVVersion, FormattedAVVersion, objNewInstance, strWMINamespace, ParentWMINamespace
dim ReturnedstrValue1, strValue, FormattedPatternAge, CalculatedPatternAge, CurrentDate
dim WMINamespace, strWMIClassWithQuotes, strWMIClassNoQuotes, colClasses, objClass 
dim strTestKeyPath, Registry, Registry32, Registry64, arrValueNames, strSymantecESKeyPath, WshShell
dim ReturnedBinaryArray1(7), bytevalue, i, f, SymantecAvPatternDate
dim SymantecAVMonth, SymantecAVYear, SymantecAVDate, ProductUpToDate, wbemCimtypeBoolean, OnAccessScanningEnabled, ReturneddwValue3, strTrendRealTimeKeyPath, InputRegistryKey4, InputRegistryKey7, strTrendRealTimeKeyPathNew
dim AddressWidth, colItems, colItems2, RawOnAccessScanningEnabled, Revision, strSymantecAVKeyPath, strSophosAVKeyPath, strSophosAVVersionPath, strSophosUpdateStatusPath
dim RawAVDate, strSymantecAVRealTimePath, strTrendPatternAgeKeyPath, strTrend7KeyPath, objComponentMgr, objConfigMgr
Dim strTrendVersionKeyPath, VIPREAvPatternDate, VIPREAVMonth, VIPREAVYear, VIPREAVDate, strVIPREAVRealTimePath, strVIPREESFolderPath
Dim McAfeeDatVersion, strMcAfeeVersionPath, McAfeeDatOAS, strMcAfeePath, InputRegistryKey5, strMcAfeeOASPath, OASEnabled, RAWOASEnabled, mcAfeeEndPointSecurityVersion, mcAfeeEndpointSecurityDefDatePath, mcAfeeEndpointSecurityBuildNum, mcAfeeEndpointSecurityVerNum, mcAfeeEndpointSecurity, mcAfeeEndpointSecurityDefDate
Dim strVIPREAVKeyPath, objFSO, ProgramFiles, Path, objFile, strLine, objNode, output, ProductVersion, AVDatDate, AVDatVersion, OutOfDateDays
Dim strVIPREEnterpriseKeyPath, InstallLocation, strTrendProductVersionKeyPath, strTrendAVVersionKeyPath, ReturneddwValue4, strKasperskyAV2012KeyPath, ReturneddwValue5
Dim strKaspersky2012AVDatePath, strKasperskyAV2012Path, strKasperskyAV60KeyPath, strKasperskyAV60DatePath, strKasperskyAV60Path, strKasperskyKES8KeyPath, strKasperskyKES8AVDatePath
Dim strSecurityEssentialsKeyPath, RawAVDefDate, strEndpointSecurityKeyPath   
Dim InputRegistryKey6, ProductName, strVIPREBusiness5KeyPath, strVIPREAV2012KeyPath, strESETKeyPath, DateStart
Dim strKasperskyKESServerKeyPath, strKasperskyKESServerAVDatePath, InstalledApp, ProductVersionKey, strForefrontKeyPath, strKasperskyKES8ServerKeyPath, strKasperskyKES6ServerKeyPath 
Dim Version, WSHStdOut, filename, cscriptExec, strKasperskySOS2KeyPath, strKasperskySOS3KeyPath, Return, strTotalDefenseKeyPath, strWMIQuery
Dim strAviraKeyPath, NoAVBehavior, strWMINamespace2, HexProductState, HexScannerState, HexAVDefState, strAviraServerKeyPath, DeleteWMINamespace
Dim strFSecureRegPath0, strFSecureRegPath00, strFSecureRegPath000, strFSecureRegPath1, strFSecureRegPath2, strFSecureRegPath3, strFSecureRegPath4, strFSecureRegPath5, strFSecureRegPath6, strFSecureRegPathLoc,strFSecureInstallPath
Dim strSEPCloudRegPath0, strSEPCloudRegpath2,strSEPCloudRegpath3,strSEPCloudRegpath4,strSEPCloudDefPath, strSEPCloudDefPath2, OSName, strKES10KeyPath, strKES10KeyPathSP1, strKES10KeyPathSP2
Dim strPandaCloudEPPath64, strPandaCloudEPPath32, strPandaAVDefinitionPath
Dim strWindowsDefenderPath, DisableRealtimeMonitoring, strTrendProductVersion
Dim stravg2014regpath32, stravg2014regpath64, stravg2014defpath
Dim stravg2013regpath32, stravg2013regpath64, stravg2013defpath
dim lngBias, dtmDate,lngHigh,lngLow
Dim strWebRootStatusPath, strWebRootStatusPath32
Dim strTrendVerLen, InstalledAV1, serviceactive
Dim strAvastRegPath32, strAvastInstallPath, strAvastRegPath64
Dim strViprebusinessAgt, strViprebusiness64Agt , strVipreBusinessAgtLoc, strViprebusinessAgt1, strVIPREBusinessOnlineKeyPath
Dim strMalwareBytesRegPath64, SCEPInstalled, FoundGUID, StatusCode, StatusText
Dim sMonth, sDay, sYear, sHour, sMinutes, sSeconds, strTMMSARegPath, recentFile, NamespacetoCheck, strTMDSARegPath, fileSystem, folder, file, newestfile, ProgramFiles64, stravg2016defpath, stravg2016regpath, colServices, objService
Dim strNormanregpath32, strNormanregpath64, strNormanrootpath, boolNormanversion9, strNormandefpath, strKasperskyStandAlonePath, LastUpdateDate, AVGBusSecDataFolder, arrIniFileLines, ProviderRealTimeScanningEnabled, UserRealTimeScanningDisabled
Dim objFileToRead, objFileToWrite, node, UpToDateState, strFortiClientPath, FortiClientInstallPath, objApp, strKasperskyKESServerAVVersionPath, strSophosVirtualAVKeyPath, RawProtectionStatus, strPandaAdaptiveDefencePath64, strPandaAdaptiveDefencePath32
Dim ProgramData, objFolder, objSubFolders, objSubFolder, oShell, oExec, sLine, sExecPath, sNewestFolder, dPrevDate, SCEPUninstallString, objXMLHTTP, objADOStream, S1HelperObj, S1AgentStatus, IndexOfAgentVersion, LenOfVersion


' Specify values for some of the variables

Version = "2.37"

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
InstallLocation = WshShell.ExpandEnvironmentStrings("%AllUsersProfile%")
CurrentDate=Now
Set objXMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1") 




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
  ' output.writeline "- The command-line parameter (either WRITE or DONOTWRITE) for choosing whether or not to write data to WMI if an A/V product isn't found was not specified. This script will write data to WMI regardless of whether or not an A/V product is discovered."                
End If







	
'This is a meat of the script - where all of the functions are called.

  output.writeline "- This is version " & Version & " of the script." 
  
  OSType 'This function determines whether this is a 32-bit or 64-bit OS
  OSVersion 'This function figures out what OS the machine is running
  
  DetectInstalledAV   'This function will detect what AV software is installed
  
  output.writeline "- " & InstalledAV & " has been detected."
    
  If (InstalledAV="Trend Micro Apex One" OR InstalledAV="Trend Micro Worry-Free Business Security 6" OR InstalledAV="WFBSS" OR InstalledAV="Trend Micro WFBSS" OR InstalledAV="Trend Micro Worry-Free Business Security" OR InstalledAV="Trend Micro OfficeScan" OR InstalledAV="Trend Micro Worry-Free Business Security Services" OR InstalledAV="Trend Micro WFBSH_Agent") Then
    ObtainTrendMicroData 'Call the function we created to grab info about Trend Micro from the registry
	    
  ElseIf InstalledAV="Trend Micro Worry-Free Business Security 7" Then
    ObtainTrend7AVData 'Call the function we created to grab info about Trend WFBS 7 from the registry
  
  ElseIf InstalledAV="Symantec Endpoint Protection" Then
   ObtainSymantecESData 'Call the function we created to grab info about Symantec Endpoint Security from the registry
  
  ElseIf InstalledAV="Symantec AntiVirus" Then
    ObtainSymantecAVData 'Call the function we created to grab info about Symantec AntiVirus from the registry

  
  ElseIf InstalledAV="Sophos Anti-Virus" Then
    ObtainSophos10AVData 'Call the function we created to grab info about Sophos Anti-Virus from the registry

  ElseIf InstalledAV="Sophos Endpoint Protection" Then
    ObtainSophos10AVData 'Call the function we created to grab info about Sophos Endpoint Protection from the registry
    
       
  ElseIf InstalledAV="Sophos Anti-Virus 10" Then
    ' If the script is launched on a 64-bit machine, let's re-launch it in the 32-bit command prompt. This will allow the script to properly detect Sophos.
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
    ObtainSophos10AVData 'Call the function we created to grab info about Sophos Anti-Virus from the registry
 
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

  ElseIf InstalledAV="Microsoft System Center Endpoint Protection" OR InstalledAV="Microsoft Security Essentials" OR InstalledAV="Microsoft Forefront" OR InstalledAV="Microsoft System Center Endpoint Protection (Managed Defender)" Then
    ObtainSecurityEssentialsAVData 'Call the function we created to grab info about MS Essentials from the registry
              
  ElseIf InstalledAV="Kaspersky Endpoint Security 8" Then
    ObtainKES8Data 'Call the function we created to grab info about Kaspersky Endpoint Security 8 from the registry

    
  ElseIf InstalledAV="VIPRE Business Antivirus" Then
    ObtainVIPREEnterpriseData 'Call the function we created to grab info about Vipre Business from the registry


  ElseIf InstalledAV="VIPRE Antivirus 2012" Then
    ObtainVIPREAVData 'Call the function we created to grab info about Vipre Antivirus 2012 from the registry
    
    
  ElseIf InstalledAV="ESET NOD32 Antivirus" OR InstalledAV="ESET Endpoint Antivirus" Then
    ObtainESETAVData 'Call the function we created to grab info about ESET from the registry
    
    
  ElseIf (InstalledAV="ESET File Security" OR InstalledAV="ESET Mail Security" OR InstalledAV="ESET Endpoint Security") Then
    ObtainESETFSData 'Call the function we created to grab info about ESET File Security from the registry   
    
    
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
    
  ElseIf InstalledAV="Kaspersky Small Office Security 3" Then
    ObtainKasperskySO3Sata 'Call the function we created to grab info about Kaspersky Endpoint Security 6 from the registry    
       

  ElseIf InstalledAV="Total Defense R12 Client" Then
    ObtainTotalDefenseAVData 'Call the function we created to grab info about Total Defense from the registry
    
  ElseIf InstalledAV="Avira AntiVirus" Then
    ObtainAviraAVData 'Call the function we created to grab info about Avira from the registry
          
  ElseIf InStr(1,InstalledAV,"F-Secure")>0 Then
    ObtainFSecureAVData 'Call the function we created to grab info about F-Secure from registry and folder

  ElseIf InstalledAV="Symantec Endpoint Protection Cloud" then
    ObtainSEPCloudData
  
  ElseIf InstalledAV="Avast!" Then
	 ObtainAvastData
                
  ElseIf InstalledAV="VIPRE Business Agent" Then
	 ObtainVIPREBusinessAgentData
               
  ElseIf InstalledAV="VIPRE Business Online" Then
	 ObtainVIPREBusinessAgentData
                

  ElseIf InstalledAV="Kaspersky Endpoint Security 10" Then
    strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\10.1.0.0\" 

        Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
        oReg.EnumKey HKEY_LOCAL_MACHINE, Registry & "KasperskyLab\Components\34\Connectors\KES", arrSubKeys
        For Each subkey In arrSubKeys
            strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\" & subkey & "\"
	   	Next

		    Path = InstallLocation & "\Kaspersky Lab\KES10\Data\u0607g.xml"
    ObtainKESServerData 'Call the function we created to grab info about Kaspersky from the registry

  ElseIf InstalledAV="Kaspersky Endpoint Security 10 for Windows" Then
    strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\WSEE\10.0.0.0\" 
    Path = InstallLocation & "\Kaspersky Lab\KES10\Data\u0607g.xml"
    ObtainKESServerData 'Call the function we created to grab info about Kaspersky from the registry
    
  ElseIf InstalledAV="Kaspersky Endpoint Security 11 for Windows" Then
    strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\11.0.0.0\" 
    Path = InstallLocation & "\Kaspersky Lab\KES10\Data\u0607g.xml"
    ObtainKESServerData 'Call the function we created to grab info about Kaspersky from the registry

		
  ElseIf InstalledAV="Kaspersky Endpoint Security 10 SP1" Then
    strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\10.2.2.0\" 

        Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
        oReg.EnumKey HKEY_LOCAL_MACHINE, Registry & "KasperskyLab\Components\34\Connectors\KES", arrSubKeys
        For Each subkey In arrSubKeys
            strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\" & subkey & "\"
	   	Next

     Path = InstallLocation & "\Kaspersky Lab\KES10SP1\Data\u0607g.xml"
    ObtainKESServerData 'Call the function we created to grab info about Kaspersky from the registry
    
  ElseIf InstalledAV="Kaspersky Endpoint Security 10 SP2" Then
    strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\10.3.0.0\" 

        Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
        oReg.EnumKey HKEY_LOCAL_MACHINE, Registry & "KasperskyLab\Components\34\Connectors\KES", arrSubKeys
        For Each subkey In arrSubKeys
            strKasperskyKESServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KES\" & subkey & "\"
            'output.writeline strKasperskyKESServerKeyPath 'This is a debug line, added in to troubleshoot NCI-8949
	   	Next

     Path = InstallLocation & "\Kaspersky Lab\KES10SP2\Data\u1313g.xml"
    ObtainKESServerData 'Call the function we created to grab info about Kaspersky from the registry

  ElseIf InstalledAV="Panda Adaptive Defense 360 32 Bit" OR InstalledAV="Panda Adaptive Defense 360 64 Bit" Then
    ObtainPandaAdaptiveDefenceData 'Call the function we created to grab info about Panda Adaptive Defense from the registry


  ElseIf InstalledAV="Panda Endpoint Protection 10 32 Bit" Then
    strPandaAVDefinitionPath = "C:\Program Files\Panda Security\WaAgent\WalUpd\Data\Catalog"
    ObtainPandaCloudOfficeData 'Call the function we created to grab info about Kaspersky Endpoint Security 6 from the registry

  ElseIf InstalledAV="Panda Endpoint Protection 10 64 Bit" Then
    strPandaAVDefinitionPath = "C:\Program Files (x86)\Panda Security\WaAgent\WalUpd\Data\Catalog"
    ObtainPandaCloudOfficeData 'Call the function we created to grab info about Kaspersky Endpoint Security 6 from the registry
	
  ElseIf InstalledAV="Windows Defender" then
          ObtainWindowsDefenderData
                
  ElseIf InstalledAV="AVG 2013" Then
  				ObtainAVG2013Data
                
  ElseIf (InstalledAV="AVG 2014" OR InstalledAV="AVG Protection")  Then
  				ObtainAVG2014Data
                
  ElseIf InstalledAV="Webroot SecureAnywhere" Then
	 ObtainWebrootAnywhereAVData 'Call the function to grab info about webroot from the registry
  
  ElseIf InstalledAV="Malwarebytes' Corporate Edition" Then
	 ObtainMalwarebytesCorporate 'Call the function to grab info about webroot from the registry
  
  ElseIf InstalledAV="McAfee Endpoint Security" Then
	 ObtainMcAfeeEndpointSecurity 'Call the function to grab info about McAfee endpoint security from the registry
  
  ElseIf InstalledAV="McAfee Endpoint Security 10.1" Then
	 ObtainMcAfeeEndpointSecurity101 'Call the function to grab info about McAfee endpoint security from the registry
  
  ElseIf InstalledAV="Trend Micro Deep Security Agent" Then
	 ObtainTrendMicroDeepSecurity 'Call the function to grab info about Deep Security from the registry

  ElseIf InstalledAV="McAfee Move AV Client" Then
	 ObtainMcAfeeMove 'Call the function to grab info about Deep Security from the registry
   
  ElseIf InstalledAV="Trend Micro Messaging Security Agent" Then
	 ObtainTMMSA 'Call the function to grab info about TMMSA from the registry
   
  ElseIf InstalledAV="Norman Endpoint Protection" Then
    ObtainNormanEndpointProtection 'Call the function to grab info about Norman Endpoint Protection from the registry
    
  ElseIf InstalledAV="AVG Business Security" Then
    ObtainAVGBusinessSecurity 'Call the function to grab info about Norman Endpoint Protection from the registry

  ElseIf InstalledAV="FortiClient" Then
    ObtainFortiClient 'Call the function to grab info about FortiClient from the registry
    
  ElseIf InstalledAV="Cisco Advanced Malware Protection (AMP)" Then
    ObtainCiscoAMPData 'Call the function to grab info about Cisco AMP
    
  ElseIf InstalledAV="Sophos for Virtual Environments" Then
    ObtainSophosVirtualAVData 'Call the function to grab info about Sophos for Virtual Environments
    
  ElseIf InstalledAV="Palo Alto Networks Traps" Then
    ObtainPaloAltoTrapsAVData 'Call the function to grab info about Palo Alto Networks Traps  

  ElseIf InstalledAV="SentinelOne" Then
    ObtainSentinelOneData 'Call the function to grab info about SentinelOne
    
  ElseIf InstalledAV="Bitdefender Endpoint Security Tools" Then
    ObtainBESTData 'Call the function to grab info about Bitdefender Endpoint Security Tools   

  ElseIf InstalledAV="Cb Defense Sensor" Then
    ObtainCarbonBlackData 'Call the function to grab info about Carbon Black 
 
  End If

  
      
  'Check to see if an instance of the WMI namespace exists; if it does, 
  'check to see if the WMI class exists. If the class exists, delete it, recreate it, and populate it
  If WMINamespaceExistanceCheck(strWMINamespace)="1" Then
      ' output.writeline "- The Namespace already exists."
      If WMIClassExists(strWMINamespace, strComputer,strWMIClassWithQuotes) Then
          ' output.writeline "- The WMI Class exists; let's delete it so that we don't have any duplicate data laying around in WMI."
          WMINamespace.Delete strWMIClassNoQuotes
          CreateWMIClass
          PopulateWMIClass
      Else
          ' output.writeline "- The Namespace exists, but the WMI class does not. Curious." 
          CreateWMIClass
          PopulateWMIClass      
      End If
  Else
      'Create the WMI Namespace (if it doesn't already exist), the WMI Class, and populate the class with data.              
      ' output.writeline "- The WMI Namespace and Class do not exist"
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
    ProgramFiles = WshShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%") 'It's useful to know if we need to access C:\Program Files or C:\Program Files(x86) - especially for Vipre A/V
    ProgramFiles64 = WshShell.ExpandEnvironmentStrings("%PROGRAMW6432%")
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
  
' Let's grab the %ProgramData% value, as it'll be used in detecting the installed AV product
ProgramData = WshShell.ExpandEnvironmentStrings("%PROGRAMDATA%")

' Let's figure out the locale of this device, so that we can correctly grab/parse dates in the correct format.
' output.writeline "- This device is in the following locale: " & GetLocale()   'Re-enable this line for debug purposes, if needed

End Sub   



  
' *****************************  
' Sub: DetectInstalledAV
' *****************************
Sub DetectInstalledAV
   output.writeline InstalledAV

    strMcAfeePath = Registry & "McAfee\AVEngine\DAT"
    strTrendKeyPath = Registry & "TrendMicro\PC-cillinNTCorp\CurrentVersion\Misc.\"
    strTrend7KeyPath = "SOFTWARE\TrendMicro\UniClient\1600\Update\PatternOutOfDateDays"
    strSymantecESKeyPath = Registry & "Symantec\Symantec Endpoint Protection\AV\"
    strSymantecAVKeyPath = Registry & "Symantec\Symantec AntiVirus\"
    strSophosAVKeyPath = Registry & "Sophos\"
    strSophosVirtualAVKeyPath = Registry32 & "Sophos\Sophos for Virtual Environments\"   
    strVIPREAVKeyPath = Registry & "Sunbelt Software\VIPRE Antivirus\"
    strVIPREEnterpriseKeyPath = Registry & "Sunbelt Software\Sunbelt Enterprise Agent\"
    strKasperskyAV2012KeyPath = Registry & "KasperskyLab\protected\AVP12\settings\"
    strKasperskyAV60KeyPath = Registry & "KasperskyLab\protected\AVP80\settings\"
    strSecurityEssentialsKeyPath = "SOFTWARE\Microsoft\Microsoft Security Client\"
    strEndpointSecurityKeyPath = Registry & "Microsoft\Windows\CurrentVersion\Uninstall\" 
    strKasperskyKES8KeyPath = Registry & "KasperskyLab\protected\KES8\settings\"
    strVIPREBusiness5KeyPath = Registry & "GFI Software\GFI Business Agent\"
    strVIPREBusinessOnlineKeyPath = Registry & "GFI Software\VIPRE Business Online\"
    strVIPREAV2012KeyPath = Registry & "GFI Software\VIPRE Antivirus\" 
    strESETKeyPath = "SOFTWARE\ESET\ESET Security\CurrentVersion\"  
    strKasperskyKES8ServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KAVFSEE\8.0.0.0\"
    strKasperskyKES6ServerKeyPath = Registry & "KasperskyLab\Components\34\Connectors\KAVFSEE\6.0.0.0\"
    strForefrontKeyPath = "SOFTWARE\Microsoft\Microsoft Forefront\Client Security\1.0\AM\"
    strKasperskySOS2KeyPath =  Registry & "KasperskyLab\protected\AVP9\settings\"
    strKasperskySOS3KeyPath =  Registry & "KasperskyLab\protected\ksos13\settings\"
    strTotalDefenseKeyPath =  "SOFTWARE\CA\TDClient\"
    strAviraKeyPath = Registry & "Avira\AntiVir Desktop\"
    strAviraServerKeyPath = Registry & "Avira\AntiVir Server\"
    strFSecureRegPath0= "SOFTWARE\Wow6432Node\Data Fellows\F-Secure\F-Secure GUI\PUB\" 
    strFSecureRegPath00= "SOFTWARE\Data Fellows\F-Secure\F-Secure GUI\PUB\"
    strFSecureRegPath000= "SOFTWARE\Wow6432Node\F-Secure\OneClient\" 
    strFSecureRegPath1= "SOFTWARE\Wow6432Node\Data Fellows\F-Secure\Anti-Virus\" 
    strFSecureRegPath2= "SOFTWARE\Data Fellows\F-Secure\Anti-Virus\" 
    strFSecureRegPath4= "SOFTWARE\Wow6432Node\F-Secure\Anti-Virus\" 
    strFSecureRegPath3= "SOFTWARE\F-Secure\Anti-Virus\" 
    strFSecureRegPath5= "SOFTWARE\Wow6432Node\Data Fellows\F-Secure\Anti-Virus Definition Databases\" 
    strFSecureRegPath6= "SOFTWARE\F-Secure\Anti-Virus Definition Databases\"
    strFSecureRegPath000= "SOFTWARE\Wow6432Node\F-Secure\OneClient\" 
    strKES10KeyPath = Registry & "KasperskyLab\protected\KES10\settings\"
    strKES10KeyPathSP1 = Registry & "KasperskyLab\protected\KES10SP1\settings\"
    strKES10KeyPathSP2 = Registry & "KasperskyLab\protected\KES10SP2\settings\"
    strWindowsDefenderPath = "SOFTWARE\Microsoft\Windows Defender\"
   	strPandaCloudEPPath64 = "Software\wow6432node\panda software\setup"
	strPandaCloudEPPath32 = "Software\panda software\setup"
    stravg2014regpath32 = "Software\Avg\Avg2014"
    stravg2014regpath64 = "Software\Wow6432Node\Avg\Avg2014"
    stravg2014defpath = ProgramData & "\AVG2014\avi"
    stravg2013regpath32 = "Software\Avg\Avg2013"
    stravg2013regpath64 = "Software\Wow6432Node\Avg\Avg2013"
    stravg2013defpath = ProgramData & "\AVG2013\avi"
    strAvastRegPath32 = "SOFTWARE\AVAST SOFTWARE\Avast"
    strAvastRegPath64 = "SOFTWARE\Wow6432Node\AVAST SOFTWARE\Avast"
    strViprebusinessAgt = "SOFTWARE\VIPRE Business Agent"
    strViprebusiness64Agt = "SOFTWARE\Wow6432Node\VIPRE Business Agent"
    strVipreBusinessAgtLoc= ""
	strWebRootStatusPath = Registry & "WRData\Status\"
	strWebRootStatusPath32 = "SOFTWARE\WRData\Status\"
	mcAfeeEndPointSecurityVersion = "SOFTWARE\McAfee\Endpoint\AV"
	mcAfeeEndpointSecurityDefDatePath = "SOFTWARE\McAfee\Endpoint\Common\TPSConnector\Subsystems\Update\Configuration"
	strSophosAVVersionPath = Registry & "Sophos\SAVService\Application"
	strSophosUpdateStatusPath = Registry & "Sophos\SAVService\Status\"
    strMalwareBytesRegPath64 = Registry & "Malwarebytes' Anti-Malware"
    strTMMSARegPath = "SOFTWARE\TrendMicro\ScanMail for Exchange\CurrentVersion\"
    strTMDSARegPath = "SOFTWARE\TrendMicro\Deep Security Agent\"
    stravg2016defpath = ProgramData & "\AVG\AV\avi"
    stravg2016regpath =  "Software\Wow6432Node\Avg\AV"
    strNormanregpath32 = "SOFTWARE\Norman Data Defense Systems"
    strNormanregpath64 = "SOFTWARE\Wow6432Node\Norman Data Defense Systems"
    strFortiClientPath = "SOFTWARE\Fortinet\FortiClient\FA_FMON"
    strPandaAdaptiveDefencePath64 = "Software\wow6432node\Panda Security\Nano Av\Setup"
    strPandaAdaptiveDefencePath32 = "Software\Panda Security\Nano Av\Setup"

    

    'Check if N-able's Endpoint Security is installed - if it is, we should exit the script immediately (dumping data into WMI negatively affects ES' ability to run scans)
    'We need to check two different registry values, as it changes depending on what OS is installed.
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strEndpointSecurityKeyPath & "AVTC64","DisplayName",ReturnedstrValue1
    If ReturnedstrValue1 = "Endpoint Security Manager" Then
      output.writeline "- N-able's Endpoint Security product has been detected on this machine. This script will now exit."
      wscript.quit(0)
    End If 
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strEndpointSecurityKeyPath & "AVNT64","DisplayName",ReturnedstrValue1
    If ReturnedstrValue1 = "Endpoint Security Manager" Then
      output.writeline "- N-able's Endpoint Security product has been detected on this machine. This script will now exit."
      wscript.quit(0)
    End If  
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strEndpointSecurityKeyPath & "AVTC","DisplayName",ReturnedstrValue1
    If ReturnedstrValue1 = "Endpoint Security Manager" Then
      output.writeline "- N-able's Endpoint Security product has been detected on this machine. This script will now exit."
      wscript.quit(0)
    End If 
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strEndpointSecurityKeyPath & "AVNT","DisplayName",ReturnedstrValue1
    If ReturnedstrValue1 = "Endpoint Security Manager" Then
      output.writeline "- N-able's Endpoint Security product has been detected on this machine. This script will now exit."
      wscript.quit(0)
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    ' Check for AV Defender on 64-bit, modern operating systems
    If (objFSO.FolderExists(ProgramData & "\N-able Technologies\AVDefender\Configuration") AND isProcessRunning(".","NableAVDBridge.exe")) Then
      output.writeline "- AV Defender has been detected on this machine. This should be monitored via the dedicated AV Defender Status Service. This script will now exit."
      wscript.quit(0)
    End If
    
    ' Check for AV Defender on XP and 2K3
    If (objFSO.FolderExists("C:\Documents and Settings\All Users\Application Data\N-able Technologies\AVDefender\Configuration") AND isProcessRunning(".","NableAVDBridge.exe")) Then
      output.writeline "- AV Defender has been detected on this machine. This should be monitored via the dedicated AV Defender Status Service. This script will now exit."
      wscript.quit(0)
    End If
  

    'Check to see what A/V product is installed
    
    If RegKeyExists("HKLM\" & strTrendKeyPath & "ProductName") Then
      strValue = "ProductName"
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strTrendKeyPath,strValue,RegstrValue
      InstalledAV = RegstrValue
      
      
    ElseIf (objFSO.FolderExists(ProgramData & "\Bitdefender\Endpoint Security") AND isProcessRunning(".","epag.exe")) Then
      InstalledAV = "Bitdefender Endpoint Security Tools"
      

    ElseIf objFSO.FolderExists(ProgramFiles64 & "\SentinelOne") Then
    'ElseIf (objFSO.FolderExists(ProgramFiles64 & "\SentinelOne") AND isProcessRunning(".","SentinelAgent.exe")) Then
      InstalledAV = "SentinelOne"


    ElseIf RegKeyExists("HKLM\" & strTrendKeyPath & "ProgramVer") Then
      InstalledAV = "Trend Micro OfficeScan"
                                                                              

    Elseif RegKeyExists("HKLM\" & strSymantecESKeyPath & "ScanEngineVendor") Then
      InstalledAV = "Symantec Endpoint Protection"
      

    ElseIf RegKeyExists ("HKLM\" & strSymantecAVKeyPath & "CorporateFeatures") Then
       InstalledAV = "Symantec AntiVirus"
          
       
    ElseIf RegKeyExists ("HKLM\" & strTrend7KeyPath) Then 
       InstalledAV = "Trend Micro Worry-Free Business Security 7"

       
       
    ElseIf RegKeyExists ("HKEY_LOCAL_MACHINE\" & strMcAfeePath) Then
       InstalledAV = "McAfee AntiVirus"


    ElseIf RegKeyExists ("HKLM\" & strVIPREAVKeyPath & "ProductCode") Then
       InstalledAV = "VIPRE AntiVirus"



    ElseIf RegKeyExists ("HKLM\" & strVIPREEnterpriseKeyPath & "ProductCode") Then
       InstalledAV = "Sunbelt Enterprise Agent"

       
       
    ElseIf RegKeyExists ("HKLM\" & strKasperskyAV2012KeyPath & "SettingsVersion") Then
       InstalledAV = "Kaspersky Anti-Virus 2012"


       
    ElseIf RegKeyExists ("HKLM\" & strKasperskyAV60KeyPath & "SettingsVersion") Then
       InstalledAV = "Kaspersky Anti-Virus 6.0"
       
       
    ElseIf RegKeyExists ("HKLM\" & strKasperskyKES8KeyPath & "SettingsVersion") Then
       InstalledAV = "Kaspersky Endpoint Security 8"
       
    '--- Check for VIPRE Business Agent  ---	  
    ElseIf RegKeyExists ("HKLM\" & strViprebusinessAgt & "\Version") Then
	   InstalledAV = "VIPRE Business Agent"
	   objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strViprebusinessAgt ,"InstallPath",strVipreBusinessAgtLoc
        strViprebusinessAgt1 = strViprebusinessAgt
    ElseIf RegKeyExists ("HKLM\" & strViprebusiness64Agt & "\Version") Then
	  InstalledAV = "VIPRE Business Agent"
	  objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strViprebusiness64Agt ,"InstallPath",strVipreBusinessAgtLoc
        strViprebusinessAgt1 = strViprebusiness64Agt

    '--- Check for VIPRE Business Antivirus  ---    
    ElseIf RegKeyExists ("HKLM\" & strVIPREBusiness5KeyPath & "ProductCode") Then
       InstalledAV = "VIPRE Business Antivirus"
       strVIPREEnterpriseKeyPath = Registry & "GFI Software\GFI Business Agent\"
       
    ElseIf RegKeyExists ("HKLM\" & strVIPREBusinessOnlineKeyPath) Then
       InstalledAV = "VIPRE Business Online"
       strVIPREEnterpriseKeyPath = Registry & "GFI Software\VIPRE Business Online\"
       strViprebusinessAgt1 = strVIPREBusinessOnlineKeyPath
             
    ElseIf RegKeyExists ("HKLM\" & strVIPREAV2012KeyPath & "ProductCode") Then
       InstalledAV = "VIPRE Antivirus 2012"
       strVIPREAVKeyPath = Registry & "GFI Software\VIPRE Antivirus\"
       
            
    ElseIf RegKeyExists ("HKLM\" & strESETKeyPath & "\info\ProductName") Then
       objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\ESET\ESET Security\CurrentVersion\info\","ProductName",InstalledAV 

       
       
    ElseIf RegKeyExists ("HKLM\" & strKasperskyKES8ServerKeyPath & "ProdDisplayName") Then
       InstalledAV = "Kaspersky Anti-Virus 8.0 For Windows Servers Enterprise Edition"
       

    ElseIf RegKeyExists ("HKLM\" & strKasperskyKES6ServerKeyPath & "ProdDisplayName") Then
       InstalledAV = "Kaspersky Anti-Virus 6.0 For Windows Servers Enterprise Edition"
       
       
    ElseIf RegKeyExists ("HKLM\" & strForefrontKeyPath & "InstallLocation") Then
       InstalledAV = "Microsoft Forefront"
       
       
    ElseIf RegKeyExists ("HKLM\" & strKasperskySOS2KeyPath & "Ins_DisplayName") Then
       InstalledAV = "Kaspersky Small Office Security"

    ElseIf RegKeyExists ("HKLM\" & strKasperskySOS3KeyPath ) Then
       InstalledAV = "Kaspersky Small Office Security 3"


    ElseIf RegKeyExists ("HKLM\" & strTotalDefenseKeyPath & "ProductName") Then
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strTotalDefenseKeyPath,ProductName,InstalledAV
      InstalledAV = "Total Defense R12 Client"
       
       
    ElseIf RegKeyExists ("HKLM\" & strAviraKeyPath & "EngineVersion") Then
       InstalledAV = "Avira AntiVirus"
       
       
    ElseIf RegKeyExists ("HKLM\" & strAviraServerKeyPath & "EngineVersion") Then
       InstalledAV = "Avira AntiVirus"
             
    ElseIf  WMINamespaceExistanceCheck("FSECURE") Then 'If the F-Secure WMI namespace exists, let's use that. If not, we'll fall back to the alternate method.
        Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\FSECURE")
        ' Let's determine what version of F-Secure has been installed
        Set colItems = objWMIService.ExecQuery("SELECT Version,Name FROM Product", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
        For Each objItem in colItems
            FormattedAVVersion = objItem.Version
            InstalledAV = objItem.Name
        Next
           
    ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath0 & "ProductName") Then
    output.writeline "- Found it!"
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPath0,"ProductName",InstalledAV      

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
      End If
                   
                   
      ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath00 & "ProductName") Then
       objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPath00,"ProductName",InstalledAV
       

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
        End If


    ElseIf RegKeyExists ("HKLM\" & strFSecureRegPath000 & "ProductName") Then
       objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPath000,"ProductName",InstalledAV
       

                                                

  ElseIf RegKeyExists ("HKLM\" & strKES10KeyPath & "SettingsVersion") Then
      InstalledAV = "Kaspersky Endpoint Security 10"
  
  ElseIf RegKeyExists ("HKLM\" & strKES10KeyPathSP1 & "SettingsVersion") Then
      InstalledAV = "Kaspersky Endpoint Security 10 SP1"
  
  ElseIf RegKeyExists ("HKLM\" & strKES10KeyPathSP2 & "SettingsVersion") Then
      InstalledAV = "Kaspersky Endpoint Security 10 SP2"
      
  ElseIf RegKeyExists ("HKLM\" & Registry & "KasperskyLab\Components\34\Connectors\KES\11.0.0.0\" & "ConnectorVersion") Then
      InstalledAV = "Kaspersky Endpoint Security 11 for Windows"
  
  ElseIf RegKeyExists ("HKLM\" & Registry & "KasperskyLab\Components\34\1103\1.0.0.0\Statistics\AVState\" & "Protection_NagentVersion") Then
      InstalledAV = "Kaspersky Endpoint Security 10 for Windows"
      
  ElseIf RegKeyExists ( "HKLM\" & strPandaAdaptiveDefencePath32 & "\Path") Then
      InstalledAV = "Panda Adaptive Defense 360 32 Bit"	

  ElseIf RegKeyExists ( "HKLM\" & strPandaAdaptiveDefencePath64 & "\Path") Then
      InstalledAV = "Panda Adaptive Defense 360 64 Bit"

  ElseIf RegKeyExists ( "HKLM\" & strPandaCloudEPPath32 & "\InitialProductName") Then
      InstalledAV = "Panda Endpoint Protection 10 32 Bit"	

  ElseIf RegKeyExists ( "HKLM\" & strPandaCloudEPPath64 & "\InitialProductName") Then
      InstalledAV = "Panda Endpoint Protection 10 64 Bit"
	
  ElseIf RegKeyExists ( "HKLM\" & stravg2013regpath32 & "\ProdType") Then
      InstalledAV = "AVG 2013"
	
  ElseIf RegKeyExists ( "HKLM\" & stravg2013regpath64 & "\ProdType") Then
      InstalledAV = "AVG 2013"

  ElseIf RegKeyExists ( "HKLM\" & stravg2014regpath32 & "\ProdType") Then
      InstalledAV = "AVG 2014"
	
  ElseIf RegKeyExists ( "HKLM\" & stravg2014regpath64 & "\ProdType") Then
      InstalledAV = "AVG 2014"
      
  ElseIf RegKeyExists ( "HKLM\" & stravg2016regpath & "\ProdType") Then
      InstalledAV = "AVG Protection"    


  '--- Check for Webroot Anywhere ---	  
  ElseIf RegKeyExists ("HKLM\" & strWebRootStatusPath & "Version") Then
	  InstalledAV = "Webroot SecureAnywhere"

  '--- Check for Webroot Anywhere 32 bit on 64 bit computer ---	  
  ElseIf RegKeyExists ("HKLM\" & strWebRootStatusPath32 & "Version") Then
	  InstalledAV = "Webroot SecureAnywhere"
	  strWebRootStatusPath = strWebRootStatusPath32

  '--- Check for AVAST! ---	  
  ElseIf RegKeyExists ("HKLM\" & strAvastRegPath32 & "\Version") Then
	  InstalledAV = "Avast!"
	  objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strAvastRegPath32,"ProgramFolder",strAvastInstallPath  

  '--- Check for AVAST! ---	  
  ElseIf RegKeyExists ("HKLM\" & strAvastRegPath64 & "\Version") Then
	  InstalledAV = "Avast!"
	  objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strAvastRegPath64,"ProgramFolder",strAvastInstallPath  




  '--- Check for McAfee Endpoint Security V10  ---	  
  ElseIf RegKeyExists ("HKLM\" & mcAfeeEndpointSecurityDefDatePath ) Then
	  InstalledAV = "McAfee Endpoint Security"


  '--- Check for McAfee Endpoint Security V10.1  ---	  
  ElseIf RegKeyExists ("HKLM\" & mcAfeeEndPointSecurityVersion & "\ProductVersion" ) Then
	  InstalledAV = "McAfee Endpoint Security 10.1"

    '--- Check for Symantec Endpoint Protection Cloud  ---                
    ElseIf RegKeyExists ("HKLM\SOFTWARE\Norton\SecurityStatusSDK\SDKProduct") Then
        strSEPCloudRegPath0 = "SOFTWARE\Norton\"
        InstalledAV="Symantec Endpoint Protection Cloud"
        FoundGUID = FALSE
        Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
        Set arrSubKeys = Nothing
        oReg.EnumKey HKEY_LOCAL_MACHINE, strSEPCloudRegPath0, arrSubKeys
    
        If Not IsNull (arrSubKeys) Then
            For Each subkey In arrSubKeys
  	         ' Endpoint Cloud stores the information we're after in a GUID-titled sub-key. Let's figure out that GUID.
                If RegKeyExists("HKLM\" & strSEPCloudRegPath0  & subkey & "\PRODUCTVERSION") Then
                    strSEPCloudRegPath2 = strSEPCloudRegPath0 & subkey
                    FoundGUID = TRUE
                    Exit For         
  		        End If 
  	         Next
      
        ' It looks like newer versions of Endpoint Cloud no longer log to the Wow6432Node on 64-bit machines. This next bit of code will handle that scenario
        If FoundGUID = FALSE Then
            Set arrSubKeys = Nothing
            oReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\NORTON\", arrSubKeys
            For Each subkey In arrSubKeys
    		  ' Endpoint Cloud stores the information we're after in a GUID-titled sub-key. Let's figure out that GUID.
                If RegKeyExists("HKLM\SOFTWARE\NORTON\"  & subkey & "\PRODUCTVERSION") Then
                    strSEPCloudRegPath2 = strSEPCloudRegPath0 & subkey
                    Exit For         
    		  End If 
  		    Next
        End If
    Else
        output.writeline "- Unable to find more details about Symantec Endpoint Protection Cloud from the registry. Is it perhaps uninstalled?"
        Wscript.Quit
    End If

           
              

    
    '--- Check for Trend Micro Messaging Security Agent  ---                                    
    ElseIf RegKeyExists("HKLM\" & strTMMSARegPath & "DebugLevel") Then
      InstalledAV = "Trend Micro Messaging Security Agent"
      
      
    '--- Check for Trend Micro Deep Security Agent  ---                                    
    ElseIf objFSO.FileExists(ProgramFiles64 & "\Trend Micro\Deep Security Agent\dsa.exe") Then
      InstalledAV = "Trend Micro Deep Security Agent"


    '--- Check for McAfee Move  ---                                   
    ElseIf objFSO.FileExists(ProgramFiles & "\McAfee\MOVE AV Client\mvadm.exe") Then
      InstalledAV = "McAfee Move AV Client"
      
      
    '--- Check for Norman Endpoint Protection 32-bit ---                                    
    ElseIf RegKeyExists("HKLM\" & strNormanregpath32 & "\RootPath") Then
      InstalledAV = "Norman Endpoint Protection"
      objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strNormanregpath32,"RootPath",strNormanrootpath

    '--- Check for Norman Endpoint Protection 64-bit ---                                    
    ElseIf RegKeyExists("HKLM\" & strNormanregpath64 & "\RootPath") Then
      InstalledAV = "Norman Endpoint Protection"
      objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strNormanregpath64,"RootPath",strNormanrootpath


    '--- Check for Cylance PROTECT  ---                                    
    ElseIf RegKeyExists("HKLM\SOFTWARE\Cylance\Desktop\Path") Then
      InstalledAV = "Cylance PROTECT"
      'Cylance is a different type of AV product - it doesn't have the traditional concept of AV definition updates.
      ' As a result, all we're looking for here is if the Cylance process is running; all other values will be hard-coded.
      ServiceActive = isProcessRunning(strComputer,"cylancesvc.exe")
	    If (ServiceActive) Then
		    OnAccessScanningEnabled = TRUE
	    Else
		    OnAccessScanningEnabled = FALSE
	   End If 
     
     output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled 
     ProductUpToDate = "TRUE"
     FormattedAVVersion = "Unknown"  
     
    '--- Check for AVG Business Security ---                                    
    ElseIf objFSO.FileExists(ProgramData & "\AVG\Persistent Data\Antivirus\Logs\update.log") Then
      InstalledAV = "AVG Business Security"
      
    '--- Check for FortiClient ---   
    ElseIf RegKeyExists ("HKLM\" & strFortiClientPath & "\enabled") Then
        InstalledAV = "FortiClient"
      
    '--- Check for Sophos AV ---   
    ElseIf RegKeyExists ("HKLM\" & strSophosAVVersionPath & "\MarketingVersion") Then
        objReg.GetStringValue HKEY_LOCAL_MACHINE,strSophosAVVersionPath,"MarketingVersion",RegstrValue
        If inStr(RegstrValue, "9.")  Then
            InstalledAV = "Sophos Anti-Virus"
        ElseIf inStr(RegstrValue, "10.1") OR  inStr(RegstrValue, "10.2") OR  inStr(RegstrValue, "10.3") Then
            InstalledAV = "Sophos Anti-Virus"
        ElseIf inStr(RegstrValue, "2.2.")  Then
            InstalledAV = "Sophos Endpoint Protection"
        Else
		    InstalledAV = "Sophos Anti-Virus 10" 
	   End If

    '--- Check for Sophos for Virtual Environments --- 
    ElseIf RegKeyExists ("HKLM\" & strSophosVirtualAVKeyPath & "SGVM Deployment Service\InstalledPath") Then
        InstalledAV = "Sophos for Virtual Environments"    
        
    '--- Check for Carbon Black a.k.a. "Cb Defense Sensor 64-bit" --- 
    ElseIf isServiceRunning("CbDefense") Then
        InstalledAV = "Cb Defense Sensor"     


    '--- Check for the ATP-managed version of SCEP, and Microsoft Security Essentials  ---    
    ElseIf RegKeyExists ("HKLM\SOFTWARE\Microsoft\Microsoft Antimalware\InstallLocation") Then
       ' We know that either Microsoft SCEP is installed, or Microsoft Security Essentials. Now let's figure out which one.
       objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Security Client\","DisplayName",SCEPInstalled
       objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Security Client\","UninstallString",SCEPUninstallString
       If InStr(SCEPInstalled,"System Center Endpoint Protection") AND InStr(SCEPUninstallString,"Managed Defender") Then
        InstalledAV = "Microsoft System Center Endpoint Protection (Managed Defender)"     
       ElseIf InStr(SCEPInstalled,"System Center Endpoint Protection") AND InStr(SCEPUninstallString,"Microsoft Security Client") Then
        InstalledAV = "Microsoft System Center Endpoint Protection"
       ElseIf InStr(SCEPInstalled,"System Center 2012 Endpoint Protection") Then
        InstalledAV = "Microsoft System Center 2012 Endpoint Protection" 
       Else
        InstalledAV = "Microsoft Security Essentials"
       End If


        
    '--- Check for Microsoft System Center Endpoint Protection  ---	
    ElseIf RegKeyExists ("HKLM\" & strSecurityEssentialsKeyPath & "LastSuccessfullyAppliedPolicy") Then
        InstalledAV = "Microsoft System Center Endpoint Protection"
 
      
    '--- Check for Cisco AMP ---
    ElseIf objFSO.FileExists(ProgramFiles64 & "\Cisco\AMP\local.xml") Then
        InstalledAV = "Cisco Advanced Malware Protection (AMP)"

      
    '--- Check for Malwarebytes' Corporate Edition  ---	  
    ElseIf RegKeyExists ("HKLM\" & strMalwareBytesRegPath64 & "\InstallPath" ) Then
	   InstalledAV = "Malwarebytes' Corporate Edition"
       
       
    '--- Check for Palo Alto Traps  ---	
    ElseIf RegKeyExists ("HKLM\SOFTWARE\Palo Alto Networks\Traps\ProtectionStatus") Then
        InstalledAV = "Palo Alto Networks Traps"
 
       
                  	  

    Else 
  
    If WMINamespaceExistanceCheck(strWMINameSpace2) = "1" Then 'Check to make sure that the SecurityCenter2 namespace exists before we do this check, lest the script error out when checking for the presence of the AntiVirusProduct WMI class   
        If WMIClassExists (strWMINameSpace2, strComputer, "AntiVirusProduct") Then
            output.writeline "- Looking in the root\SecurityCenter2 WMI Namespace."
            ' This next line calls the 'ObtainSecurityCenter2Data sub-routine, which queries the root\SecurityCenter2 WMI Namespace for info.
            ObtainSecurityCenter2Data                 
        ElseIf RegKeyExists ("HKLM\" & strWindowsDefenderPath & "ProductStatus") Then
            output.writeline "- The root\securityCenter2 namespace exists, but the AntiVirusProduct WMI class does not - that's really weird."
            output.writeline "- Windows Defender has been detected, by querying the registry."
            InstalledAV = "Windows Defender"
        End If
    ' If the root\SecurityCenter2 WMI namespace doesn't exist (and it doesn't if you're looking at a Server-class OS, then let's see if Windows Defender is installed)
    ElseIf RegKeyExists ("HKLM\" & strWindowsDefenderPath & "ProductStatus") Then
        output.writeline "- Windows Defender has been detected, by querying the registry."
        InstalledAV = "Windows Defender"
             
    Else 

      If NoAVBehavior = "WRITE" Then
        output.writeline "- Unable to determine installed AV."
        InstalledAV = "An Unknown A/V product, or no A/V product at all"
        onAccessScanningEnabled = "FALSE"
        ProductUpToDate = "FALSE"
        FormattedAVVersion = "Unknown"
      ElseIf NoAVBehavior = "DONOTWRITE" Then
        output.writeline "- The script could not detect an A/V product on this device. As the DONOTWRITE command-line parameter was specified, no data will be written to WMI, and this script will now exit."
        ' We need to clear out the existing data from the AntiVirusProduct WMI Class
        If WMIClassExists(strWMINamespace, strComputer,strWMIClassWithQuotes) Then
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
    
  End If
  
End Sub  


' *****************************  
' Function: ObtainTrendMicroData
' *****************************
Sub ObtainTrendMicroData
    'Grab the A/V Pattern Version from the Registry
    strTrendRealTimeKeyPath = Registry & "TrendMicro\PC-cillinNTCorp\CurrentVersion\Real Time Scan Configuration\"
    strTrendProductVersion =  Registry & "TrendMicro\PC-cillinNTCorp\CurrentVersion\Misc.\" 
    InputRegistryKey1 =  "InternalPatternVer"
    InputRegistryKey2 =  "PatternDate"
    InputRegistryKey3 = "InternalNonCrcPatternVer"
    InputRegistryKey4 = "Enable"
    InputRegistryKey5 =  "NonCrcPatternDate"
    InputRegistryKey6 = "LastUpdate"
    InputRegistryKey7 = "TmListen_Ver"

    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strTrendRealTimeKeyPath,InputRegistryKey4,ReturneddwValue3
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strTrendKeyPath,InputRegistryKey1,ReturneddwValue1
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strTrendKeyPath,InputRegistryKey3,ReturneddwValue2
    
    
    ' Let's grab the version of Trend Micro that is installed on this machine
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strTrendProductVersion,InputRegistryKey7,ProductVersion

   If ReturneddwValue3 = 1 Then
    OnAccessScanningEnabled = TRUE
   ElseIf ReturneddwValue3 = 0 Then
    OnAccessScanningEnabled = FALSE
   Else
    OnAccessScanningEnabled = FALSE
   End If
   
   
   output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled 
    
 
    
    ' If Smart Scan is installed, the version of the A/V definition that's being used will be reported in the InternalNonCrcPatternVer key, 
    ' and the InternalPatternVer key will be 0.
    ' If Smart Scan isn't installed, the version of the A/V definition that's being used will be reported in the InternalPatternVer key, 
    ' and the InternalNonCrcPatternVer key will be 0.
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,Registry & "TrendMicro\PC-cillinNTCorp\CurrentVersion\iCRC Scan\","scantype",ReturneddwValue4
    If ReturneddwValue4 = 1 Then
        RawAVVersion =  cstr(ReturneddwValue2)
        output.writeline "- Smart Scan is enabled, so we'll use the InternalNonCrcPatternVer registry value."
    Else
        output.writeline "- Smart Scan must not be installed. Using the InternalPatternVer registry key."
        RawAVVersion =  ReturneddwValue1
    End If 
    
    
    'Convert the dwValue variable from hex to decimal, and pretty it up
	   strTrendVerLen = Len(RawAVVersion)
	   If strTrendVerLen = 6 Then
		  FormattedAVVersion =  Left(RawAVVersion,1) & "." & Mid(RawAVVersion,2,3) & "." & Right(RawAVVersion,2)
     ElseIf strTrendVerLen = 7 Then
		  FormattedAVVersion =  Left(RawAVVersion,2) & "." & Mid(RawAVVersion,3,3) & "." & Right(RawAVVersion,2)
	   Else
		  FormattedAVVersion = "No Version"
	   End If
   
    
    ' Format the Pattern Age that was in the registry to make it more human-readable
    If InStr(InstalledAV,"OfficeScan") Then 'The date for OfficeScan is stored in a nice,human-readable format that just needs to have slashes inserted 
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strTrendKeyPath,InputRegistryKey2,ReturnedstrValue1
      FormattedPatternAge = Left(ReturnedstrValue1,4) & "/" & Mid(ReturnedstrValue1,5,2) & "/" & Right(ReturnedstrValue1,2)
    Else 'The date for other Trend products is stored as a timestamp value that needs to be converted into a date
      objReg.GetDWORDValue HKEY_LOCAL_MACHINE,Registry & "TrendMicro\PC-cillinNTCorp\CurrentVersion\UpdateInfo\",InputRegistryKey6,ReturnedstrValue1
      FormattedPatternAge = DateAdd("s", ReturnedstrValue1, "01/01/1970 00:00:00")
    End If
    
    ' Calculate how old the A/V Pattern really is
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate) 
   
    output.writeline "- The A/V product version is: " & FormattedAVVersion
    output.writeline "- The date of the A/V Definition File is: " & FormattedPatternAge
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
    ElseIf SymantecAvDate = "0" Then
            SymantecAvDate = "01"            
    End If

    FormattedPatternAge =  SymantecAVYear & "/" & SymantecAvMonth & "/" & SymantecAvDate
    output.writeline "- The formatted date of the A/V Definition File is: " & FormattedPatternAge
    
    FormattedAVVersion = FormattedPatternAge & " r" & Revision
    
    output.writeline "- The formatted A/V Version is: " & FormattedAVVersion
    
    'Calculate how old the A/V Pattern really is
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
    


End Sub



 
 
 
 
 
 
 ' *****************************  
' Sub: ObtainSophosAVData
' *****************************
Sub ObtainSophosAVData
  ' Let's call the Sophos script/function so that the registry gets populated
  SophosScript

  ' Let's figure out when Sophos was last updated
  If RegKeyExists ("HKLM\SOFTWARE\WOW6432Node\Sophos\SavService\Status\UpToDateState") Then
    Output.Writeline "- The UpToDateState registry key has been found. Let's use that to determine whether or not Sophos is up-to-date."
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,"SOFTWARE\WOW6432Node\Sophos\SavService\Status\","UpToDateState",UpToDateState
    If UpToDateState = 1 Then
      ProductUpToDate = TRUE 
    Else
      ProductUpToDate = FALSE
    End If
    
    Output.Writeline "- Is Sophos up-to-date? " & ProductUpToDate
      
  Else
    ' Note that as of Dec 2018, this mechanism is not recommended by Sophos anymore. I'm only keeping it here for backwards compatibility.
    Output.Writeline "- The UpToDateState registry was not found. Querying the Sophos API directly."
    Set objComponentMgr = CreateObject("Infrastructure.ComponentManager.1")
    Set objConfigMgr = objComponentMgr.FindComponent("ConfigurationManager")
    Set objNode = objConfigMgr.GetNode(2, "ProductInfo/updateDate")      
    FormattedPatternAge = objNode.GetAttributeValue("year") & "/" & objNode.GetAttributeValue("month") & "/" & objNode.GetAttributeValue("day") 
    output.writeline   "- The A/V Definition file for Sophos was last updated on: " & FormattedPatternAge
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
    output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
    CalculateAVAge 'Call the function to determine how old the A/V Definitions are     
  End If
    


End Sub


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
  
  '(SavControl property name, key to create, value name, default, type )
  properties  = Array(_   
                  Array("Version.Major", "SavService\Version", "Major", "-1", DWORD_VALUE), _
                  Array("Version.Minor", "SavService\Version", "Minor", "-1", DWORD_VALUE), _
                  Array("Version.AV.Data", "SavService\Version", "Data",  "", STRING_VALUE), _
                  Array("Version.Extra", "SavService\Version", "Extra", "", STRING_VALUE), _
                  Array("Protection.OnAccess", "SavService\Status\Policy", "OnAccessScanningEnabled", "-1", BOOL_VALUE) _
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
          output.writeline "- Update in progress; can't make any queries to Sophos right now."
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
          'Values available directly from SAVControl
          Dim i
          i =0
          Dim p
          For Each p In properties
              value = ""
              value = sc.GetProperty(p(SAV_PROPERTY_NAME))
              If Err.number = 0 And CStr(value) <> "" Then
                  properties(i)(THE_VALUE) = value
                                   output.writeline "- Got " & p(SAV_PROPERTY_NAME) & "; it reported a value of " & value 
              Else
                  output.writeline "- Failed to get " & p(SAV_PROPERTY_NAME) & ". The error code returned was: " & Err.number
              End If
              i = i+1
          Next
      Else
          output.writeline "- Failed to create SavControl object"
      End If
      
  
      value = ""
      value = GetLastUpdatedTimeOfSAV(oReg)
      If value <> "" Then
          LastUpdatedStr= value
      End If
  End If 'not updating
  
  'set formatted AV Version and onaccess scannning state
                  FormattedAVVersion = left(properties(0)(THE_VALUE),2) & "." & right(properties(1)(THE_VALUE),1) & " virus data " & properties(2)(THE_VALUE) 
                  output.writeline "- The A/V Version is: " & FormattedAVVersion
                  OnAccessScanningEnabled = properties(4)(THE_VALUE)
      output.writeline "- Is Real-Time Scanning Enabled? " & OnAccessScanningEnabled
      
  
  
  Err.number = 0


End Sub

 ' *****************************  
' Function: ObtainSophos10AVData
' *****************************
Sub ObtainSophos10AVData
  Set objComponentMgr = CreateObject("Infrastructure.ComponentManager")
  Set objConfigMgr = objComponentMgr.FindComponent("ConfigurationManager")
  

  ' Let's figure out when Sophos was last updated
  If RegKeyExists ("HKLM\SOFTWARE\WOW6432Node\Sophos\SavService\Status\UpToDateState") Then
    ' output.Writeline "- The UpToDateState registry key has been found. Let's use that to determine whether or not Sophos is up-to-date."
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,"SOFTWARE\WOW6432Node\Sophos\SavService\Status\","UpToDateState",UpToDateState
    If UpToDateState = 1 Then
      ProductUpToDate = TRUE 
    Else
      ProductUpToDate = FALSE
    End If
    
    output.Writeline "- Is Sophos up-to-date? " & ProductUpToDate
      
  Else
    ' Note that as of Dec 2018, this mechanism is not recommended by Sophos anymore. I'm only keeping it here for backwards compatibility.
    output.Writeline "- The UpToDateState registry was not found. Querying the Sophos API directly."
    Set objNode = objConfigMgr.GetNode(2, "ProductInfo/updateDate")      
    FormattedPatternAge = objNode.GetAttributeValue("year") & "/" & objNode.GetAttributeValue("month") & "/" & objNode.GetAttributeValue("day") 
    output.writeline   "- The A/V Definition file for Sophos was last updated on: " & FormattedPatternAge
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
    output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
    CalculateAVAge 'Call the function to determine how old the A/V Definitions are     
  End If
  
  
  ' Let's find out what version of Sophos is installed on this machine
  If InstalledAV = "Sophos Endpoint Protection" Then
    FormattedAVVersion = RegstrValue
  Else 'This chunk of code works perfectly for Sophos AV 10, but not for Sophos Endpoint Protection - so I've split things apart
    Set node = objConfigMgr.GetNode(0, "ProductInfo/productVersion")   
    Dim verString 
    FormattedAVVersion = node.GetAttributeValue("major") & "." &_
            node.GetAttributeValue("minor") & "." _
     & node.GetAttributeValue("build") 
  End If

  

  
  
  ' Let's see whether or not On-Access Scanning is enabled
  Dim icManager
  Set icManager = objComponentMgr.FindComponent("ICManager")
  
  Select Case icManager.GetFilterState
      Case 1
         OnAccessScanningEnabled = FALSE
      Case 2
         OnAccessScanningEnabled = TRUE
      Case Else
         OnAccessScanningEnabled = TRUE 
  End Select 
  
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled 
  
  Set icManager = Nothing

    
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
   Else 
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
    output.writeline "- Creating the WMI namespace"
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
    ' output.writeline "- Creating the WMI class"
    objClassCreator.Put_

End Sub
    
    
' *****************************  
' Function: WMINamespaceExistanceCheck
' Thanks to http://www.cruto.com/resources/vbscript/vbscript-examples/misc/wmi/List-All-WMI-Namespaces.asp for this code 
' *****************************
Function WMINamespaceExistanceCheck(WMINamespace)
                WMINamespaceExistanceCheck = "0"
                Set NamespacetoCheck = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\root")
                Set colNamespaces = NamespacetoCheck.InstancesOf("__Namespace")
                For Each objNamespace In colNamespaces
                      If instr(objNamespace.Path_.Path,WMINamespace & chr(34)) Then
                        WMINamespaceExistanceCheck = "1"
                         ' output.writeline "- Found the " & WMINamespace & " WMI namespace."   Turn this on for debugging purposes, if needed.                     
                        Exit For
                      End If
                       
                Next
                If WMINamespaceExistanceCheck = "0" Then
                  ' output.writeline "- Unable to find the root\" &  WMINamespace & " WMI namespace." Turn this on for debugging purposes, if needed. 
                End If
                Set colNamespaces = Nothing
End Function


  


' *****************************  
' Function: WMIClassExists
' Thanks to http://gallery.technet.microsoft.com/ScriptCenter/en-us/a1b23364-34cb-4b2c-9629-0770c1d22ff0 for this code 
' *****************************
Function WMIClassExists(strWMINamespace, strComputer, strWMIClassWithQuotes)
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

	' Added to script by Tim Wiser 10/03/2015 to catch rare cases where no AV is installed yet the 
	 ' two health variables were being set incorrectly.
	 If InstalledAV = "No AV installed" Then
	   ProductUpToDate = FALSE
	   OnAccessScanningEnabled = FALSE
	 Else
    If (InstalledAV1 = "ESET File Security" OR InstalledAV1 = "ESET Mail Security") Then
      InstalledAV = InstalledAV1 & " " & ProductVersion & "<br/>" & "Status Message from ESET: " & StatusText 
    Else
      InstalledAV = InstalledAV & " " & ProductVersion
    End If  
   End If
   
   'Debug code, if needed
   ' output.writeline "- Version Number: " & FormattedAVVersion
   ' output.writeline "- Installed AV: " & InstalledAV
   ' output.writeline "- On-Access Scanning: " & OnAccessScanningEnabled
   ' output.writeline "- Product Up-to-Date: " & ProductUpToDate
   
   
	 'Create an instance of the WMI class using SpawnInstance_
    Set WMINamespace = GetObject("winmgmts:\\" & strComputer & "\root\" & strWMINamespace)
    Set objGetClass = WMINamespace.Get(strWMIClassNoQuotes)
    Set objNewInstance = objGetClass.SpawnInstance_
    objNewInstance.VersionNumber =  FormattedAVVersion
    objNewInstance.displayName = InstalledAV
    objNewInstance.onAccessScanningEnabled = OnAccessScanningEnabled
    objNewInstance.ProductUpToDate = ProductUpToDate
    objNewInstance.ScriptExecutionTime = FormatDateTime(Date & " " & Time)
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
      strMcAfeePath = Registry & "McAfee\AVEngine\"
      strMcAfeeVersionPath = Registry & "McAfee\DesktopProtection\"
      strMcAfeeOASPath = Registry & "Mcafee\DesktopProtection\"
      'Lets find the state of the OAS  
	objReg.GetStringValue HKEY_LOCAL_MACHINE,strMcAfeeVersionPath,ProductVersionKey,ProductVersion
	If InStr(ProductVersion,"8.7") Then    'It's version 8.7 or older!
        objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strMcAfeeOASPath,InputRegistryKey3,RAWOASEnabled
          If RAWOASEnabled = 1 Then
            OnAccessScanningEnabled = TRUE
          Else
            OnAccessScanningEnabled = FALSE
         End If
      Else 'It's at least version 8.8!
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
 
    
    
    FormattedAVVersion = AvDatDate & " SDAT: " & AVDatVersion
    
    output.writeline "- The version of the A/V Definition File is: " & FormattedAVVersion
    output.writeline "- The version of McAfee AntiVirus running on this machine is: " & ProductVersion
    
    
    'Calculate how old the A/V Pattern really is
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



' *************************************  
' Function: ObtainVIPREBusinessAgentData
' *************************************
Function ObtainVIPREBusinessAgentData
  'Grab the VIPRE data from the text file from the A/V install path
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strViprebusinessAgt1 ,"InstallPath",strVipreBusinessAgtLoc
  Path = strVipreBusinessAgtLoc & "Definitions\DefVer.txt"

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
      
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The A/V Definition File is " & CalculatedPatternAge & " days old."
     
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  

  'Let's figure out if Real Time Scanning is enabled or not
  Dim ServiceActive
	ServiceActive = isProcessRunning(strComputer,"SBAMSvc.exe")
	
	If ( ServiceActive) Then
		OnAccessScanningEnabled = TRUE
	Else
		OnAccessScanningEnabled = FALSE
	End If

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  'Let's figure out the version number
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strViprebusinessAgt1 ,"Version",FormattedAVVersion
  output.writeline "- The version of Vipre A/V running is: " & FormattedAVVersion

End Function



' *****************************  
' Sub: CalculateAVAge
' *****************************
Sub CalculateAVAge

  If CalculatedPatternAge < OutOfDateDays Then
      ProductUpToDate = TRUE
      output.writeline "- A/V is up-to-date."
  Else
      ProductUpToDate = FALSE
      output.writeline "- A/V is not up-to-date."
  End If

End Sub


' *************************************  
' Sub: ObtainKaspersky2012AVData
' *************************************
Sub ObtainKaspersky2012AVData
  strKaspersky2012AVDatePath = Registry & "KasperskyLab\protected\AVP12\"
  objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyAVDatePath & "Data\","LastSuccessfulUpdate",AVDatVersion
  
  ' The AVDatVersion value is a UNIX timestamp - so this script needs to convert it to a human-readable format before continuing.
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
      If RegKeyExists("HKLM\Software\Microsoft\Microsoft Antimalware\Signature Updates\AVSignatureVersion") Then
        strSecurityEssentialsKeyPath = "Software\Microsoft\Microsoft Antimalware\"
      Else
        strSecurityEssentialsKeyPath = "Software\Microsoft\Windows Defender\"
      End If
  End If
  
 
  
  ' Let's grab what A/V definition version Security Essentials is using
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strSecurityEssentialsKeyPath & "Signature Updates\","AVSignatureVersion",FormattedAVVersion
  output.writeline "- The A/V version " & InstalledAV & " is running is: " & FormattedAVVersion
  
  
  ' Let's figure out if Real-Time Scanning is enabled or not
  If RegKeyExists ("HKLM\" & strSecurityEssentialsKeyPath & "Real-Time Protection\DisableRealtimeMonitoring") Then
    output.writeline "- The DisableRealTimeMonitoring registry value exists, so let's check and see how it's been configured."
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strSecurityEssentialsKeyPath & "Real-Time Protection\","DisableRealTimeMonitoring",RawOnAccessScanningEnabled
    If RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = FALSE
    Else
      OnAccessScanningEnabled = TRUE
    End If
  Else
    ' output.writeline "- The DisableRealTimeMonitoring registry value didn't exist, which means that Real-Time Scanning is enabled."
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
  FormattedPatternAge = DateAdd("s", AVDatVersion, #1/1/1970#)
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time Kasperky was updated was " & CalculatedPatternAge & " days ago."



  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

    Set oShell = CreateObject("WScript.Shell")
    Set oExec = oShell.Exec(ProgramFiles & "\Kaspersky Lab\Kaspersky Endpoint Security 8 for Windows\avp.com status FM")
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
  objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strESETKeyPath & "\info","ProductVersion",ProductVersion
  ObtainESETRegistryData
  
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled
  
  ' Let's figure out how long it's been since ESET was updated
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time " & InstalledAV & " was updated was " & CalculatedPatternAge & " days ago."
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are


End Sub





' *************************************  
' Sub: ObtainESETFSData
' *************************************
Sub ObtainESETFSData
  objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strESETKeyPath&"\info","ProductVersion",ProductVersion
  
  ' Let's append the version number to the name of the installed AV product.
  output.writeline "- The installed version of ESET File Security is: " & ProductVersion
  
  ' Let's grab the A/V Definition version
  objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strESETKeyPath&"\info","ScannerVersion",AVDatVersion
  ' That A/V Definition version is composed of two pieces of data - the build number and the date (   i.e. 6762 (20120102)   ). We need to separate them out.
  DateStart = InStr(AVDatVersion,"(") ' This determines the position of the "(" character
  RawAVDate = Mid(AVDatVersion, DateStart+1) ' This grabs everything to the right of the "(" character
  FormattedPatternAge = Left(RawAVDate,4) & "/" & Mid(RawAVDate,5,2) & "/" & Mid(RawAVDate,7,2)
  
  ' Now let's grab the '
  FormattedAVVersion = "Version: " & Left(AVDatVersion,DateStart -1) & "Date: " & FormattedPatternAge 

  'Let's get the data about ESET from WMI
  If WMINamespaceExistanceCheck("ESET")="1" Then
    Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\ESET")
    Set colItems = objWMIService.ExecQuery("SELECT StatusCode,StatusText FROM ESET_Product", "WQL", _
    wbemFlagReturnImmediately + wbemFlagForwardOnly)
  
    For Each objItem in colItems
        StatusCode = objItem.StatusCode
        StatusText = objItem.StatusText
        output.writeline "- Status Code: " & StatusCode 'Use this line for debug purposes
        output.writeline "- Status Text: " & StatusText 'Use this line for debug purposes
    Next
  
    If StatusCode = "0" Then
        OnAccessScanningEnabled = TRUE
    ElseIf ((StatusCode = "1") OR (StatusCode = "2")) Then
      InstalledAV1 = InstalledAV
      If InStr(StatusText,"Besturingssysteem is niet up-to-date") OR InStr(StatusText,"Operating system is not up") OR InStr(StatusText,"Your license will expire soon") _
      OR InStr(StatusText,"Het wordt aangeraden de computer opnieuw op te starten") OR InStr(StatusText,"De computer moet opnieuw worden opgestart") OR InStr(StatusText,"A computer restart is recommended") _
      OR InStr(StatusText,"Licentie van e-mailserverbeveiliging verloopt binnenkort") Then

        OnAccessScanningEnabled = TRUE
      Else
        OnAccessScanningEnabled = FALSE
      End If
    End If
  
  ' If the root\ESET WMI namespace doesn't exist, then let's try querying the registry
  Else
    output.writeline "- Because the root\ESET WMI namespace doesn't exist, let's query the registry."
    ObtainESETRegistryData ' Call the sub that grabs the data
    
  End If

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled
  
  ' Let's figure out how long it's been since ESET was updated
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time " & InstalledAV & " was updated was " & CalculatedPatternAge & " days ago."
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are


End Sub





' *************************************  
' Sub: ObtainESETRegistryData
' *************************************
Sub ObtainESETRegistryData
  ' Let's grab the A/V Definition version
  objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,strESETKeyPath & "\info","ScannerVersion",AVDatVersion
  ' That A/V Definition version is composed of two pieces of data - the build number and the date (   i.e. 6762 (20120102)   ). We need to separate them out.
  DateStart = InStr(AVDatVersion,"(") ' This determines the position of the "(" character
  RawAVDate = Mid(AVDatVersion, DateStart+1) ' This grabs everything to the right of the "(" character
  FormattedPatternAge = Left(RawAVDate,4) & "/" & Mid(RawAVDate,5,2) & "/" & Mid(RawAVDate,7,2)
  
  ' Now let's grab the '
  FormattedAVVersion = "Version: " & Left(AVDatVersion,DateStart -1) & "Date: " & FormattedPatternAge
  output.writeline "- " & FormattedAVVersion 


   If RegKeyExists ("HKLM\" & strESETKeyPath&"Scanners\01010100\Profiles") Then
 	  objReg.GetDWordValue HKEY_LOCAL_MACHINE,strESETKeyPath&"Scanners\01010100\Profiles","Enable", OnAccessScanningEnabled
   Else
	   ServiceActive = isProcessRunning(strComputer,"ekrn.exe")
	   If ( ServiceActive) Then
		  OnAccessScanningEnabled = TRUE
	   Else
		  OnAccessScanningEnabled = FALSE
	   End If
    End If 
End Sub 
 
 
' *************************************  
' Sub: ObtainKESServerData
' *************************************
Sub ObtainKESServerData
   strKasperskyKESServerAVDatePath = Registry & "KasperskyLab\Components\34\1103\1.0.0.0\Statistics\AVState\"
   strKasperskyKESServerAVVersionPath = Registry & "KasperskyLab\WSEE\10.1\Install\"
  
  'Let's figure out the version number
  If RegKeyExists ("HKLM\" & strKasperskyKESServerKeyPath & "ProdVersion") Then
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyKESServerKeyPath,"ProdVersion",FormattedAVVersion
    output.writeline "- The version of Kaspersky running is: " & FormattedAVVersion
  ElseIf RegKeyExists ("HKLM\" & strKasperskyKESServerAVVersionPath & "Version") Then
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVVersionPath,"Version",FormattedAVVersion
    output.writeline "- The version of Kaspersky running is: " & FormattedAVVersion
  Else
    output.writeline "- Unable to find what version of Kaspersky is running on this PC."
    FormattedAVVersion ="Unknown"
  End If   
  

   
   If RegKeyExists ("HKLM\" & strKasperskyKESServerAVDatePath & "Protection_BasesDate") Then
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath,"Protection_BasesDate",AVDatVersion
    If Len(AVDatVersion) = 0 Then 'If the registry value is Null, let's get the date from the XML file
      output.writeline "- The age of the Kaspersky A/V definitions wasn't found in the registry; checking the XML file instead."   'Re-enable this for debugging purposes, if needed
      'Grab the Kaspersky data from the text file from the A/V install path
	
      Set objFSO = CreateObject("Scripting.FileSystemObject") 

	    dim path2
	    path2 = InstallLocation & "\Kaspersky Lab\KES10SP1\Data\U1313g.xml"

	    If objFSO.FileExists(Path) Then
			 output.writeline "- Using File :" & Path
	    ElseIf objFSO.FileExists(path2) Then
			 Path = path2
			 output.writeline "- Using File :" & path2
	    Else
			 output.writeline "- ERROR - Kaspersky Ux file not found"
	    End If

      Set objFile = objFSO.OpenTextFile(Path, 1) 
      Set f = objFSO.GetFile(Path)
      Set objFile = objFSO.GetFile(Path) 
	  AVDatVersion = objFile.DateLastModified
	  output.writeline "- The file was last modified on: " & AVDatVersion
	  output.writeline "- The current date is: " & CurrentDate
	  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
      output.writeline "- The last time Kaspersky was updated was " & CalculatedPatternAge & " days ago."
	
    Else 'If the registry value isn't Null, let's use it to figure things out
    
     'output.writeline "- The age of the Kaspersky A/V defintions is in the registry."  Re-enable this for debugging purposes, if needed
      RawAVDate = LEFT(AVDatVersion,10)

    
      'Looks like KES 10 changes things, and store dates in a format that needs transposition. Older versions of KES still do though, so we'll check the 
      'version number and decide which action to take (transpose date values or not).


        If CInt(Left(FormattedAVVersion,InStr(FormattedAVVersion,".")-1)) <= 10 Then
            FormattedPatternAge = DateValue(Mid(RawAVDate,7,4) & "/" & Mid(RawAVDate,4,2) & "/" & Left(RawAVDate,2) )
        Else
            If (GetLocale() = 2057 OR GetLocale()=1031 OR GetLocale()=1043 OR GetLocale()=2067) Then  'If the device is in the UK, Germany or the Netherlands, let's transpose the day and month, so that we get things right.  
                FormattedPatternAge = DateValue(Mid(RawAVDate,1,2) & "/" & Mid(RawAVDate,4,2) & "/" & Mid(RawAVDate,7,4) )
            Else
                FormattedPatternAge = DateValue(Mid(RawAVDate,4,2) & "/" & Mid(RawAVDate,1,2) & "/" & Right(RawAVDate,4) )
            End If
        End If

      
      output.writeline "- The current date is: " & CurrentDate 
      output.writeline "- According to Kaspersky, A/V definitions were last updated on: " & FormattedPatternAge
      CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
      output.writeline "- The last time Kaspersky was updated was " & CalculatedPatternAge & " days ago."
   
    End If
  End If

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  
  ' Let's determine if Real-Time Scanning is enabled or not
  If RegKeyExists ("HKLM\" & strKasperskyKESServerAVDatePath & "Protection_BasesDate") Then
    output.writeline "- This installation of Kaspersky is managed by a KES server."
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath,"Protection_RtpState",RawOnAccessScanningEnabled
    ' Values below were sourced from https://support.kaspersky.com/13758
    If RawOnAccessScanningEnabled = 4 OR RawOnAccessScanningEnabled = 5 OR RawOnAccessScanningEnabled = 6 OR RawOnAccessScanningEnabled = 7 OR RawOnAccessScanningEnabled = 8 Then
      OnAccessScanningEnabled = TRUE
    Else
      OnAccessScanningEnabled = FALSE
    End If
  
    Else
    output.writeline "- This must be a stand-alone version of " & InstalledAV
    If InStr(InstalledAV,"SP1") Then
      strKasperskyStandAlonePath = Registry & "KasperskyLab\protected\KES10SP1"  
    ElseIf InStr(InstalledAV,"SP2") Then
      strKasperskyStandAlonePath = Registry & "KasperskyLab\protected\KES10SP2"
    ElseIf InStr(InstalledAV,"SP3") Then
      strKasperskyStandAlonePath = Registry & "KasperskyLab\protected\KES10SP3"
    End If 
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyStandAlonePath & "\settings\def\","EnableSelfProtection",RawOnAccessScanningEnabled
    If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
    ElseIf RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
    End If
  End If
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
	ServiceActive = isServiceRunning("WRSVC")
	
	If ( ProtectionEnabled = 1 AND ServiceActive ) Then
		OnAccessScanningEnabled = TRUE
	Else
		OnAccessScanningEnabled = FALSE
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

End Sub

' *************************************  
' Sub: ObtainKasperskySOSData
' *************************************
Sub ObtainKasperskySO3Sata
   strKasperskyKESServerAVDatePath = Registry & "KasperskyLab\protected\KSOS13"
   objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath & "\Data\","LastSuccessfulUpdate",AVDatVersion
  

  ' The AVDatVersion value is a UNIX timestamp - so this script needs to convert it to a human-readable format before continuing.
  output.writeline "The current date is: " & CurrentDate
  FormattedPatternAge = DateAdd("s", AVDatVersion, #1/1/1970#)
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time Kaspersky was updated was on " & FormattedPatternAge & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

   objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath,"Enabled",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
  Else
      OnAccessScanningEnabled = FALSE
  End If
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  
  
  'Let's figure out the version number
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strKasperskyKESServerAVDatePath & "environment\","ProductVersion",FormattedAVVersion
  output.writeline "- The version of Kaspersky running is: " & FormattedAVVersion

End Sub





' *************************************  
' Sub: ObtainTotalDefenseAVData
' *************************************
Sub ObtainTotalDefenseAVData
   objReg.GetStringValue HKEY_LOCAL_MACHINE,strTotalDefenseKeyPath,"AMSigsVersion",FormattedAVVersion
   objReg.GetStringValue HKEY_LOCAL_MACHINE,strTotalDefenseKeyPath,"Version",Version
   
   
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

    If  WMINamespaceExistanceCheck("FSECURE") Then 'If the F-Secure WMI namespace exists, let's use that. If not, we'll fall back to the alternate method.
        output.writeline "- The F-Secure WMI namespace is present on this device. We'll use that to gather the required information."  
        
        Set colItems2 = objWMIService.ExecQuery("SELECT RealTimeScanningEnabled,AvDefinitionsAgeInHours FROM AntiVirus2", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
  
        For Each objItem2 in colItems2
            OnAccessScanningEnabled = objItem2.RealTimeScanningEnabled
            AVDatVersion = objItem2.AvDefinitionsAgeInHours
        Next
        
        ' The Value we got back from the AvDefinitionsAgeInHours property is measured in hours, and we need it to be in days. Let's do the math.
        If (AVDatVersion <= 24) Then
            AVDatVersion = Now
        Else
            AVDatVersion=  DateAdd("h","-"& AVDatVersion,Now)
            output.writeline "- AV DatVersion is: " & AVDatversion
        End If    
        
        
        
        
    Else
      If RegKeyExists ("HKLM\Software\Wow6432Node\Data Fellows\F-Secure\FSAVCSIN\CurrentVersion") Then
          objReg.GetStringValue HKEY_LOCAL_MACHINE,"Software\Wow6432Node\Data Fellows\F-Secure\FSAVCSIN","CurrentVersion",FormattedAVVersion
      ElseIf RegKeyExists ("HKLM\" & strFSecureRegPathLoc & "CurrentVersion") Then
          objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPathLoc,"CurrentVersion",FormattedAVVersion
      ElseIf RegKeyExists ("HKLM\SOFTWARE\Wow6432Node\F-Secure\OneClient\FriendlyVersion") Then
          objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Wow6432Node\F-Secure\OneClient\","FriendlyVersion",FormattedAVVersion
      End If
     
     
      If RegKeyExists("HKLM\" & strFSecureRegPathLoc & "InstallationDirectory") Then
          objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPathLoc,"InstallationDirectory",strFSecureInstallPath
      ElseIf RegKeyExists("HKLM\" & strFSecureRegPathLoc & "Path") Then
          objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPathLoc,"Path",strFSecureInstallPath
      ElseIf RegKeyExists("HKLM\" & strFSecureRegPath000 & "InstallPath") Then
          objReg.GetStringValue HKEY_LOCAL_MACHINE,strFSecureRegPath000,"InstallPath",strFSecureInstallPath 
      End If
   
   
      ' We're going to look at the 'Last Modified' date of a Total Defense file to determine if the A/V Definitions are up-to-date or not.
      Set objFSO = CreateObject("Scripting.FileSystemObject") 
    
      If objFSO.FolderExists(strFSecureInstallPath) Then
        If objFSO.FileExists(strFSecureInstallPath & "\FS@aqua.ini") Then
          Path = strFSecureInstallPath & "\FS@aqua.ini"
        ElseIf objFSO.FileExists(strFSecureInstallPath & "\aquarius\FS@aqua.ini") Then
          Path = strFSecureInstallPath & "\aquarius\FS@aqua.ini"
        ElseIf objFSO.FileExists(strFSecureInstallPath & "\FS@hydra.ini") Then
          Path = strFSecureInstallPath & "\FS@hydra.ini"
        Else
          Path = strFSecureInstallPath & "\FS@orion.ini"
        End If
        output.writeline "- We'll use the following file to figure out how up-to-date F-Secure might be: " & Path
  
  
        Set objFile = objFSO.OpenTextFile(Path) 
        Do Until objFile.AtEndOfStream 
        strLine = objFile.ReadLine 
        If InStr(strLine, "File_set_visible_version=") Then 
          AVDatVersion= left(right(strLine,13),10)
        End If 
        If InStr(strLine, "the requested operation failed") Then 
        output.writeline "- Operation Failed" 
        End If 
        Loop 
    
        objFile.Close 
                                   
                                    
      Else
        AVDatVersion="01/01/2001"
      End If
    
    
      If isProcessRunning(".","fsav32.exe") Then
        OnAccessScanningEnabled=TRUE
      Else
        OnAccessScanningEnabled=FALSE
      End If
    
    End If
    
    output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled
    output.writeline "- The current date is: " & CurrentDate
    CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
    output.writeline "- The last time F-Secure was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."
    
    CalculateAVAge 'Call the function to determine how old the A/V Definitions are


End Sub




' *************************************  
' Sub: ObtainSEPCloudData
' *************************************
Sub ObtainSEPCloudData


   
                
   ' Let's figure out the version of Symantec Endpoint Protection Cloud that's being run             
   If RegKeyExists("HKLM\" & strSEPCloudRegpath2 & "\PRODUCTVERSION") Then
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strSEPCloudRegpath2 ,"PRODUCTVERSION",Version
   Else
    output.writeline "- Unable to find out what version of Symantec Endpoint Protection Cloud is installed on this machine."
   End If
   
   
   
   
   ' Now let's figure out the last time that Symantec updated it's AV definitions. We'll do that by finding out where the versioninfo.dat file is locate, and then looking at when it was last modified.
   If RegKeyExists("HKLM\" & strSEPCloudRegpath2 & "\SharedDefs\SDSDefs\") Then
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strSEPCloudRegpath2 & "\SharedDefs\SDSDefs","AVDEFMGR",strSEPCloudDefPath
   Else
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strSEPCloudRegpath2 & "\SharedDefs\","AVDEFMGR",strSEPCloudDefPath
   End If
 
  ' We're going to look at the 'Last Modified' date of the versioninfo.dat file to determine if the A/V Definitions are up-to-date or not.
  Set objFSO = CreateObject("Scripting.FileSystemObject") 
  If objFSO.FileExists(strSEPCloudDefPath & "\versioninfo.dat") Then
    Set objFile = objFSO.GetFile(strSEPCloudDefPath & "\versioninfo.dat") 
    AVDatVersion = objFile.DateLastModified
    output.writeline "- The file was last modified on: " & AVDatVersion
  Else
    ' output.writeline "- Unable to find the versioninfo.dat file. Let's try an alternative method to figure out if the AV definitions are up-to-date."
    ' As the strSEPCloudDefPath variable includes way too much information, lets just strip out the part we're interested in - the date value
    ' AVDatVersion = Right(Left(strSEPCloudDefPath,(Len(strSEPCloudDefPath)-4)),8)
    AVDatVersion = Left(Right(strSEPCloudDefPath,12),8)
    ' Now that we have just the date value, let's format it properly by inserting forwardslashes between the yyyy/mm/dd
    AVDatVersion = Mid(AVDatVersion,1,4) & "/" &  Mid(AVDatVersion,5,2) & "/" & Mid(AVDatVersion,7,2)
  End If
  

  If isProcessRunning(".","AVAgent.exe") Then
    OnAccessScanningEnabled=TRUE
  Else
    OnAccessScanningEnabled=FALSE
  End If
  

  FormattedAVVersion =  Version '& " - " & AVDatVersion
  output.writeline "- Symantec Endpoint Protection Cloud is running the following A/V Definition Version: " & FormattedAVVersion 

  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time SEP Cloud was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub




'ObtainPandaCloudOfficeData
' *************************************  
' Sub: ObtainPandaCloudOfficeData
' *************************************
Sub ObtainPandaCloudOfficeData

  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  strPandaAVDefinitionPath = "C:\Program Files\Panda Security\WaAgent\WalUpd\Data\Catalog"

  If objFSO.FolderExists(strPandaAVDefinitionPath) Then
	  Set objFile = objFSO.GetFile( strPandaAVDefinitionPath & "\Web_Catalog") 
	  AVDatVersion = objFile.DateLastModified
	  output.writeline "The file was last modified on: " & AVDatVersion
  ElseIf objFSO.FolderExists(strPandaAVDefinitionPath) Then
	  Set objFile = objFSO.GetFile( strPandaAVDefinitionPath & "\Last_Catalog") 
	  AVDatVersion = objFile.DateLastModified
	  output.writeline "The file was last modified on: " & AVDatVersion
  Else

		AVDatVersion="01/01/2001"

	    strPandaAVDefinitionPath = "C:\Program Files (x86)\Panda Security\WaAgent\WalUpd\Data\Catalog"
	  
	    If objFSO.FolderExists(strPandaAVDefinitionPath) Then
	  	  Set objFile = objFSO.GetFile( strPandaAVDefinitionPath & "\Web_Catalog") 
	  	  AVDatVersion = objFile.DateLastModified
	  	  output.writeline "The file was last modified on: " & AVDatVersion
	    ElseIf objFSO.FolderExists(strPandaAVDefinitionPath) Then
	  	  Set objFile = objFSO.GetFile( strPandaAVDefinitionPath & "\Last_Catalog") 
	  	  AVDatVersion = objFile.DateLastModified
	  	  output.writeline "The file was last modified on: " & AVDatVersion

	  		AVDatVersion="01/01/2001"
	    End If  
  End If  

	If isProcessRunning(".","WAHost.exe") Then
		OnAccessScanningEnabled	=TRUE
	Else
		OnAccessScanningEnabled	=FALSE
	End If

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

  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time Panda Cloud Endpoint Protection was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub



'ObtainAVG2013Data
' *************************************  
' Sub: ObtainAVG2013Data
' *************************************
Sub ObtainAVG2013Data

  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  If objFSO.FolderExists(stravg2013defpath) Then
	  Set objFile = objFSO.GetFile( stravg2013defpath & "\incavi.avm") 
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
	
	  If RegKeyExists ( "HKLM\" & stravg2013regpath32 & "\ProdType") Then
		   objReg.GetStringValue HKEY_LOCAL_MACHINE,stravg2013regpath32,"ProdType",RawAVVersion
      ElseIf RegKeyExists ( "HKLM\" & stravg2013regpath64 & "\ProdType") Then
		   objReg.GetStringValue HKEY_LOCAL_MACHINE,stravg2013regpath64,"ProdType",RawAVVersion
      End If

	FormattedAVVersion =  InstalledAV & " - " & RawAVVersion
	output.writeline "- " & FormattedAVVersion & " has been found on this device."

  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time AVG 2013 was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub



'ObtainAVG2014Data
' *************************************  
' Sub: ObtainAVG2014Data
' *************************************
Sub ObtainAVG2014Data

  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  If objFSO.FolderExists(stravg2014defpath) Then
    Set objFile = objFSO.GetFile(stravg2014defpath & "\incavi.avm")
    AVDatVersion = objFile.DateLastModified
  ElseIf objFSO.FolderExists(stravg2016defpath) Then
	 Set objFile = objFSO.GetFile(stravg2016defpath & "\incavi.avm")   
	 AVDatVersion = objFile.DateLastModified
  Else
		AVDatVersion="01/01/2001"
	  output.writeline "- The AV Definition file was not found. Odd."
  End If  

	If (isProcessRunning(".","avgwdsvc.exe") OR (isProcessRunning(".","avgwdsvca.exe"))) Then
		OnAccessScanningEnabled	=TRUE
	Else
		OnAccessScanningEnabled	=FALSE
	End If
	
	  If RegKeyExists ( "HKLM\" & stravg2014regpath32 & "\ProdType") Then
		  objReg.GetStringValue HKEY_LOCAL_MACHINE,stravg2014regpath32,"ProdType",RawAVVersion
    ElseIf RegKeyExists ( "HKLM\" & stravg2014regpath64 & "\ProdType") Then
		  objReg.GetStringValue HKEY_LOCAL_MACHINE,stravg2014regpath64,"ProdType",RawAVVersion
    ElseIf RegKeyExists ( "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AVG\DisplayVersion") Then
		  objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AVG","DisplayVersion",RawAVVersion
    End If

	FormattedAVVersion =  InstalledAV & " - " & RawAVVersion

  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time " & InstalledAV & " was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub



' *************************************  
' Sub: ObtainAvastData
' *************************************
Sub ObtainAvastData
	
  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  If objFSO.FolderExists(strAvastInstallPath & "\Defs") Then
	  Set objFile = objFSO.GetFile( StrAvastInstallPath & "\defs\aswdefs.ini") 
	  AVDatVersion = objFile.DateLastModified
  Else
		AVDatVersion="01/01/2001"
	    output.writeline "the AV Definition file is not found"
  End If  
  
	If isProcessRunning(".","avastsvc.exe") Then
		OnAccessScanningEnabled	=TRUE
	Else
		OnAccessScanningEnabled	=FALSE
	End If

	If RegKeyExists ( "HKLM\" & strAvastRegPath32 & "\Version"	  ) Then
	 objReg.GetStringValue HKEY_LOCAL_MACHINE,strAvastRegPath32,"Version",FormattedAVVersion
	ElseIf RegKeyExists ( "HKLM\" & strAvastRegPath64 & "\Version") Then
	 objReg.GetStringValue HKEY_LOCAL_MACHINE,strAvastRegPath64,"Version",FormattedAVVersion
  End If
	
  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time Avast! was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  	
  
  
End Sub



' *************************************  
' Sub: ObtainMalwarebytesCorporate
' *************************************
Sub ObtainMalwarebytesCorporate
   objReg.GetStringValue HKEY_LOCAL_MACHINE,strMalwareBytesRegPath64 & "\","dbversion",AVDatVersion
  

  ' The AVDatVersion value needs to be compared against the current date, so that we can figure out if Malwarebytes is out of date or not.
  output.writeline "- The current date is: " & CurrentDate
    ' Because the DB version string ends with a minute value (which we don't need), lets strip that off
  AVDatVersion = Left(AVDatVersion,11)
  ' Because the DB version string starts with the letter "v", let's strip that off
  AVDatVersion = Right(AVDatVersion,10)
  ' Now let's replace the periods between the year/month/day with forwardslashes
  FormattedPatternAge = Replace(AVDatVersion,".","/") 
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time Malwarebytes Corporate Edition was updated was on " & FormattedPatternAge & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

  ' Note that this section is a bit hypothetical, as I wasn't able to confirm if this is the right registry key
   objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strMalwareBytesRegPath64 & "\","alwaysscanfiles",RawOnAccessScanningEnabled
  If RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
  Else
      OnAccessScanningEnabled = FALSE
  End If                            
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  
  
  'Let's figure out the version number
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strMalwareBytesRegPath64 & "\","programversion",FormattedAVVersion
  output.writeline "- The version of Malwarebytes Corporate Edition running is: " & FormattedAVVersion

End Sub




' *************************************  
' Sub: ObtainTMMSA
' *************************************
Sub ObtainTMMSA
    objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strTMMSARegPath,"PatternStringFormatted",AVDatVersion 'Grab the AV version that TMMSA is currently running
    objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strTMMSARegPath,"Version",FormattedAVVersion 'Grab the software/application version of TMMSA
  
    
  
    
    'Let's hardcode "Real Time Scanning" to be turned on, as this AV product doesn't really have that option
      OnAccessScanningEnabled = TRUE
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set recentFile = Nothing
    For Each objFile in objFSO.GetFolder("C:\Program Files\Trend Micro\Messaging Security Agent\engine\vsapi\latest\pattern").Files
      If instr(objFile.name, "lpt$vpn") = 1 Then  
        If (recentFile is Nothing) Then
          Set recentFile = objFile
        ElseIf (objFile.DateLastModified > recentFile.DateLastModified) Then
          Set recentFile = objFile
        End If
      End If  
    Next
    

    output.writeline "- The formatted date of the A/V Definition File is: " & recentFile.DateLastModified
    

  ' The AVDatVersion value needs to be compared against the current date, so that we can figure out if TMMSA is out of date or not.
  output.writeline "- The current date is: " & CurrentDate

  FormattedPatternAge = recentFile.DateLastModified
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time Trend Micro Messaging Security Agent was updated was on " & FormattedPatternAge & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  

End Sub




  
' *************************************  
' Sub: ObtainMcAfeeEndpointSecurity
' *************************************
Sub ObtainMcAfeeEndpointSecurity

    objReg.GetStringValue HKEY_LOCAL_MACHINE,mcAfeeEndpointSecurityDefDatePath,"CasperContentTS",AVDatVersion
    objReg.GetStringValue HKEY_LOCAL_MACHINE,mcAfeeEndPointSecurityVersion,"BuildNumber",mcAfeeEndpointSecurityBuildNum
    objReg.GetStringValue HKEY_LOCAL_MACHINE,mcAfeeEndPointSecurityVersion,"ProductVersion",mcAfeeEndpointSecurityVerNum
    FormattedAVVersion = mcAfeeEndpointSecurityVerNum & "." & mcAfeeEndpointSecurityBuildNum
    mcAfeeEndpointSecurityDefDate = AVDatVersion
  
	output.writeline "McAfee Endpoint Protection Version " & mcAfeeEndpointSecurityVerNum & "." & mcAfeeEndpointSecurityBuildNum
  
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

  If isProcessRunning(".","mcshield.exe") Then
    OnAccessScanningEnabled=TRUE
  Else
    OnAccessScanningEnabled=FALSE
  End If

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled
  output.writeline "- The current date is: " & CurrentDate
  
   
  
  If right(left(AVDatVersion,2),1) = "/" Then
	  If right(left(AVDatVersion,4),1) = "/" Then
	   '1/1/...
		  FormattedPatternAge = right(AVDatVersion,4) & "/" & left(AVDatVersion,1) & "/" & right(left(AVDatVersion,3),1)
	  Else
		  '1/10/
		  FormattedPatternAge = right(AVDatVersion,4) & "/" & left(AVDatVersion,1) & "/" & right(left(AVDatVersion,4),2)
	  End If

  Else
	  If right(left(AVDatVersion,5),1) = "/" Then
		  '10/1/...
		  FormattedPatternAge = right(AVDatVersion,4) & "/" & left(AVDatVersion,2) & "/" & right(left(AVDatVersion,4),1)
	  Else
		  '10/10/..
		  FormattedPatternAge = right(AVDatVersion,4) & "/" & left(AVDatVersion,2) & "/" & right(left(AVDatVersion,5),2)
	  End If
	End If
  
  
  
  
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  output.writeline "- The last time McAfee Endpoint Security was updated was on " & FormattedPatternAge & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  	
End Sub



' *************************************  
' Sub: ObtainMcAfeeEndpointSecurity101
' *************************************
Sub ObtainMcAfeeEndpointSecurity101

   objReg.GetStringValue HKEY_LOCAL_MACHINE,mcAfeeEndPointSecurityVersion ,"BuildNumber",mcAfeeEndpointSecurityBuildNum
   objReg.GetStringValue HKEY_LOCAL_MACHINE,mcAfeeEndPointSecurityVersion ,"ProductVersion",mcAfeeEndpointSecurityVerNum
	FormattedAVVersion = mcAfeeEndpointSecurityVerNum & "." & mcAfeeEndpointSecurityBuildNum
  
	Dim filename,fso,f,Linevalue,bfound,goodline

    If objFSO.FileExists (ProgramData & "\mcafee\agent\updatehistory.ini") Then
	   filename = ProgramData & "\mcafee\agent\updatehistory.ini"
       ' Enable for Debug purposes output.writeline "- The updatehistory.ini file has been found under %ProgramData%\mcafee\agent\"
    Else
        ' Enable for Debug purposes  output.writeline "- The updatehistory.ini file was not found under %ProgramData%\mcafee\agent\ - let's try looking under the Common Framework path."
        If objFSO.FileExists (ProgramData & "\mcafee\common framework\updatehistory.ini") Then
            filename = ProgramData & "\mcafee\common framework\updatehistory.ini"
            ' Enable for Debug purposes output.writeline "- The updatehistory.ini file has been found under %ProgramData%\mcafee\common framework\"
        Else
            output.writeline "- Unable to fine the updatehistory.ini file. That's.....not.....good."
        End If 
    End If
    
	goodline=""
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(filename)

	Do Until f.AtEndOfStream
		Linevalue = f.ReadLine
		If(InStr(1,Linevalue,"CatalogVersion=")) Then
			If(bfound <> "True") Then
				bfound = "True"
				goodline = Linevalue
			End If
		End If
	Loop

	goodline = left(right(Left(goodline,23),8),4) & "-" & left(right(Left(goodline,23),4),2) & "-" & right(Left(goodline,23),2)

	f.Close
	FormattedPatternAge = goodline
	AVDatVersion = goodline
		

  If isProcessRunning(".","mcshield.exe") Then
    OnAccessScanningEnabled=TRUE
  Else
    OnAccessScanningEnabled=FALSE
  End If

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled
  output.writeline "- The current date is: " & CurrentDate
  
  
  CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  output.writeline "- The last time McAfee Endpoint Security was updated was on " & FormattedPatternAge & ", which was " & CalculatedPatternAge & " days ago."

  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  	
End Sub




Function isProcessRunning(byval strComputer,byval strProcessName)

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
   Set colItems = objWMIService.ExecQuery("SELECT * from AntiVirusProduct", "WQL", wbemFlagReturnImmediately)

    
   
  If (colItems.Count = "0" OR IsNull(colItems)) Then
     output.writeline "- There are no enties in the SecurityCenter2 namespace."
     ' Because there's nothing in the SecurityCenter2 namespace, the last check we need to perform is to see if Windows Defender is installed
     If RegKeyExists ("HKLM\" & strWindowsDefenderPath & "ProductStatus") Then
      InstalledAV = "Windows Defender"
      output.writeline "- " & InstalledAV & " has been detected."
     End If  
     Exit Sub 'Because there was nothing in the SecurityCenter2 namespace, we should continue querying for other A/V products
   ElseIf colItems.Count = "1" Then
      output.writeline "- There is 1 entry in the SecurityCenter2 namespace."
   Else
      output.writeline "- There are " & colItems.Count & " entries in the SecurityCenter2 namespace."
   End If
   


    If colItems.Count > 0 Then
      For Each objItem In colItems
      If  objItem.displayName = "Windows Defender" Then
        InstalledAV = objItem.displayName
      End If
      
      If ((objItem.displayName <> "AVG update module") AND (objItem.displayName <> "Windows Defender")) Then
        
        onAccessScanningEnabled = "FALSE"
        ProductUpToDate = "FALSE"
  		
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
  			 
         
        If CLng("&H" & HexScannerState) = 16 Then
  		    OnAccessScanningEnabled = TRUE
  			Else
  		        OnAccessScanningEnabled = FALSE
  			End If
  			output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled  
  			
  			' The last of the 3 HEX values tells us if the A/V definition files are up-to-date or not
  			HexAVDefState =  Right(HexProductState,2)
  			' Keep this line around for debugging purposes. output.writeline CLng("&H" & HexAVDefState)
  			If CLng("&H" & HexAVDefState) = 0 Then
  			   ProductUpToDate = TRUE
  			Else
  			    ProductUpToDate = FALSE
  			End If
  			output.writeline "- Are the AV Definitions up-to-date? " & ProductUpToDate
        
            CalculateAVAge 'Call the function to determine how old the A/V Definitions are 

  		End If
        If  InStr(objItem.displayName,"FireEye Endpoint Security") Then Exit For
    Next
    End If    
      
    
      
   
  		

End Sub


' *************************************  
' Sub: ObtainTrendMicroDeepSecurity
' *************************************
sub ObtainTrendMicroDeepSecurity
  ' Is the scanning service running?
  If isProcessRunning(".","coreServiceShell.exe") Then
    OnAccessScanningEnabled=TRUE
  Else
    OnAccessScanningEnabled=FALSE
  End If
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled


  'Let's get the version of Deep Security that's been installed
  objReg.GetStringValue HKEY_LOCAL_MACHINE,strTMDSARegPath ,"InstalledVersion",FormattedAVVersion

                
  'Let's figure out when the Deep Security Agent was last updated
  Set fileSystem = CreateObject("Scripting.FileSystemObject") 
  If fileSystem.FolderExists("C:\Users\All Users\Trend Micro\Deep Security Agent\iaurepo\packages") Then
    Set folder = fileSystem.GetFolder("C:\Users\All Users\Trend Micro\Deep Security Agent\iaurepo\packages") 
    For Each file In folder.Files         
      If file.DateLastModified > newestfile Then         
        newestfile = file.DateLastModified
      End If
    Next
  Else
    output.writeline "- The file C:\Users\All Users\Trend Micro\Deep Security Agent\iaurepo\packages was not found."
  End If
  AVDatVersion = newestfile 
  CalculatedPatternAge = DateDiff("d",newestfile,CurrentDate)
  output.writeline "- The last time Trend Micro Deep Security was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."
  
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
 
End Sub



' *************************************  
' Sub: ObtainMcAfeeMove
' *************************************
sub ObtainMcAfeeMove
  ' Is the scanning service running?
  If isProcessRunning(".","masvc.exe") Then
    OnAccessScanningEnabled=TRUE
  Else
    OnAccessScanningEnabled=FALSE
  End If
  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled


  'Let's get the version of the McAfee Move Agent that's been installed
  If RegKeyExists ("HKLM\" & Registry & "McAfee\Agent\AgentVersion") Then
    objReg.GetStringValue HKEY_LOCAL_MACHINE, Registry & "McAfee\Agent\","AgentVersion",FormattedAVVersion
    output.writeline "- Found the Version key for McAfee Move: " & FormattedAVVersion
  Else
    output.writeline "- Didn't find the key"
  End If

                
  'Let's figure out when the McAfee Move Agent was last updated
  Set fileSystem = CreateObject("Scripting.FileSystemObject") 
  Set objFile = fileSystem.GetFile("C:\Users\All Users\McAfee\Agent\UpdateHistory.ini")
  AVDatVersion = objFile.DateLastModified 
 
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time the McAfee Move Agent was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."
  
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
 
End Sub



' *************************************  
' Sub: ObtainNormanEndpointProtection
' *************************************
Sub ObtainNormanEndpointProtection
  ' Is the scanning service running?
   
  Set fileSystem = CreateObject("Scripting.FileSystemObject")
  boolNormanversion9 = fileSystem.FileExists(strNormanrootpath & "\Nvc\bin\nvcoas.exe")
 
  On Error Resume Next
  Err.Clear
  If boolNormanversion9 Then
    Set file = fileSystem.OpenTextFile(strNormanrootpath & "\Nvc\bin\nvcoas.exe", 8, False) ' 8 = Append
  Else
    Set file = fileSystem.OpenTextFile(strNormanrootpath & "\Nvc\bin\gzfltum.dll", 8, False) ' 8 = Append
  End If
 
  If Err.Number = 0 Then
    file.Close
    OnAccessScanningEnabled=FALSE
  ElseIf Err.Number <> 70 Then ' 70 = Permission denied (in-use)
    OnAccessScanningEnabled=FALSE
  Else
    OnAccessScanningEnabled=TRUE
  End If
  On Error Goto 0

  output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

  'Let's get the version of Norman Endpoint Protection that's been installed
  FormattedAVVersion = "unknown"
  objReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", arrSubKeys
  For Each subkey In arrSubKeys
    objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & subkey,"DisplayName",RegstrValue
    If InStr(1, RegstrValue, "Norman Endpoint Protection", vbTextCompare) = 1 Then
      objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & subkey,"DisplayVersion",FormattedAVVersion
      Exit For
    End If
  Next
  ' On 64-bit systems, the product version registry key might be stored in the 32-bit branch
  If FormattedAVVersion = "unknown" And RegKeyExists ("HKLM\SOFTWARE\Wow6432Node\") Then
    objReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", arrSubKeys
    For Each subkey In arrSubKeys
      objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & subkey,"DisplayName",RegstrValue
      If InStr(1, RegstrValue, "Norman Endpoint Protection", vbTextCompare) = 1 Then
        objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & subkey,"DisplayVersion",FormattedAVVersion
        Exit For
      End If
    Next
  End If
  output.writeline "- Found the Version of Norman Endpoint Protection: " & FormattedAVVersion

  'Let's figure out when Norman Endpoint Protection was last updated

  On Error Resume Next
  If boolNormanversion9 Then
     Set objFile = fileSystem.GetFile(strNormanrootpath & "\Nse\Bin\Nvcincr.def")
  Else
    Set objFile = fileSystem.GetFile(strNormanrootpath & "\Nse\def00\descr.dta")
  End If

  If Err.Number <> 0 Then
    AVDatVersion="01/01/2001"
    output.writeline "- The AV Definition file is not found"
  Else
    AVDatVersion = objFile.DateLastModified
    output.writeline "- The last time Norman Endpoint Protection was updated was on " & AVDatVersion
  End If
  On Error Goto 0
 
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are

End Sub


' *************************************  
' Sub: ObtainAVGBusinessSecurity
' *************************************
Sub ObtainAVGBusinessSecurity
  ' Let's figure out what version of AVG Business Security has been installed
  If RegKeyExists("HKLM\SOFTWARE\WOW6432Node\AVG\Antivirus\Version") Then
    objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\WOW6432Node\AVG\Antivirus\","Version",FormattedAVVersion
    InstalledAV = InstalledAV & " " & FormattedAVVersion
  End If  
  
  
  ' Let's figure out where AVG has been installed, so that we can interrogate a few INI files for information
  If RegKeyExists("HKLM\" & Registry & "AVG\Antivirus\DataFolder") Then
    objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\WOW6432Node\AVG\Antivirus\","DataFolder",AVGBusSecDataFolder
     output.writeline "- The Data folder for " & InstalledAV & " is " &  AVGBusSecDataFolder
  Else
    output.writeline "- The Data Folder for AVG isn't listed on this device. That's......not.....very......good."
  End If
  ' Now that we know the path to those INI files, let's find out if Real-Time Scanning is turned on or off
  Set objFSO = CreateObject("Scripting.FileSystemObject") 
  If objFSO.FileExists(AVGBusSecDataFolder & "\FileSystemShield.ini") Then
    Set objFile = objFSO.OpenTextFile(AVGBusSecDataFolder & "\FileSystemShield.ini", 1, False, -1)
    Dim arrIniFileLines() 
    i = 0 
    Do Until objFile.AtEndOfStream 
      Redim Preserve arrIniFileLines(i) 
      arrIniFileLines(i) = objFile.ReadLine 
      i = i + 1 
    Loop 
    objFile.Close
  Else
    output.writeline "The FileSystemShield.ini file does not exist. That's.......not.......very.....good."
  End If
 
 
  'Let's see if real-time scanning is enabled or disable through the AVG GUI 
  For Each strLine in arrIniFileLines
    If InStr(strLine,"ProviderEnabled") Then 'Check if it's the right line before we look for the value - it'll start with the word 'ProviderEnabled'
      If InStr(strLine,"ProviderEnabled=1") Then
        ProviderRealTimeScanningEnabled = TRUE
      ElseIf InStr(strLine,"ProviderEnabled=0") Then
        ProviderRealTimeScanningEnabled = FALSE
      Else
        output.writeline "- Here's what we saw in the log file: " & strLine
        ProviderRealTimeScanningEnabled = "UNKNOWN"    
      End If
      Exit For
    End If
  Next
    
  output.writeline "- Has real-time scanning been enabled through the AVG GUI? " & ProviderRealTimeScanningEnabled
  
  'Let's see if real-time scanning is enabled or disable through the AVG system tray menu
  UserRealTimeScanningDisabled = FALSE 'Let's start off by assuming that it hasn't been disabled, unless we find evidence to the contrary 
  For Each strLine in arrIniFileLines
    If InStr (strLine,"TemporaryDisabled") Then 'Check if it's the right line before we look for the value - it'll start with the word 'TemporaryDisabled'
      If InStr(strLine,"TemporaryDisabled=0") Then
        UserRealTimeScanningDisabled = FALSE
      ElseIf InStr(strLine,"TemporaryDisabled=1") Then
        UserRealTimeScanningDisabled = TRUE 
      End If
      Exit For
    End If
  Next  
  output.writeline "- Has real-time scanning been disabled through the system tray icon? " & UserRealTimeScanningDisabled
  
  
  If ((ProviderRealTimeScanningEnabled = TRUE) And (UserRealTimeScanningDisabled = FALSE)) Then
    OnAccessScanningEnabled = TRUE
  Else
    OnAccessScanningEnabled = FALSE
  End If
    output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled    
  
  
  ' Let's figure out when AVG Business Security last updated itself
  Set objFSO = CreateObject("Scripting.FileSystemObject") 
  Set objFile = objFSO.OpenTextFile(ProgramData & "\AVG\Persistent Data\Antivirus\Logs\update.log", 1) 

  Dim arrFileLines() 
  i = 0 
  Do Until objFile.AtEndOfStream 
  Redim Preserve arrFileLines(i) 
  arrFileLines(i) = objFile.ReadLine 
  i = i + 1 
  Loop 
  objFile.Close 
 
  'Then you can iterate it like this 
 
  For Each strLine in arrFileLines
    If InStr(strLine,"vps") Then
      LastUpdateDate = strLine
    End If
  Next
  
  ' Now that we have the line from the log file, let's grab the date at the very end
  AVDatVersion = Replace(Mid(LastUpdateDate,7,2) & "/" & Mid(LastUpdateDate,10,2) & "/" & Mid(LastUpdateDate,2,4),"-","/")
  output.writeline "- The last time AVG Business Security was updated was on " & AVDatVersion
 
  CalculatedPatternAge = DateDiff("d",DateValue(AVDatVersion),CurrentDate)
  
  CalculateAVAge 'Call the function to determine how old the A/V Definitions are
  

  
End Sub

' *****************************  
' Sub: FortiClient
' *****************************
Sub ObtainFortiClient
    ' Let's figure out when FortiClient last updated its AV definitions
    ' First let's figure out where FortiClient has been installed
    If RegKeyExists ("HKLM\SOFTWARE\Fortinet\FortiClient\INSTALLDIR") Then
        objReg.GetSTRINGValue HKEY_LOCAL_MACHINE,"SOFTWARE\Fortinet\FortiClient","INSTALLDIR",FortiClientInstallPath
    Else
        output.writeline "- Unable to find where FortiClient has been installed. Assuming it's under C:\Program Files\Fortinet\FortiClient\"
        FortiClientInstallPath = "C:\Program Files\Fortinet\FortiClient\"
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists (FortiClientInstallPath & "vir_sig\vir_high") Then
        Set objFile = objFSO.GetFile(FortiClientInstallPath & "vir_sig\vir_high")
        AVDatVersion = objFile.DateLastModified
        output.writeline "- The file was last modified on: " & AVDatVersion
	    output.writeline "- The current date is: " & CurrentDate
	    CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
        output.writeline "- The last time FortiClient was updated was " & CalculatedPatternAge & " days ago."
        CalculateAVAge 'Call the function to determine how old the A/V Definitions are
    Else
        output.writeline "- Unable to find the vir_high file. This is not a good thing."
    End If
    
    ' Now let's figure out whether or not Real-Time Scanning is enabled
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strFortiClientPath,"enabled",RawOnAccessScanningEnabled
    If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
    ElseIf RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
    End If
    
    output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

    'Let's figure out the version number
    Set colItems = objWMIService.ExecQuery("Select Version from Win32_Product where Name like '%FortiClient%'")
    For Each objApp in colItems
        FormattedAVVersion = objApp.Version
    Next
    output.writeline "- The version of FortiClient installed on this machine is: " & FormattedAVVersion



End Sub


' *************************************  
' Sub: ObtainPandaAdaptiveDefenceData
' *************************************
Sub ObtainPandaAdaptiveDefenceData

  Set objFSO = CreateObject("Scripting.FileSystemObject") 

  'Let's figure out where Panda stores its NanoRepository.bin file
  If objFSO.FolderExists(ProgramData & "\Panda Security\Panda Security Protection") Then
    strPandaAVDefinitionPath = ProgramData & "\Panda Security\Panda Security Protection"
  ElseIf objFSO.FolderExists(ProgramData & "\Panda Security\Security Protection") Then
    strPandaAVDefinitionPath = ProgramData & "\Panda Security\Security Protection"
  Else
    output.writeline "- Unable to determine the location of the NanoRepository.bin file."
  End If
  
  ' Now that we know where that bin file is located, let's see when it was last modified	  
  Set objFile = objFSO.GetFile( strPandaAVDefinitionPath & "\NanoRepository.bin") 
	  AVDatVersion = objFile.DateLastModified

	If isProcessRunning(".","PSANHost.exe") Then
		OnAccessScanningEnabled	=TRUE
	Else
		OnAccessScanningEnabled	=FALSE
	End If

	If RegKeyExists ( "HKLM\" & strPandaCloudEPPath32 & "\Normal") Then
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strPandaCloudEPPath32 ,"Normal",RawAVVersion 
  ElseIf RegKeyExists ( "HKLM\" & strPandaCloudEPPath64 & "\Normal") Then
      objReg.GetStringValue HKEY_LOCAL_MACHINE,strPandaCloudEPPath64 ,"Normal",RawAVVersion 
  Else
    	RawAVVersion = "Unknown Version"
	End If

    FormattedAVVersion =  InstalledAV & " - " & RawAVVersion
	output.writeline "- " & InstalledAV & " has been found on this device."
	output.writeline "- Panda Cloud Endpoint Protection is running the following A/V Definition Version: " & FormattedAVVersion 

  output.writeline "- The current date is: " & CurrentDate
  CalculatedPatternAge = DateDiff("d",AVDatVersion,CurrentDate)
  output.writeline "- The last time Panda Cloud Endpoint Protection was updated was on " & AVDatVersion & ", which was " & CalculatedPatternAge & " days ago."
  output.writeline "- Is Real-Time Scanning Enabled? " & OnAccessScanningEnabled

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
' Sub: ObtainCiscoAMPData
' *************************************
Sub ObtainCiscoAMPData
    onAccessScanningEnabled = True 'Real time scanning is controlled through the AMP Management Console so this option is N/A
    ProductUpToDate = True 'There aren't pattern file versions that can be monitored from the machine itself - only from the Management Console so this option is N/A
    FormattedAVVersion = "See Cisco AMP Cloud Management Console at https://console.amp.cisco.com/"
End Sub


' *************************************  
' Sub: ObtainPaloAltoTrapsAVData
' *************************************
Sub ObtainPaloAltoTrapsAVData
    ' Let's determine what version of Palo Alto Traps has been installed
    objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Cyvera\Client\","Product Version",FormattedAVVersion
    
    ' Let's determine whether or not Real-Time Scanning is enabled
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,"SOFTWARE\Palo Alto Networks\Traps\","ProtectionStatus",RawProtectionStatus
    
    If RawProtectionStatus=3 Then
        onAccessScanningEnabled = True
    Else
        onAccessScanningEnabled = False
    End If
    
    ' Let's figure out when Palo Alto last updated itself
    ProductUpToDate = True
End Sub


  
' *************************************  
' Sub: ObtainSentinelOneData
' *************************************
Sub ObtainSentinelOneData
    'SentinelOne is a different type of AV product - it doesn't have the traditional concept of AV definition updates.
    ' As a result, we're going to run the SentinelCtl executable to find out how the product is doing.
    
    ' First, we need to find the location of the SentinelCtl executable, as it's stored in a folder whose name changes with each version of SentinelOne.
    ' As there could be many folders (for example, if multiple versions of SentinelOne have been installed) we need to find the newest folder
    Set objFolder = objFSO.GetFolder(ProgramFiles64 & "\SentinelOne")
    Set objSubFolders = objFolder.SubFolders
    sNewestFolder = NULL
    For Each objSubFolder In objSubFolders
        If IsNull(sNewestFolder) Then
            sNewestFolder = objSubFolder.Path
            dPrevDate = objSubFolder.DateLastModified
        ElseIf dPrevDate < objSubFolder.DateLastModified Then
            sNewestFolder = objSubFolder.Path
        End If
    Next
    
    If objFSO.FileExists (sNewestFolder & "\SentinelCtl.exe") Then
        ' output.writeline "- " & "The SentinelCtl executable is located under " & sNewestFolder & "\"
    Else
        output.writeline "- The SentinelCtl executable was not found under the most recently modified folder (" & sNewestFolder & "). That is not good. Exiting, as we cannot determine the status of SentinelOne."
        wscript.quit(1)
    End If

    
    ' Now that we know where to go to find the SentinelCtl executable, let's run it and parse the output.
    Set oExec = WshShell.Exec(sNewestFolder & "\SentinelCtl.exe status")
    OnAccessScanningEnabled = FALSE
    Do 'Because the status command outputs many lines of text, we need to use a Do/Loop command to grab all of it
        sLine = oExec.StdOut.ReadLine()
        If InStr(sLine, "SentinelMonitor is loaded") <> 0 Then
            OnAccessScanningEnabled = TRUE
            Exit Do
        End If
    Loop While Not oExec.Stdout.atEndOfStream  
                

    
    output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled 
    

    ' Now let's determine if this S1 install is one that was purchased through SolarWinds, or another vendor.
    set oExec = WshShell.Exec(sNewestFolder & "\SentinelCtl.exe config server.mgmtServer")
    sLine = oExec.StdOut.ReadLine
                
    If InStr(sLine, "swprd") <> 0 Then
        InstalledAV = "SolarWinds EDR"
        output.writeline "- This is an installation of SolarWinds EDR."
    Else
        output.writeline "- This is an installation of SentinelOne that is not managed through SolarWinds."
    End If
     
     ' Note that because we have no way of determining if SentinelOne is up-to-date or not, the ProductUpToDate value is hardcoded to be TRUE
     ProductUpToDate = "TRUE"
     
     ' Let's call the GetAgentStatusJSON object to determine what version of S1 is installed on this device 
    Set S1HelperObj = CreateObject("SentinelHelper.1")
    S1AgentStatus = s1HelperObj.GetAgentStatusJSON()
    Set S1HelperObj = nothing
    
    If InStr(S1AgentStatus, "agent-version") <> 0 Then
        IndexOfAgentVersion = InStr(S1AgentStatus, "agent-version") + 16
        LenOfVersion = InStr(IndexOfAgentVersion, S1AgentStatus, """") - IndexOfAgentVersion
        FormattedAVVersion = Mid(S1AgentStatus, IndexOfAgentVersion, lenOfVersion)
    Else
        FormattedAVVersion = Replace (Right(sNewestFolder,9)," ","")
    End If

End Sub   
     

' *************************************  
' Sub: ObtainSophosVirtualAVData
' *************************************
Sub ObtainSophosVirtualAVData
    ' Let's determine whether or not Real-Time Scanning is enabled
    objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strSophosVirtualAVKeyPath & "SGVM Scanning Service\On-Access Policy\","On-Access Protection",RawOnAccessScanningEnabled
    If RawOnAccessScanningEnabled = 0 Then
      OnAccessScanningEnabled = FALSE
    ElseIf RawOnAccessScanningEnabled = 1 Then
      OnAccessScanningEnabled = TRUE
    End If
    
    output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled
    
    
    ' Let's determine what version of Sophos for Virtual Environments has been installed
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strSophosVirtualAVKeyPath & "SGVM Management Service\Installed Components\Sophos GVM Scanning Service\","Version",FormattedAVVersion
    
    ' Let's figure out when the A/V Definitions were last updated
    objReg.GetQWORDValue HKEY_LOCAL_MACHINE,strSophosVirtualAVKeyPath & "SGVM Scanning Service\On-Access Policy\","Virus Data Age",AVDatVersion
    ' The AVDatVersion value is a UNIX timestamp - so this script needs to convert it to a human-readable format before continuing.
    FormattedPatternAge = DateAdd("s", AVDatVersion, #1/1/1970#)
    CalculatedPatternAge = DateDiff("d",FormattedPatternAge,CurrentDate)
    output.writeline "- The last time Sophos for Virtual Environments was updated was " & CalculatedPatternAge & " days ago."
    
    

    ProductUpToDate = True 'There aren't pattern file versions that can be monitored from the machine itself - only from the Management Console so this option is N/A

End Sub  


' *************************************  
' Sub: ObtainBESTData
' *************************************
Sub ObtainBESTData
    
    ' Let's see if the eps.rmm.exe file is in C:\Temp
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists("C:\Temp\eps.rmm.exe") Then
        output.writeline "- The Bitdefender RMM SDK is present on this device."
             
    ElseIf DownloadFile("http://download.bitdefender.com/SMB/RMM/Tools/Win/1.0.0.88/x64/eps.rmm.exe", "C:\Temp\") = True Then
        output.writeline "- The eps.rmm.exe file was not found in C:\temp, but it's been successfully downloaded."
    
    Else
        output.writeline "- The Bitdefender RMM SDK is not present on this device, and could not be downloaded."
        output.writeline "- Please download the Bitdefender RMM SDK from http://download.bitdefender.com/SMB/RMM/Tools/Win/ and place it in the C:\Temp folder."
        output.writeline "- Exiting the script."
        Wscript.quit(0)
    End If
    
    
    ' Now that we know where to go to find the eps.rmm executable, let's use it to determine what version of Bitdefender is installed on the device.
    Set oExec = WshShell.Exec("C:\Temp\eps.rmm.exe -getProductVersion")
    sLine = oExec.StdOut.ReadLine
                
    If InStr(sLine, ".") <> 0 Then
        FormattedAVVersion = sLine
    Else
        FormattedAVVersion = "Unknown"
    End If
        
    output.writeline "- The version of Bitdefender running on this machine is: " & FormattedAVVersion
        
    ' Now let's use the eps.rmm.exe to determine whether or not Bitdefender is up-to-date.
    Set oExec = WshShell.Exec("C:\Temp\eps.rmm.exe -isUpToDate")
    sLine = oExec.StdOut.ReadLine
    
    output.writeline "- Returned value: " & sLine
    output.writeline "- If 1, then Up-To-Date; if 0, BEST is not up-to-date."
               
    If sLine="1" Then
        ProductUpToDate = True
    Else
        ProductUpToDate = False
    End If
        
    output.writeline "- Is Bitdefender up-to-date? " & ProductUpToDate
        
    ' Now let's use the eps.rmm.exe to determine whether or not Real-Time Scanning is enabled.
    Set oExec = WshShell.Exec("C:\Temp\eps.rmm.exe -getFeatureStatus")
    sLine = oExec.StdOut.ReadLine
                
    If InStr(sLine, "AntivirusScan=1") <> 0 Then
        OnAccessScanningEnabled = True
    Else
        OnAccessScanningEnabled = False
    End If
        
        output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

End Sub  


' *************************************  
' Sub: ObtainCarbonBlackData
' *************************************
Sub ObtainCarbonBlackData
    
    ' Let's see if the repcli.exe tool is in %ProgramFiles%\Confer
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(ProgramFiles64 & "\Confer\RepCLI.exe") Then
        ' output.writeline "- Successfully located the RepCLI.exe tool."   
    Else
        output.writeline "- Unable to locate the RepCLI.exe tool."
        output.writeline "- Exiting the script."
        Wscript.quit(0)
    End If
    
    
    ' Now that we know where to go to find the repcli executable, let's use it to determine what version of Carbon Black is installed on the device.
    Set oExec = WshShell.Exec(ProgramFiles64 & "\Confer\repcli.exe status")
    FormattedAVVersion = "Unknown"
    Do 'Because the status command outputs many lines of text, we need to use a Do/Loop command to grab all of it
        sLine = oExec.StdOut.ReadLine()
        If InStr(sLine, "Sensor Version") <> 0 Then
            FormattedAVVersion = Mid(sLine,17,10)
            Exit Do
        End If
    Loop While Not oExec.Stdout.atEndOfStream            

        
    output.writeline "- The version of Carbon Black running on this machine is: " & FormattedAVVersion
        
    ' Now let's determine whether or not Carbon Black is up-to-date.
    If objFSO.FileExists(ProgramFiles64 & "\Confer\db_cfg") Then
        Set objFile = objFSO.GetFile(ProgramFiles64 & "\Confer\db_cfg")
        AVDatVersion = objFile.DateLastModified
    Else
        AVDatVersion="01/01/2001"
	   output.writeline "- The AV Definition file was not found. Odd."
    End If
    
    CalculateAVAge 'Call the function to determine how old the A/V Definitions are

        
    ' Now let's use the repcli executable to determine whether or not Real-Time Scanning is enabled.
    Set oExec = WshShell.Exec(ProgramFiles64 & "\Confer\repcli.exe status")
    OnAccessScanningEnabled = False
    Do 'Because the status command outputs many lines of text, we need to use a Do/Loop command to grab all of it
        sLine = oExec.StdOut.ReadLine()
        If InStr(sLine, "Sensor State[Enabled]") <> 0 Then
            OnAccessScanningEnabled = True
            Exit Do
        End If
    Loop While Not oExec.Stdout.atEndOfStream   
        
    output.writeline "- Is Real Time Scanning Enabled? " & OnAccessScanningEnabled

End Sub          
        

' *************************************  
' Sub: ObtainWindowsDefenderData
' *************************************
Sub ObtainWindowsDefenderData

  
  
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


Function IsServiceRunning(ServiceName)
  Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  
  Set colServices = objWMIService.ExecQuery("Select * from Win32_Service WHERE Name = '" & ServiceName & "'")
  
  For Each objService in colServices 
    ' output.writeline "- The service name is : " & ServiceName
    ' output.writeline "- The service status is : " & objService.State
    If objService.State = "Running" Then
      IsServiceRunning = TRUE
    Else
      IsServiceRunning = FALSE
    End If
  Next

End Function



Function DownloadFile(strFileURL, strHDLocation)

    ' Sourced from: https://serverfault.com/questions/29707/download-file-from-vbscript
    ' Fetch the file
    
    output.writeline "- Will attempt to download the following file: " & strFileURL
    output.writeline "- The file will be stored at the following path: " & strHDLocation
    objXMLHTTP.open "GET", strFileURL, false
    objXMLHTTP.send()

    If objXMLHTTP.Status = 200 Then
        Set objADOStream = CreateObject("ADODB.Stream")
        objADOStream.Open
        objADOStream.Type = 1 'adTypeBinary

        objADOStream.Write objXMLHTTP.ResponseBody
        objADOStream.Position = 0    'Set the stream position to the start

        Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.FileExists(strHDLocation) Then 
            objFSO.DeleteFile strHDLocation
            Set objFSO = Nothing
        End If
            
        objADOStream.SaveToFile strHDLocation & "eps.rmm.exe"
        objADOStream.Close
        Set objADOStream = Nothing
      
        DownloadFile = TRUE
    
    Else
        DownloadFile = FALSE
    End If

    Set objXMLHTTP = Nothing

End Function