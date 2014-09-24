'##########################################
'# Copyright 2009 by N-able Technologies  #
'# All Rights Reserved                    #
'# May not be reproduced or redistributed #
'# Without written consent from N-able    #
'# Technologies			                      #
'# www.n-able.com			                    #
'##########################################


' Script Name: SetSNMP.vbs
' Author: Chris Reid
' Date: Septh. 17th, 2009
' Description: This VBS script will add localhost (127.0.0.1) as a host from which the device will accept SNMP queries, and will add an SNMP community string using the one specified. 
' Version: 1.0

'NOTE: The SNMP service does not need to be restarted for these changes to take effect



'Let's declare some variables
option explicit
dim strComputer, Host, strHostKeyPath, strCommunityKeyPath, Count, arrValueNames,arrValueTypes, objReg, KeyPresenceTest, CommunityString, CommunityType
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
strHostKeyPath = "SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers"
strCommunityKeyPath = "SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities"
Host = "127.0.0.1"
CommunityType = "4" 'According to http://technet.microsoft.com/en-us/library/cc736898%28WS.10%29.aspx this will set the community string to be Read Only. 4=ReadOnly, 8=Read/Write



'Let's make sure that we got the SNMP community string as a command line parameter
If WScript.Arguments.Count = 1 Then
 CommunityString = WScript.Arguments.Item(0)
Else
 wscript.Quit(1) 	
End If


'STEP 1: ADD 127.0.0.1 as an accepted SNMP Host

'Access the registry through WMI
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'Find out how many hosts already exist, so that the script creates an entry with the right Name 
'(each entry has a numerical Name in the registry, starting with 1)
objReg.GetStringValue HKEY_LOCAL_MACHINE,strHostKeyPath,1,KeyPresenceTest

'This next line tests to see whether the entry (define by KeyPresenceTest) is null or not
If isNull (KeyPresenceTest) Then
  'Because the value is null, the registry entry we create needs to have a Name of 1
  Count = 1
Else
  'Get the entries, and stick it into an array
  objReg.EnumValues HKEY_LOCAL_MACHINE, strHostKeyPath,arrValueNames, arrValueTypes
  'Grab the upper value of the array, and add 2 so that it creates a sequential Name value 
  Count = Ubound(arrValueNames) + 2
End if

'Add the entry into the registry
objReg.SetStringValue HKEY_LOCAL_MACHINE,strHostKeyPath,Count,Host

'Clear the variables
objReg=Null
Count=Null
arrValueNames=Null
arrValueTypes=Null



'STEP 2: Add the specified SNMP Community String
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
objReg.SetDWordValue HKEY_LOCAL_MACHINE,strCommunityKeyPath,CommunityString,CommunityType


'Clear the variable
objReg=Null