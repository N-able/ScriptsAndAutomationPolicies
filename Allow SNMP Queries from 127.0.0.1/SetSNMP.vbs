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
' Date: March 5th, 2009
' Description: This VBS script will add localhost (127.0.0.1) as a host from which the device will accept SNMP queries. 
' Version: 1.0


'Let's declare some variables
option explicit
dim strComputer, Host, strKeyPath, Count, arrValueNames,arrValueTypes, objReg, KeyPresenceTest
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
strKeyPath = "SYSTEM\CurrentControlSet\Services\SNMP\Parameters\PermittedManagers"
Host = "127.0.0.1"


'Access the registry through WMI
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")


'Find out how many hosts already exist, so that the script creates an entry with the right Name 
'(each entry has a numerical Name in the registry, starting with 1)
objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,1,KeyPresenceTest

'This next line tests to see whether the entry (define by KeyPresenceTest) is null or not
If isNull (KeyPresenceTest) Then
  'Because the value is null, the registry entry we create needs to have a Name of 1
  Count = 1
Else
  'Get the entries, and stick it into an array
  objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath,arrValueNames, arrValueTypes
  'Grab the upper value of the array, and add 2 so that it creates a sequential Name value 
  Count = Ubound(arrValueNames) + 2
End if

'Add the entry into the registry
objReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,Count,Host

'Clear the variables
objReg=Null
Count=Null
arrValueNames=Null
arrValueTypes=Null