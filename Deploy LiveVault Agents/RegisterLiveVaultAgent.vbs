'##################################################################
'# Copyright 2008 by N-able Technologies                          #
'# All Rights Reserved                                            #
'# May not be reproduced or redistributed                         #
'# Without written consent from N-able Technologies               #
'# www.n-able.com			                                            #
'##################################################################




'###########################################################################################
'#  USER CONFIGURABLE SECTION

Option Explicit
Dim WshShell, HostnameCMD, Hostname, RegisterCMD, Username, Password, EncryptionKey, LiveVaultServer


Username = "LIVEVAULTUSERNAME"
Password = "LIVEVAULTPASSWORD"
EncryptionKey = "ENCRYPTIONKEY"
LiveVaultServer = "https://provisioning.livevault.com"
'###########################################################################################



'Get the hostname of the device
Set WshShell = CreateObject("Wscript.Shell")
Set HostnameCMD = WshShell.Exec("%COMSPEC% /c hostname")
Hostname = HostnameCMD.StdOut.ReadLine
Set WshShell = Nothing


' Run the registration command
Set WshShell = CreateObject("Wscript.Shell")
RegisterCMD = "" & chr(34) & "C:\Program Files\LiveVault Corporation\backupengine\LVRegister" & chr(34) & " -n -i -u " & chr(34) & Username & chr(34) & " -p " & Password & " -s " & LiveVaultServer & " -k " & EncryptionKey & " -l " & Hostname & " -w " & LiveVaultServer & ""
WshShell.Run RegisterCMD,7,TRUE
Set WshShell = Nothing 
