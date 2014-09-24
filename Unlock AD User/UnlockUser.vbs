'**********************************************************
' Script: UnlockUser.vbs
' Version: 1.0
' Author: Chris Reid
' Date: Nov. 27th, 2009
'Usage: cscript UnlockUser.vbs DOMAINAME
'**********************************************************


'Let's declare some variables
Option explicit
dim Domain, objDomain, objUser



'Let's make sure that the person running this script specified a domain name as a command line parameter
If WScript.Arguments.Count = 1 Then
 Domain = WScript.Arguments.Item(0)
Else
 wscript.Quit(1) 	
End If




'Let's get a list of all users whose accounts are locked
Set objDomain = GetObject("WinNT://" & Domain &",domain")
objDomain.Filter = Array("User")


'Now let's figure out what accounts are locked, and unlock them
For Each objUser In objDomain 
    If objUser.IsAccountLocked then
	objUser.IsAccountLocked = False
	objUser.SetInfo
    End if	 
Next