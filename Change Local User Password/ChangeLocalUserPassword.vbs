'**********************************************************
' Script: ChangeLocalUserPassword.vbs
' Version: 1.1
' Author: Chris Reid
' Date: Nov. 27th, 2009
' Usage: cscript ChangeLocalUserPassword.vbs USER PASSWORD
'**********************************************************

' Version History

' 1.1 - Feb 17th, 2010 - Fixed an issue where the objUser variable was not being defined (thanks Steve Grabowski!)
' 1.0 - Nov. 27th 2009 - Initial Release
                           



'Let's declare some variables
Option explicit
dim User, Password, WshNetwork, Device, objUser



'Let's make sure that the person running this script specified the User and Password
If WScript.Arguments.Count = 2 Then
 User = WScript.Arguments.Item(0)
 Password = WScript.Arguments.Item(1)
Else
 wscript.Quit(1)    
End If


'Let's get the name of this device (it's needed for changing the user account)
Set WshNetwork = WScript.CreateObject("WScript.Network")
Device = WshNetwork.ComputerName



'Let's bind to the user's account, and change his password
Set objUser = GetObject("WinNT://"& Device & "/"& User)
objUser.SetPassword("" & Password)