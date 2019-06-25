'****************************************************************************************
' Script: UninstallAVG.vbs
' Version: 1.0
' Author: Chris Reid
' Description: This script will unintstall the 32-bit versions of AVG 8.x and 9.x
' Date: Dec. 11th. 2009
'***************************************************************************************


'**************************************************************************************
'NOTE: This script *will* cause the machine to be rebooted!!! Use with CAUTION!!
'**************************************************************************************



' Declare variables
option explicit
dim strSystemDrive, strFileURL, strHDLocation, objXMLHTTP, objADOStream, objFSO, Shell, strPartialTitle, objWord, colTasks, strWindowTitle, objTask


'Set the URL where the file should be downloaded from
strFileURL = "http://download.avg.com/filedir/util/avg_arm_sup_____.dir/avgremover.exe"




'We want to place the avgremover.exe on a the root of the system drive - so we need to find out the system drive
Set Shell = CreateObject("WScript.Shell")
    strSystemDrive = Shell.ExpandEnvironmentStrings("%SystemDrive%")
Set Shell = Nothing 

  
'Set the location where the file will be saved 
strHDLocation = "" & strSystemDrive & "\avgremover.exe"
    

'Run the download sub-routine, so that we can download the AVG uninstaller
Download


'Now that the file has been downloaded. create a shell and run it
'Create the Shell
Set Shell = CreateObject("Wscript.Shell")
'Run the AVG Uninstaller
Shell.Run "%COMSPEC% /c " & strHDLocation
Set Shell = Nothing




'As the uninstaller prompts the user for confirmation, we need to automatically click 'Yes'. So let's find the confirmation window.
FindWarningWindow

' Now we just need to activate the window, and pass ALT-Y to it
Confirm






'**********************************************************************************
' Sub: Download
' This sub will download the file from the net, and save it to the specified path
'**********************************************************************************
sub Download
    ' Let's download and save the avgremover.exe (the code for this section was grabbed from http://blog.netnerds.net/2007/01/vbscript-download-and-save-a-binary-file/)    

    ' Fetch the file
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    objXMLHTTP.open "GET", strFileURL, false
    objXMLHTTP.send()

    If objXMLHTTP.Status = 200 Then
        Set objADOStream = CreateObject("ADODB.Stream")
        objADOStream.Open
        objADOStream.Type = 1 'adTypeBinary
        objADOStream.Write objXMLHTTP.ResponseBody
        objADOStream.Position = 0    'Set the stream position to the start

    'If the file already exists, delete it. Otherwise, place the file in the specified location
    Set objFSO = Createobject("Scripting.FileSystemObject")
    If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
        Set objFSO = Nothing
        objADOStream.SaveToFile strHDLocation
        objADOStream.Close
        Set objADOStream = Nothing
    End if

  Set objXMLHTTP = Nothing

End Sub


'**********************************************************************************
' Sub: FindWarningWindow
' This sub finds the open window whose title is 'Warning'
'**********************************************************************************
Sub FindWarningWindow
  strPartialTitle = "Warning"
  Set objWord = CreateObject("Word.Application")
  Set colTasks = objWord.Tasks
  For Each objTask In colTasks
    	If objTask.Visible = True Then
  		If InStr(objTask.Name, strPartialTitle) > 0 Then
  			strWindowTitle = objTask.Name
  			Exit For
  		Else
  			strWindowTitle = ""
  		End If
  	End If
  Next
  objWord.Quit
End Sub


'**********************************************************************************
' Sub: Confirm
' This sub passes an ALT-Y command to the open window
'**********************************************************************************
Sub Confirm
  If strWindowTitle <> "" Then
  	Set Shell = CreateObject("WScript.Shell")
  	While Shell.AppActivate(strWindowTitle) = False
  		WScript.Sleep 10
  	Wend
  	Shell.AppActivate strWindowTitle
  	WScript.Sleep 100
  	While Shell.AppActivate(strWindowTitle) = True
  		Shell.SendKeys "%Y"
  		WScript.Sleep 2000
  	Wend
  Else
  End If
End Sub