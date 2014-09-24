'This program was written by Darren Wiebe of Cougar Ridge Computer Systems.  It is
'free for use by N-Able Partners.
'On Windows XP this will silently install MBSA.  On Vista and 7 the UAC is a problem so it will
'prompt the user.

'Set temporary directory
strHDLocation = WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
'strHDLocation = "C:\Temp"

'Check Operating System so we know if we need the 32bit or 64bit version
set WshShell = CreateObject("WScript.Shell")
OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
If (OsType = "x86") Then
    strFileURL = "http://download.microsoft.com/download/4/f/3/4f3044cb-0cf1-4c59-9da8-df6f8b1df6ef/MBSASetup-x86-EN.msi"
    strHDLocation = strHDLocation  & "\MBSASetup-x86-EN.msi"
Else
    strFileURL = "http://download.microsoft.com/download/4/f/3/4f3044cb-0cf1-4c59-9da8-df6f8b1df6ef/MBSASetup-x64-EN.msi"
    strHDLocation = strHDLocation & "\MBSASetup-x64-EN.msi"
End If

'Check for Vista or XP.
const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
dim Vistaver, dwvalue, UACStatus, UACValue
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_OperatingSystem")
For Each objItem in colItems
'msgBox(objItem.Caption)
If Instr(objItem.Caption, "Vista") Then
    OperatingSystem = "UAC"
ElseIf Instr(objItem.Caption, "Windows 7") Then
    OperatingSystem = "UAC"
End If    
Next


'Begin Download
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
 
    objXMLHTTP.open "GET", strFileURL, false
    objXMLHTTP.send()
 
    If objXMLHTTP.Status = 200 Then
'      msgBox("Downloading")
      Set objADOStream = CreateObject("ADODB.Stream")
      objADOStream.Open
      objADOStream.Type = 1 'adTypeBinary
 
      objADOStream.Write objXMLHTTP.ResponseBody
      objADOStream.Position = 0    'Set the stream position to the start
 
      Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.Fileexists(strHDLocation) Then 
            objFSO.DeleteFile strHDLocation
'            msgBox("Deleting File")
        End If    
      Set objFSO = Nothing
 
      objADOStream.SaveToFile strHDLocation
      objADOStream.Close
      Set objADOStream = Nothing
    End if
 
    Set objXMLHTTP = Nothing
    

Set Command = WScript.CreateObject("WScript.Shell")
if (OperatingSystem = "UAC") Then
    answer = MsgBox("This Microsoft security application is being installed on your system to assist in ensuring your computer security.  If you have questions or are concerned that this may be a virus please contact us at 000-000-0000.",0,"IT Support")
    cmd = "msiexec.exe /I " & strHDLocation & " /QB+"
Else
    cmd = "msiexec.exe /I " & strHDLocation & " /QN"
End If
WshShell.Run (cmd) 