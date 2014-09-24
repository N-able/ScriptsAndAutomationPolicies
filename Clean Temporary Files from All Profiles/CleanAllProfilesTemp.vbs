
dim objWSH, sProfile, objFolder
dim objFSO, sProfileRoot, objProfileFolder
dim sWindows, sSysDrive, strDateToday, strUserTemp
dim OSType, strError, output

set objFSO=CreateObject("Scripting.FileSystemObject")
Set output = Wscript.stdout
strDateToday = Date

'Determine the OS version and set the user temp folder accordingly
OSType = FindOSType 
If OSType="Windows 7" Or OSType="Windows Vista" Then 
strUserTemp=userProfile & "\AppData\Local\Temp" 
ElseIf  OSType="Windows 2003" Or OSType="Windows XP" Or OSType="Windows 2000" Then 
strUserTemp="\Local Settings\Temp" 
ElseIf OSType="Other" Then
WScript.Quit (100)
End If 
output.writeline "Detected Windows Version: " & OSType

' Get user profile root folder
set objWSH    = CreateObject("WScript.Shell")
sWindows = objWSH.ExpandEnvironmentStrings("%WINDIR%")
sSysDrive = objWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
sProfileRoot = ReadReg("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\ProfilesDirectory")
output.writeline "Raw Registry Read Result: " & sProfileRoot
sProfileRoot = Replace (LCase(sProfileRoot), "%systemdrive%", sSysDrive)
output.writeline "After Variable Replacement: " & sProfileRoot
set objWSH=nothing

'Delete user temp folders
set objProfileFolder=objFSO.GetFolder(sProfileRoot)
for each objFolder in objProfileFolder.SubFolders
    select case LCase(objFolder.Name)
        case "public": ' do nothing
        case "default": ' do nothing
        case "all users": ' do nothing
        case "default user": ' do nothing
        case "localservice": ' do nothing
        case "networkservice": ' do nothing
        case else:
            output.writeline "Processing profile: " & sProfileRoot & "\" & objFolder.Name & strUserTemp
            sProfile=sProfileRoot & "\" & objFolder.Name
            DeleteFolderContents sProfile & strUserTemp
	    If Err Then Call LogError(sProfile)
    end select
next

'Delete the windows\temp folder
output.writeline "Processing folder: " & sWindows & "\Temp"
DeleteFolderContents sWindows & "\Temp"
If Err Then Call LogError(sWindows)

output.writeline "Cleanup is finished " & strDateToday

sub DeleteFolderContents(strFolder)
    ' Deletes all files and folders within the given folder
    dim objFolder, objFile, objSubFolder
    on error resume next
    
    set objFolder=objFSO.GetFolder(strFolder)
    if Err.Number<>0 then
        call LogError (strFolder)
        Exit sub ' Couldn't get a handle to the folder, so can't do anything
    end if
    for each objSubFolder in objFolder.SubFolders
        objSubFolder.Delete true
        if Err.Number<>0 then
            'Try recursive delete (ensures better result)
            Err.Clear
            DeleteFolderContents(strFolder & "\" & objSubFolder.Name)
	    If Err Then Call LogError(strFolder & "\" & objSubFolder.Name)
        end if
    next
    for each objFile in ObjFolder.Files
		objFile.Delete true
        	if Err.Number<>0 then Call LogError (objFile) ' In case we couldn't delete a file
    next
end sub

Sub LogError (strError)

output.writeline Err.Number & " " & Err.Description & " " & strError
Err.Clear

End Sub

Function ReadReg(RegPath)
     Dim objRegistry, Key
     Set objRegistry = CreateObject("Wscript.shell")
     Key = objRegistry.RegRead(RegPath)
     ReadReg = Key
     Set objRegistry = Nothing
End Function

Function FindOSType 
    'Defining Variables 
    Dim objWMI, objItem, colItems 
    Dim OSVersion, OSName 
    Dim ComputerName 
      
     ComputerName="." 
      
    'Get the WMI object and query results 
    Set objWMI = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2") 
    Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem",,48) 
  
    'Get the OS version number (first two) and OS product type (server or desktop)  
    For Each objItem in colItems 
        OSVersion = Left(objItem.Version,3) 
                 
    Next 
  
     
    Select Case OSVersion 
        Case "6.1" 
            OSName = "Windows 7" 
        Case "6.0"  
            OSName = "Windows Vista" 
        Case "5.2"  
            OSName = "Windows 2003" 
        Case "5.1"  
            OSName = "Windows XP" 
        Case "5.0"  
            OSName = "Windows 2000" 
	Case Else
	    OSName = "Other"
   End Select 
  
    'Return the OS name 
    FindOSType = OSName 
     
    'Clear the memory 
    Set colItems = Nothing 
    Set objWMI = Nothing 
End Function 