'****************************************************************************************
' Script: RunTMDiskCleaner.vbs
' Version: 1.0
' Author: Chris Reid
' Description: This script will run the 'TMDiskCleaner' utility
' Date: Dec. 15th, 2009
'***************************************************************************************




' Declare variables
option explicit
dim strFileURL, strUnZipProgram, strTemp, strHDZipLocation, objXMLHTTP, objADOStream, objFSO, Shell, bDiskCleanerExists


'Set a few constants
strFileURL = "http://solutionfile.trendmicro.com/solutionfile/1055066/EN/DiskCleaner_v1_0.zip"


'We want to place the downloaded zip file in the %TEMP% directory
'As the agent runs as LOCAL SYSTEM, this will be %SYSTEMDRIVE%\Windows\TEMP
Set Shell = CreateObject("WScript.Shell")
    strTemp = Shell.ExpandEnvironmentStrings("%TEMP%")
Set Shell = Nothing 

'Check if it has been downloaded before
Set objFSO = Createobject("Scripting.FileSystemObject")
bDiskCleanerExists = objFSO.FileExists(strTemp & "\Disk Cleaner v1.0\TMDiskCleaner.exe")
Set objFSO = Nothing
  
'If it doesn't exist, download and extract it
If Not bDiskCleanerExists Then 
    'Set the location where the file will be saved 
    strHDZipLocation = "" & strTemp & "\DiskCleaner_v1_0.zip"

    'Run the download sub-routine
    Download

    'Now that the file has been downloaded. unzip it
    Extract strHDZipLocation, strTemp
End If

'let's run the TMDiskCleaner executable
'Create the Shell
Set Shell = CreateObject("Wscript.Shell")
'Run the executable
Shell.Run """" & strTemp & "\Disk Cleaner v1.0\TMDiskCleaner.exe"" /hide /log", 7, TRUE
Set Shell = Nothing








'**********************************************************************************
' Sub: Download
' This sub will download the file from the net, and save it to the specified path
'**********************************************************************************
sub Download
    ' Let's download and save 'DiskCleaner_v1_0.zip' (the code for this section was grabbed from http://blog.netnerds.net/2007/01/vbscript-download-and-save-a-binary-file/)    

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
    	If objFSO.FileExists(strHDZipLocation) Then objFSO.DeleteFile strHDZipLocation
        Set objFSO = Nothing
        
	objADOStream.SaveToFile strHDZipLocation
        objADOStream.Close
        Set objADOStream = Nothing
    End If

  Set objXMLHTTP = Nothing

End Sub

'******************************************************************************
' Sub: Extract
' Extracts files from a ZIP file.
' From: http://www.robvanderwoude.com/vbstech_files_zip.php
'******************************************************************************

Sub Extract( ByVal myZipFile, ByVal myTargetDir )
' Function to extract all files from a compressed "folder"
' (ZIP, CAB, etc.) using the Shell Folders' CopyHere method
' (http://msdn2.microsoft.com/en-us/library/ms723207.aspx).
' All files and folders will be extracted from the ZIP file.
' A progress bar will be displayed, and the user will be
' prompted to confirm file overwrites if necessary.
'
' Note:
' This function can also be used to copy "normal" folders,
' if a progress bar and confirmation dialog(s) are required:
' just use a folder path for the "myZipFile" argument.
'
' Arguments:
' myZipFile    [string]  the fully qualified path of the ZIP file
' myTargetDir  [string]  the fully qualified path of the (existing) destination folder
'
' Based on an article by Gerald Gibson Jr.:
' http://www.codeproject.com/csharp/decompresswinshellapics.asp
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com

    Dim intOptions, objShell, objSource, objTarget

    ' Create the required Shell objects
    Set objShell = CreateObject( "Shell.Application" )

    ' Create a reference to the files and folders in the ZIP file
    Set objSource = objShell.NameSpace( myZipFile ).Items( )

    ' Create a reference to the target folder
    Set objTarget = objShell.NameSpace( myTargetDir )

    ' These are the available CopyHere options, according to MSDN
    ' (http://msdn2.microsoft.com/en-us/library/ms723207.aspx).
    ' On my test systems, however, the options were completely ignored.
    '      4: Do not display a progress dialog box.
    '      8: Give the file a new name in a move, copy, or rename
    '         operation if a file with the target name already exists.
    '     16: Click "Yes to All" in any dialog box that is displayed.
    '     64: Preserve undo information, if possible.
    '    128: Perform the operation on files only if a wildcard file
    '         name (*.*) is specified.
    '    256: Display a progress dialog box but do not show the file
    '         names.
    '    512: Do not confirm the creation of a new directory if the
    '         operation requires one to be created.
    '   1024: Do not display a user interface if an error occurs.
    '   4096: Only operate in the local directory.
    '         Don't operate recursively into subdirectories.
    '   8192: Do not copy connected files as a group.
    '         Only copy the specified files.
    intOptions = 4 Or 16 Or 512 Or 1024 

    ' UnZIP the files
    objTarget.CopyHere objSource, intOptions

    ' Release the objects
    Set objSource = Nothing
    Set objTarget = Nothing
    Set objShell  = Nothing
End Sub 