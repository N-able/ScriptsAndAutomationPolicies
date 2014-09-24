' Modified from http://msdn.microsoft.com/en-us/library/aa387102%28VS.85%29.aspx
'
' Modifications 12/16/2010 by Steve Drees sdrees@greatinsight.com
'
Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()

const EVENTLOG_SUCCESS = 0
const EVENTLOG_ERROR = 1
const EVENTLOG_WARNING = 2
const EVENTLOG_INFORMATION = 4
const EVENTLOG_AUDIT_SUCCESS = 8
const EVENTLOG_AUDIT_FAILURE = 16

Set objShell = Wscript.CreateObject("Wscript.Shell")

objShell.LogEvent EVENTLOG_INFORMATION, "Searching for updates..."


Set searchResult = _
updateSearcher.Search("IsInstalled=0 and Type='Software' and AutoSelectOnWebSites=1")

If searchResult.Updates.Count = 0 Then
  objShell.LogEvent EVENTLOG_INFORMATION, "No updates to install. Aborting."
	WScript.Quit
End If

strUpdates = "List of applicable items on the machine:" & vbCRLF
For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    strUpdates = strUpdates & update.Title & vbCRLF
Next

objShell.LogEvent EVENTLOG_INFORMATION, strUpdates



Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    updatesToDownload.Add(update)
Next

'WScript.Echo vbCRLF & "Downloading updates..."
objShell.LogEvent EVENTLOG_INFORMATION, "Downloading updates..."


Set downloader = updateSession.CreateUpdateDownloader() 
downloader.Updates = updatesToDownload
downloader.Download()

strDownloads ="List of downloaded updates: " & vbCRLF
For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    If update.IsDownloaded Then
       WScript.Echo I + 1 & "> " & update.Title
			 strDownloads = strDownloads & update.Title & vbCRLF 
    End If
Next

objShell.LogEvent EVENTLOG_INFORMATION, strDownloads


Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")


For I = 0 To searchResult.Updates.Count-1
    set update = searchResult.Updates.Item(I)
    If update.IsDownloaded = true Then
       updatesToInstall.Add(update)	
    End If
Next

strInput = "Y"


If (strInput = "N" or strInput = "n") Then 
  objShell.LogEvent EVENTLOG_INFORMATION, "Not Installing" 
	WScript.Quit
ElseIf (strInput = "Y" or strInput = "y") Then
objShell.LogEvent EVENTLOG_INFORMATION, "Installing updates..."
	Set installer = updateSession.CreateUpdateInstaller()
	installer.Updates = updatesToInstall
	Set installationResult = installer.Install()
	
	'Output results of install
	
  strResults = "Installation Result: " & _
               installationResult.ResultCode & vbCRLF & "Reboot Required: " & _ 
	             installationResult.RebootRequired & vbCRLF 
  objShell.LogEvent EVENTLOG_INFORMATION, strResults 							 
							 

End If
