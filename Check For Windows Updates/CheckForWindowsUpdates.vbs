' Script Name: CheckForWindowsUpdates.vbs
' Description: This will run a Windows Defender scan.


'Run the command
Set Shell = CreateObject("Wscript.Shell")
Set filesys = CreateObject("Scripting.FileSystemObject")
SRoot=Shell.ExpandEnvironmentStrings("%SystemRoot%")
Comspec=Shell.ExpandEnvironmentStrings("%COMSPEC%")
path = SRoot & "\System32\wuauclt.exe"
If filesys.FileExists(path) Then  
     cmdline = Comspec & " /c """ & path & """ /detectnow"
     Shell.Run cmdline,7,TRUE
End If

Set filesys = Nothing
Set Shell = Nothing