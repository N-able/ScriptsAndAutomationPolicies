' Script Name: RunDefender.vbs
' Description: This will run a Windows Defender scan.


'Run the command
Set Shell = CreateObject("Wscript.Shell")
Set filesys = CreateObject("Scripting.FileSystemObject")
PFiles=Shell.ExpandEnvironmentStrings("%ProgramFiles%")
Comspec=Shell.ExpandEnvironmentStrings("%COMSPEC%")
path = PFiles & "\Windows Defender\MpCmdRun.exe"
If filesys.FileExists(path) Then  
     cmdline = Comspec & " /c """ & path & """ -scan"
     Shell.Run cmdline,7,TRUE
End If

Set filesys = Nothing
Set Shell = Nothing