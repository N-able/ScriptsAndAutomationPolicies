Set Shell=WScript.CreateObject("Shell.Application")
Shell.MinimizeAll
On Error Resume Next
Dim WshShell,a,i
Set i=0
Set WshShell=WScript.CreateObject("WScript.Shell")
Set Shell=WScript.CreateObject("Shell.Application")

'This Removes the Yahoo Toolbar
a=WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Yahoo! Companion\UninstallString")
If a<>"" Then
WshShell.Run(""""&a&""" /S"),1,True
i=i+1
end if


'This Removes the Software Update Service
a=WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Yahoo! Software Update\UninstallString")
If a<>"" Then
WshShell.Run(""""&a&""" /S"),1,True
i=i+1
end if

'This Removes the Install Helper
a=WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\YInstHelper\UninstallString")
If a<>"" Then
WshShell.Run(""""&a&""" /S"),1,True
i=i+1
end if
