
Set Shell=WScript.CreateObject("Shell.Application")
Shell.MinimizeAll
On Error Resume Next
Dim WshShell,a,i
i=0
Set WshShell=WScript.CreateObject("WScript.Shell")
Set Shell=WScript.CreateObject("Shell.Application")


a=WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{2318C2B1-4965-11d4-9B18-009027A5CD4F}\UninstallString")
If a<>"" Then 
		WshShell.Run(a&" /S"),1,True
		i=i+1
	end if
