Set sh = CreateObject("Wscript.Shell")

cleaningPoints = array("Active Setup Temp Folders", _
									  "Content Indexer Cleaner", _
									  "Downloaded Program Files", _
									  "Hibernation File", _
									  "Internet Cache Files", _
									  "Memory Dump Files", _
									  "Microsoft Office Temp Files", _
									  "Old ChkDsk Files", _
									  "Previous Installations", _
									  "Recycle Bin", _
									  "Setup Log Files", _
									  "System error memory dump files", _
									  "System error minidump files", _
									  "Temporary Files", _
									  "Temporary Setup Files", _
									  "Temporary Sync Files", _
									  "Thumbnail Cache", _
									  "Upgrade Discarded Files", _
									  "Windows Error Reporting Archive Files", _
									  "Windows Error Reporting Queue Files", _
									  "Windows Error Reporting System Archive Files", _
									  "Windows Error Reporting System Queue Files")

function createRegKeys()
	for each cleaningPoint in cleaningPoints
		sh.regwrite"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\" & cleaningPoint & "\StateFlags1337", 2, "REG_DWORD"
	next
end function

function deleteRegKeys()
	for each cleaningPoint in cleaningPoints
		sh.regdelete"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\" & cleaningPoint & "\StateFlags1337"
	next
end function

createRegKeys()
sh.run "cleanmgr.exe /D C /sagerun:1337",7,TRUE
deleteRegKeys()

Set sh = Nothing
