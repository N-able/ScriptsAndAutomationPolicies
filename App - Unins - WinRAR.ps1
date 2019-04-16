<#
App - Unins - WinRAR.ps1

This script uninstalls WinRAR, and if
successful, removes the program folder to get rid of any
remnants

Created by:	Jon Czerwinski, Cohn Consulting Corporation
Date:		20190319
Version:
	1.0		Initial release

#>

#region Functions
Function Get-InstalledApp($Target) {
	$Apps = @()
	
	$Apps += Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall  |
		   Get-ItemProperty | Where-Object {$_.DisplayName -like "*$Target*" } | Select-Object -Property DisplayName, UninstallString
		   
#	If 64-bit OS, then search Wow6432Node Uninstall as well					
	If ($(Get-WmiObject Win32_OperatingSystem).OSArchitecture -eq "64-bit") {
		$Apps += Get-ChildItem -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
		   Get-ItemProperty | Where-Object {$_.DisplayName -like "*$Target*" } | Select-Object -Property DisplayName, UninstallString
		}
		
	$Apps
}
#endregion Functions

#region Main
 
Write-Output "Getting uninstall string for 'WinRAR'"
$WinRAR = Get-InstalledApp("WinRAR")

if ($WinRAR -ne $null) {
	Write-Output "WinRAR Found.  Uninstalling"
 
	ForEach ($item in $WinRAR) {

		Write-Output "Removing $($item.DisplayName) via $($item.UninstallString)"

		If ($item.UninstallString) {
		
			$uninst = $item.UninstallString

			Write-Output "Running $uninst"
			Start-Process $uninst -ArgumentList "/S" -NoNewWindow
			
			$Timeout = 0
			$MaxWait = 10 # Wait up to 5 minutes for the application to uninstall
			
			While (($ChkWinRAR -ne $null) -and ($Timeout -lt $MaxWait)) { 
				Start-Sleep -Seconds 30
				$Timeout++
				
				$ChkWinRAR = Get-InstalledApp("WinRAR")
			}

			If ($ChkWinRAR -eq $null) {
				Write-Output "WinRAR removed"
			} else {
				Write-Output "ERROR: WinRAR not successfully removed."
			}
		}
	}
} else {
	Write-Output "WinRAR not present.  Nothing to do."
}
#endregion Main