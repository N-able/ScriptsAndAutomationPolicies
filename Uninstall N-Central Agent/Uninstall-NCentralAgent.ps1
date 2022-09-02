# '==================================================================================================================================================================
# 'Disclaimer
# 'The sample scripts are not supported under any SolarWinds support program or service. 
# 'The sample scripts are provided AS IS without warranty of any kind. 
# 'SolarWinds further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. 
# 'The entire risk arising out of the use or performance of the sample scripts and documentation stays with you. 
# 'In no event shall SolarWinds or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
# '(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) 
# 'arising out of the use of or inability to use the sample scripts or documentation.
# '==================================================================================================================================================================

#Added function to remove new NC agents and delete folder

#EcoSystem Agent
try {
        Start-Process -NoNewWindow -FilePath "C:\Program Files (x86)\SolarWinds MSP\Ecosystem Agent\unins000.exe" -ArgumentList "/silent"
	} catch {
	    return
    }

#File Cache Service Agent
try {
        Start-Process -NoNewWindow -FilePath "C:\Program Files (x86)\MspPlatform\FileCacheServiceAgent\unins000.exe" -ArgumentList "/silent"
	} catch {
	    return
    }

#Request Handler Agent
try {
       Start-Process -NoNewWindow -FilePath "C:\Program Files (x86)\MspPlatform\RequestHandlerAgent\unins000.exe" -ArgumentList "/silent"
    } catch {
	    return
    }

#Patch Management Service
try {
        Start-Process -NoNewWindow -FilePath "C:\Program Files (x86)\MspPlatform\PME\unins000.exe" -ArgumentList "/silent"
    } catch {
        return
    }

function getAgentPath() {
    WriteToLog I "Started running getAgentPath function."
    $Keys = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
	$Items = $Keys | Foreach-Object {Get-ItemProperty $_.PsPath}
	ForEach ($Item in $Items) {
		If ($Item.DisplayName -like "Windows Agent"){
            $script:localFolder = $Item.InstallLocation
            $script:uninstallString = $Item.UninstallString
			break
		}
    }
	try {
		$Keys = Get-ChildItem HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction Stop
	} catch {
	    return
    }
	$Items = $Keys | Foreach-Object {Get-ItemProperty $_.PsPath}
	ForEach ($Item in $Items) {
		If ($Item.DisplayName -like "Windows Agent"){
            $script:localFolder = $Item.InstallLocation
            $script:uninstallString = $Item.UninstallString
			break
		}
	}
    If ($null -eq $localFolder) {
        WriteToLog I "No Windows Agent located. Exiting."
        Exit 1001
    }
    If (($script:localFolder -match '.+?\\$') -eq $false) {
	    $script:localFolder = $script:localFolder + "\"
    }
    WriteToLog I "Agent install location: " $script:localFolder
    WriteToLog I "Agent uninstall string: " $script:uninstallString
    WriteToLog I "Completed running getAgentPath function."
}
function runUninstaller() {
	WriteToLog I "Started running runUninstaller function."
    $argumentList = ($script:uninstallString -split 'MsiExec.exe ')[1] + " /qn /norestart"
	WriteToLog I  "Starting MSI Uninstaller."
    Start-Process msiexec.exe -Wait -ArgumentList $argumentList
    WriteToLog I "Uninstall process has now been completed."
    WriteToLog I "Will now check if the Windows Agent location still exists."
    WriteToLog I "Completed running runUninstaller function."
}
function checkAgentLocation() {
	WriteToLog I "Started running checkAgentLocation function."
    $pathTester = Test-Path $localFolder
    If ($pathTester -eq $false) {
        WriteToLog I "Windows Agent location no longer present"    
    } Else {
        Remove-Item $localFolder -recurse
    }
    WriteToLog I "Completed running checkAgentLocation function."
}
function cleanupAssetTag() {
	WriteToLog I "Started running cleanupAssetTag function."
    $class='NCentralAssetTag'
    $namespace='root\cimv2\NCentralAsset'
    Get-WmiObject -Namespace $namespace -Class $class | Remove-WmiObject
    $path='HKLM:\SOFTWARE\N-able Technologies\NcentralAsset'
    $name='NcentralAssetTag'
    Remove-ItemProperty -Path $path -Name $name -Force
	if (Test-Path -LiteralPath "C:\Program Files (x86)\N-able Technologies\NcentralAsset.xml") {
		try {
			Remove-Item -path "C:\Program Files (x86)\N-able Technologies\NcentralAsset.xml"
		}
		catch {
			WriteToLog E "Skipping, file not found."
		}
		WriteToLog I "Successfully removed Asset Tags"
    }
    WriteToLog I "Completed running cleanupAssetTag function."
}
function cleanupTakeControl() {
	WriteToLog I "Started running cleanupTakeControl function."
		if (Test-Path -LiteralPath "C:\Program Files (x86)\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe") {
			try { 
				Start-Process msiexec.exe -Wait "C:\Program Files (x86)\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe" /S
			}
			catch {
				WriteToLog E "Unable to find 64-bit installation, moving to 32-bit removal."
			}
			WriteToLog I "Successfully Removed Take Control 64 bit."
		}
		else {
			Start-Process msiexec.exe -Wait "C:\Program Files\BeAnywhere Support Express\GetSupportService_N-Central> uninstall.exe" /S
			WriteToLog E "Successfully Removed Take Control 32 bit."
    }
    WriteToLog I "Completed running cleanupTakeControl function."
}
function cleanupConfigBackup() {
	WriteToLog I "Started running cleanupConfigBackup function."
    Remove-Item -path "C:\ProgramData\N-Able Technologies\Windows Agent\config\ConnectionString_Agent.xml"
    WriteToLog I "Completed running cleanupConfigBackup function."
}
function getTimeStamp() {
	return "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)
}
function writeToLog($state, $message) {
	switch -regex -Wildcard ($state) {
		"I" {
			$state = "INFO"
		}
		"E" {
			$state = "ERROR"
		}
		"W" {
			$state = "WARNING"
		}
		"F"  {
			$state = "FAILURE"
		}
		""  {
			$state = "INFO"
		}
		Default {
			$state = "INFO"
		}
	 }
	Write-Host "$(getTimeStamp) - [$state]: $message"
}
function main() {
    getAgentPath
    runUninstaller
    checkAgentLocation
    cleanupAssetTag
	cleanupTakeControl
	cleanupConfigBackup
}
main