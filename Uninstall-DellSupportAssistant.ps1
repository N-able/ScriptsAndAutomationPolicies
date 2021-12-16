<#    
    ************************************************************************************************************
    Name: Uninstall-DellSupportAssistant

    Version: 0.1 (03rd May 2019)
    Purpose:    Uninstall Dell Support Assistant. Can specify a min/max version to target against.
    Based on a a script provided by JonCz

    Pre-Reqs:    Powershell 2

    0.1 + Initial Release
    ************************************************************************************************************
#>


Param (
	[Parameter(Mandatory=$false)]
    [switch]$Force = $false,
    [Parameter(Mandatory=$false)]
    [String] $SoftwareVersion = "3.2.0.90"
)

$Software = "Dell SupportAssist"

#region Functions

$ScriptName = $myinvocation.MyCommand
Function Remove-CachedScripts {
    # Removed Cached Copies of Script

    $Agent = "C:\Program Files\N-able Technologies\Windows Agent"

    if (test-path "C:\Program Files (x86)\N-able Technologies\Windows Agent") {
        $Agent = "C:\Program Files (x86)\N-able Technologies\Windows Agent"
    }
    remove-item "$Agent\cache\$scriptname" -force -erroraction SilentlyContinue
    remove-item "$Agent\Temp\Script\$scriptname"  -force -ErrorAction SilentlyContinue
}
Function Get-InstalledApp($Target) {
	$Apps = @()
	
	$Apps += Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall  |
		   Get-ItemProperty | Where-Object {$_.DisplayName -eq "$Target" } | Select-Object -Property DisplayName, DisplayVersion, UninstallString
		   
#	If 64-bit OS, then search Wow6432Node Uninstall as well					
	If ($(Get-WmiObject Win32_OperatingSystem).OSArchitecture -match "64") {
		$Apps += Get-ChildItem -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
		   Get-ItemProperty | Where-Object {$_.DisplayName -eq "$Target" } | Select-Object -Property DisplayName, DisplayVersion, UninstallString
		}
		
	$Apps
}

Function Uninstall-Software {

    if ($SoftwareUninstall -ne $null) {
        Write-Host "$($SoftwareUninstall.DisplayName) $($SoftwareUninstall.DisplayVersion) was detected."
        
		if (($($SoftwareUninstall.DisplayVersion) -ge "$SoftwareVersion") -and ($force -eq $false)) {
            Write-Host "This version of $Software is not affected by the Vulnerability" -ForegroundColor Green
            exit 0
        }
        if ($force -eq $true) {
			Write-Warning "Force Mode was enabled. Uninstalling $Software..."
            }

		if ($($SoftwareUninstall.DisplayVersion) -lt "$SoftwareVersion") {
			Write-Warning "$Software pre $SoftwareVersion Found. Uninstalling..."
            }
					ForEach ($item in $SoftwareUninstall) {
						if (($item.displayversion -le "$softwareversion") -or ($force -eq $true)) {
						Write-Output "Removing $($item.DisplayName) via $($item.UninstallString)"

						If ($item.UninstallString) {
						
                            $uninst = $item.UninstallString.Split('/X')[0]
                            $guid = $item.UninstallString.Split('/X')[2] 
                            $argument = "/x $guid /qn /norestart"
                            
							Write-Output "Running $uninst $argument"
							Start-Process $uninst -ArgumentList "$argument" -wait -passthru -NoNewWindow
							
							$Timeout = 0
							$MaxWait = 10 # Wait up to 5 minutes for the application to uninstall
							
							While (($CheckSoftware -ne $null) -and ($Timeout -lt $MaxWait)) { 
								Start-Sleep -Seconds 30
								$Timeout++
								
								$CheckSoftware = Get-InstalledApp("$Software")
							}

							If ($CheckSoftware -eq $null) {
								Write-Output "$Software removed"
						} else {
							Write-Output "ERROR: $Software not successfully removed."
						}
					}
				}
			}
		} 
else {
	Write-Host "$Software $softwareversion" -ForegroundColor Green -nonewline; Write-host " or lower is not present. Nothing to do."
}
}
#endregion Functions

#region Main
 
Write-Host "Checking for " -nonewline; Write-Host "$software $softwareversion" -ForegroundColor Green
$SoftwareUninstall = . Get-InstalledApp($software)
. Uninstall-Software
Remove-CachedScripts

#endregion Main