<#    
   *******************************************************************************************************************************
    Name:            Repair-PME.ps1
    Version:         0.1.7.1 (16/06/2020)
    Purpose:         Install/Reinstall Patch Management Engine (PME)
    Created by:      Ashley How
    Thanks to:       Jordan Ritz for initial Get-PMESetup function code. Thanks to Prejay Shah for input into script.
    Pre-Reqs:        PowerShell 2.0 (PowerShell 4.0+ for Connectivity Tests)
    Version History: 0.1.0.0 - Initial Release.
                     0.1.1.0 - Update PMESetup_details.xml URL, update install arguments, create new cleanup function 
                               as suggested by Jan Tauwinkl at Solarwinds. Thanks for your input. Rename Function
                               Install-PMESetup to Install-PME.
                     0.1.1.1 - Update Get-PMESetup function for better error detection when unable to grab
                               PMESetup_Details.xml. Thanks to Jordan Ritz.       
                     0.1.2.0 - Update Get-PMESetup function for PS 2.0 support. Import of BitsTransfer Module for PS 2.0.
                     0.1.3.0 - Update Get-LegacyHash function to load system.security assembly for OS compatibility reasons.
                               Update Get-PMESetup function to use alternative download method for consistency.
                               Update Cleanup-PME function to detect if PMESetup is running and if so terminate.
                               Update Install-PME function to resolve issue with hashing not working due to invalid parameter. 
                               Improve error handling of Get-LegacyHash, Get-PMESetup, Cleanup-PME and Install-PME functions.
                               New function 'Download-PME' to avoid duplicate code.
                     0.1.4.0 - Update script to PSScriptAnalyzer best practices.
                               Rename Function 'Cleanup-PME' to 'Clear-PME' to conform to approved verbs for PS.
                               Rename Function 'Download-PME' to 'Get-PMESetup' to conform to approved verbs for PS.
                               Rename Function 'Get-PMESetup' to 'Get-PMESetupDetails' to release use for above function.
                               New function 'Stop-PMESetup' moving code from 'Clear-PME' function.
                               New function 'Set-Start' and 'Set-End' to write event log entries.
                     0.1.4.1 - New function 'Invoke-SolarwindsDiagnostics' to capture logs for support prior to repair.
                     0.1.4.2 - New function 'Confirm-Elevation' to confirm/debug UAC issues.
                               Updated function 'Invoke-SolarwindsDiagnostics' to create destination folder rather than
                               relying upon the tool to create it. 
                               Added DEBUG output to determinate parameters given to Solarwinds Diagnostics tool.
                               Moved function 'Get-PMESetupDetails' to run later on to avoid hung installer issues.
                               Updated function 'Stop-PMESetup' to also kill CacheServiceSetup and RPCServerServiceSetup.
                     0.1.5.0 - New function 'Test-Connectivity' to perform connectivity tests to destinations required for PME
                               PowerShell 4.0 or above required, connectivity tests will be skipped for any versions below.
                               Updated script to record all throws into the event log. 
                               Updated function 'Get-PMESetupDetails' to fallback to HTTP if HTTPS to sis.n-able.com fails
                               or if connectivity tests can't be performed due to low PowerShell version.
                               Moved function 'Set-Start' to execute first so all events can be recorded.
                               Updated 'Install-PME' to describe exit code 5 and link to documentation for other exit codes.
                     0.1.5.1 - Updated function 'Get-PMESetup' to support HTTPS to HTTP fallback.
                     0.1.6.0 - Updated 'Stop-PMESetup' function to colour code status if running interactively.
                               Updated 'Stop-PMESetup' function to check for and terminate _iu14D2N.tmp or similar process.
                               Updated 'Install-PME' function to to set PME to write install logs to 
                               'C:\ProgramData\SolarWinds MSP\Repair-PME\' instead of default location.
                               Updated 'Test-Connectivity' function to resolve issues where Win 7/2008 R2 has PS 4.0+.
                               Updated 'Invoke-SolarwindsDiagnostics' function for clearer output.
                               New Function 'Get-OSVersion' required for update to 'Test-Connectivity' function.   
                               New function 'Stop-PMEServices' to stop services prior to install of PME to prevent
                               access is denied errors as the installer doesn't have logic to forcefully terminate if 
                               still running after after a timeout.
                               Remove debug output from 'Invoke-SolarwindsDiagnostics' function as no longer required.
                     0.1.6.1 - Updated 'Stop-PMEServices' function to fix sc.exe not reverting recovery options correctly.
                     0.1.6.2 - Updated 'Test-Connectivity' function as did not actually commit the code as stated in 0.1.6.0.
                     0.1.6.3 - Updated 'Test-Connectivity' function to fix error on lines 150/151/155/156. 
                               Thanks for Clint Conner for finding.
                     0.1.6.4 - Updated 'Get-PMESetup and 'Install-PME' functions to address change made since PME version
                               1.2.4.2303 where 'PMESetup.exe' is downloaded as 'PMESetup_versionnumber.exe' This was causing
                               the installer to be downloaded unnecessarily.
                0.1.7.0 Beta - New function 'Set-PMEConfig' to apply fix for NCPM-4407 (System.OutOfMemoryException) this
                               will be applied by default however if you wish to change this behaviour change variable 
                               $NCPM4407 = "Yes" to $NCPM4407 = "No". Please note memory usage is increased slightly. PME must
                               already be installed and at version 1.2.5 or above to apply.
                             - New function 'Read-PMEConfig' to check PME Config and inform of possible misconfiguration.
                             - Updated 'Stop-PMEServices' function to deal with situations where services are suspended.
                             - New function 'Confirm-PMEInstalled' to confirm if PME is already installed.
                             - Updated function 'Get-PMESetupDetails' to create $LatestVersion variable used in
                               'Get-PMEConfigurationDetails' function. 
                             - New function 'Get-PMEConfigurationDetails' to obtain latest PME release date used in 
                               'Confirm-PMEUpdatePending' function.
                             - New function 'Confirm-PMEUpdatePending' safeguards running this script if PME is awaiting update
                               but has not updated yet. Change $Days = "2" variable to number of desired days to force
                               repair after new version of PME released. Change to $Days = "0" if you want this script to run
                               without safeguards. This new function means this script can finally be used for self-healing.
                             - Moved OS architecture detection out into its own function 'Get-OSArch'.
                             - Moved PowerShell detection out into its own function 'Get-PSVersion'.  
                             - Moved functions around to better suit new changes made.     
                             - Updated 'Read-PMEConfig' function to detect if Cache Service cache size is not set to default size.
                             - Updates to colour coding status in various areas of script.
                             - Various minor adjustments.
                     0.1.7.1 - Updated 'Read-PMEConfig','Set-PMEConfig','Confirm-PMEUpdatePending' and 'Get-OSArch' functions to 
                               fix 32-bit OS issues and when files don't exist, thanks to Prejay Shah for testing and code changes.
                             - Updated 'Confirm-PMEInstalled' function to fix issue where it was unable to correctly detect if 
                               PME is not installed.
                             - Various minor adjustments and fixes.
                             - Variables for settings moved to a dedicated section.                                   
   *******************************************************************************************************************************
#>
$Version = '0.1.7.1 (16/06/2020)'
Write-Host "Repair-PME $Version`n" -ForegroundColor Yellow

# Settings
# *******************************************************************************************************************************
# Change this variable to number of days (must be a number!) to begin repair after new version of PME is released. Default is 2.
$RepairAfterUpdateDays = "2"
# Change this variable to "No" if you don't want to apply fix for "System.OutOfMemoryException". Default is Yes.
$NCPM4407 = "Yes"
# *******************************************************************************************************************************

Function Set-Start {
    New-EventLog -LogName Application -Source "Repair-PME" -ErrorAction SilentlyContinue
    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Repair-PME has started, running version $Version.`nScript: Repair-PME.ps1"
}

Function Confirm-Elevation {
    # Confirms script is running as an administrator
    Write-Host "Checking for elevated permissions" -ForegroundColor Cyan
    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
        [Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Insufficient permissions to run this script. Run PowerShell as an administrator and run this script again.`nScript: Repair-PME.ps1"    
        Throw "Insufficient permissions to run this script. Run PowerShell as an administrator and run this script again."
    }
    Else {
    Write-Host "OK: Script is running as administrator" -ForegroundColor Green
    }
}

function Get-LegacyHash {
    Param($Path)
    # Performs hashing functionality with compatibility for older OS
    Try {
        Add-Type -AssemblyName System.Security
        $csp = new-object -TypeName System.Security.Cryptography.SHA256CryptoServiceProvider
        $ComputedHash = @()
        $ComputedHash = $csp.ComputeHash([System.IO.File]::ReadAllBytes($Path))
        $ComputedHash = [System.BitConverter]::ToString($ComputedHash).Replace("-", "").ToLower()
        Return $ComputedHash
    }
    Catch {
        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to performing hashing, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
        Throw "Unable to performing hashing, aborting. Error: $($_.Exception.Message)"
    }
}
Function Get-OSVersion {
    # Get OS version
    $OSVersion = (Get-WmiObject Win32_OperatingSystem).Caption
    #Workaround for WMI timeout or WMI returning no data
    If (($null -eq $OSVersion) -or ($OSVersion -like "*OS - Alias not found*")) {
        $OSVersion = (get-item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('ProductName')
    }
    Write-Output "OS: $OSVersion"
}

Function Get-OSArch {
    $OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
    Write-Output "OS Architecture: $OSArch"
    If ($OSArch -like "*64*") {
        $NCentralLog = "C:\Program Files (x86)\N-able Technologies\Windows Agent\log"
        }
        Else {
        $NCentralLog = "C:\Program Files\N-able Technologies\Windows Agent\log"
        }
}

Function Get-PSVersion {
    $PSVersion = $($PSVersionTable.PSVersion)
    Write-Output "PowerShell: $($PSVersionTable.PSVersion)"
}

Function Confirm-PMEInstalled {
    # Check if PME is currently installed
    If ($OSArch -like '*64*') {
        Write-Host "Checking if PME is already installed..." -ForegroundColor Cyan
        $PATHS = @("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")       
        $SOFTWARE = "SolarWinds MSP Patch Management Engine"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName,DisplayVersion,InstallDate
            
            If ($null -ne $installed) {    
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "SolarWinds MSP Patch Management Engine") {
                        $InstallDate = $($app.InstallDate)
                        $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                        $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        $IsPMEInstalled = "Yes"
                        Write-Host "PME Already Installed: Yes" -ForegroundColor Green
                        Write-Output "Installed PME Version: $($app.DisplayVersion)"
                        Write-Output "Installed PME Date: $InstallDateFormatted"
                    }
                }       
            }
            Else {
                $IsPMEInstalled = "No"
                Write-Host "PME Already Installed: No" -ForegroundColor Yellow 
            } 
        }
    }
    
    If ($OSArch -like '*32*') {
        Write-Host "Checking if PME is already installed..." -ForegroundColor Cyan
        $PATHS = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")       
        $SOFTWARE = "SolarWinds MSP Patch Management Engine"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName,DisplayVersion,InstallDate
    
            If ($null -ne $installed) {    
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "SolarWinds MSP Patch Management Engine") {
                        $InstallDate = $($app.InstallDate)
                        $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                        $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        $IsPMEInstalled = "Yes"
                        Write-Host "PME Already Installed: Yes" -ForegroundColor Green
                        Write-Output "Installed PME Version: $($app.DisplayVersion)"
                        Write-Output "Installed PME Date: $InstallDateFormatted"
                    }
                }       
            }
            Else {
                $IsPMEInstalled = "No"
                Write-Host "PME Already Installed: No" -ForegroundColor Yellow 
            } 
        }
    }   
}

Function Get-PMESetupDetails {
    # Declare static URI of PMESetup_details.xml
    If ($Fallback -eq "Yes") {
        $PMESetup_detailsURI = "http://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml" 
    }
    Else {
        $PMESetup_detailsURI = "https://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml"  
    }
    Try {
        $request = $null
        [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$PMESetup_detailsURI") -split '<\?xml.*\?>')[-1]
        $PMEDetails = $request.ComponentDetails
        $LatestVersion = $request.ComponentDetails.Version
    }
    Catch [System.Net.WebException] {
        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Error fetching PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
        Throw "Error fetching PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message)"
    }
    Catch [System.Management.Automation.MetadataException] {
        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Error casting to XML, could not parse PMESetup_details.xml, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
        Throw "Error casting to XML, could not parse PMESetup_details.xml, aborting. Error: $($_.Exception.Message)"
    }
    Catch {
        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Error occurred attempting to obtain PMESetup details, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
        Throw "Error occurred attempting to obtain PMESetup details, aborting. Error: $($_.Exception.Message)"
     }     
}

Function Get-PMEConfigurationDetails {
    # Declare static URI of PmeConfiguration_details.xml
     $PMEConfigurationDetailsURI = "https://sis.n-able.com/ComponentData/RMM/all/PmeConfiguration_details.xml"  
     Try {
         $request = $null
         [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$PMEConfigurationDetailsURI") -split '<\?xml.*\?>')[-1]
         $PMEConfigurationDetails = $request.ComponentDetails
         $PMEConfigurationDate = $PMEConfigurationDetails.Version
         $PMEConfigurationDate = $PMEConfigurationDate.Substring(0,$PMEConfigurationDate.Length-3)
         Write-Output "Latest PME Version: $LatestVersion"
         Write-Output "Latest PME Release Date: $PMEConfigurationDate"
     }
     Catch [System.Net.WebException] {
        Write-Output "Error fetching PMESetup_Details.xml check your source URL!"
        Throw
     }
     Catch [System.Management.Automation.MetadataException] {
        Write-Output "Error casting to XML; could not parse PMESetup_details.xml"
        Throw
     }   
 }

Function Confirm-PMEUpdatePending {
    # Check if PME is awaiting update for new release but has not updated yet (normally within 48 hours)
    $PMEWrapperFile = "$NcentralLog\PMEWrapper.log"

    If (Test-Path $PMEWrapperFile) {
        $PMEWrapper = get-content "$NcentralLog\PMEWrapper.log" 
        $NewVersion = "Newer PME version detected"
        $LastInstall = "Installing PME"
        If (($IsPMEInstalled -eq "Yes") -and ($PMEWrapper -ne "") -and ($PMEWrapper -match $NewVersion) -and ($PMEWrapper -match $LastInstall)) {
            $Date = Get-Date -Format 'yyyy.MM.dd'
            Write-Output "Current Date: $Date"
            $NewVersionMatch = ($PMEWrapper -match $NewVersion)[-1] 
            $LastInstallDateVersion = $NewVersionMatch.Split('')[10].Trim("(),")
            $LastInstallMatch = ($PMEWrapper -match $LastInstall)[-1]
            $LastInstallDate = $LastInstallMatch.Split('')[1].Trim() + " " + $LastInstallMatch.Split('')[2].Trim()
            Write-Host "INFO: Last automatic PME check detected version ($LastInstallDateVersion) and installed it on ($LastInstallDate)" -ForegroundColor Cyan
            $ConvertPMEConfigurationDate = Get-Date "$PMEConfigurationDate"
            $SelfHealingDate = $ConvertPMEConfigurationDate.AddDays($RepairAfterUpdateDays).ToString('yyyy.MM.dd')
            Write-Host "INFO: Repair-PME will only proceed ($RepairAfterUpdateDays) days after a new version of PME has been released" -ForegroundColor Cyan
            $DaysElapsed = (New-TimeSpan -Start $SelfHealingDate -End $Date).Days
            $DaysElapsedReversed = (New-TimeSpan -Start $Date -End $SelfHealingDate).Days

            # Only run if current $Date is greater than or equal to $SelfHealingDate and $LatestVersion is greater than $LastInstallDateVersion
            If (($Date -ge $SelfHealingDate) -and ($LatestVersion -ge $LastInstallDateVersion)) {
                Write-Output "($DaysElapsed) days has elapsed since a new version of PME has been released, script will proceed"
            }
            Else {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Only ($DaysElapsed) days has elapsed since a new version of PME has been released, aborting.`nScript: Repair-PME.ps1"
            Throw "Only ($DaysElapsedReversed) of ($RepairAfterUpdateDays) set days has elapsed since a new version of PME has been released, aborting."
            Break
            }
        }
        Else {
            Write-Host "WARNING: Skipping update pending check as PMEWrapper.log does not currently contain update information" -ForegroundColor Yellow
        } 
    }
    Else {
        Write-Host "WARNING: Skipping update pending check as PMEWrapper.log does not currently exist" -ForegroundColor Yellow    
    } 
}

Function Test-Connectivity {
    # Performs connectivity tests to destinations required for PME
    If (($PSVersionTable.PSVersion -ge "4.0") -and (!($OSVersion -match 'Windows 7')) -and (!($OSVersion -match '2008 R2'))) {
        Write-Host "Performing HTTPS connectivity tests for PME required destinations..." -ForegroundColor Cyan
        $List1= @("sis.n-able.com")
        $HTTPSError = @()
        $List1 | ForEach-Object {
            $Test1 = Test-NetConnection $_ -port 443
            If ($Test1.tcptestsucceeded -eq $True) {
                Write-Host "OK: Connectivity to https://$_ ($(($Test1).RemoteAddress.IpAddressToString)) established" -ForegroundColor Green
                $HTTPSError += "No"    
            }
            Else {
                Write-Host "ERROR: Unable to establish connectivity to https://$_ ($(($Test1).RemoteAddress.IpAddressToString))" -ForegroundColor Red
                $HTTPSError += "Yes"
            }
        }

        Write-Host "Performing HTTP connectivity tests for PME required destinations..." -ForegroundColor Cyan
        $HTTPError = @()
        $List2= @("sis.n-able.com","download.windowsupdate.com","fg.ds.b1.download.windowsupdate.com")
        $List2 | ForEach-Object {
            $Test1 = Test-NetConnection $_ -port 80
            If ($Test1.tcptestsucceeded -eq $True) {
                Write-Host "OK: Connectivity to http://$_ ($(($Test1).RemoteAddress.IpAddressToString)) established" -ForegroundColor Green
                $HTTPError += "No"
            }
            Else {
                Write-Host "ERROR: Unable to establish connectivity to http://$_ ($(($Test1).RemoteAddress.IpAddressToString))" -ForegroundColor Red
                $HTTPError += "Yes"
            }
        }

        If (($HTTPError[0] -like "*Yes*") -and ($HTTPSError[0] -like "*Yes*")) {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "ERROR: No connectivity to $($List2[0]) can be established, aborting.`nScript: Repair-PME.ps1"  
            Throw "ERROR: No connectivity to $($List2[0]) can be established, aborting"    
        }
        ElseIf (($HTTPError[0] -like "*Yes*") -or ($HTTPSError[0] -like "*Yes*")) {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "WARNING: Partial connectivity to $($List2[0]) established, falling back to HTTP.`nScript: Repair-PME.ps1"  
            Write-Host "WARNING: Partial connectivity to $($List2[0]) established, falling back to HTTP" -ForegroundColor Yellow
            $Fallback = "Yes"
        }

        If ($HTTPError[1] -like "*Yes*") {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "WARNING: No connectivity to $($List2[1]) can be established, you will be unable to download Microsoft Updates!`nScript: Repair-PME.ps1"  
            Write-Host "WARNING: No connectivity to $($List2[1]) can be established, you will be unable to download Microsoft Updates!" -ForegroundColor Red
        }

        If ($HTTPError[2] -like "*Yes*") {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "WARNING: No connectivity to $($List2[2]) can be established, you will be unable to download Windows Feature Updates!`nScript: Repair-PME.ps1"  
            Write-Host "WARNING: No connectivity to $($List2[2]) can be established, you will be unable to download Windows Feature Updates!" -ForegroundColor Red    
        }
    }
    Else {
        Write-Host "WARNING: Skipping connectivity tests as OS is Windows 7/Server 2008 R2 and/or Powershell 4.0 or above is not installed" -ForegroundColor Yellow
        $Fallback = "Yes"    
    }
}

Function Invoke-SolarwindsDiagnostics {
    # Invokes Solarwinds official diagnostics tool to capture logs for support
    $ZipPath = "/`"ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs/SWDiagnostics$(Get-Date -Format 'yyyyMMdd-hhmmss').zip`""
    If ($OSArch -like '*64*') {
        # 32-bit program files on 64-bit
        $SolarwindsDiagnosticsFolderPath = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")+"\SolarWinds MSP\PME\Diagnostics"
        $SolarwindsDiagnosticsExePath = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")+"\SolarWinds MSP\PME\Diagnostics\SolarwindsDiagnostics.exe"
        If (Test-Path $SolarwindsDiagnosticsExePath) {
            Write-Output "Solarwinds Diagnostics located at '$SolarwindsDiagnosticsExePath'"
            If (Test-Path "C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs") {
                Write-Output "Directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' already exists, no need to create directory"
            }
            Else {
                Try {
                    Write-Output "Directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' does not exist, creating directory"
                    New-Item -ItemType Directory -Path "C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs" -Force | Out-Null
                }
                Catch {
                    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to create directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' required for saving log capture. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
                    Throw "Unable to create directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' required for saving log capture. Error: $($_.Exception.Message)"
                }
            }     
            Write-Output "Starting Solarwinds Diagnostics"
            #Write-Output "DEBUG: Solarwinds Diagnostics started with:- Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait"
            Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait
            Write-Output "Solarwinds Diagnostics completed, file saved to 'C:\ProgramData\SolarWinds MSP\Repair-PME\Diagnostic Logs'"    
        }
        Else {
            Write-Host "WARNING: Unable to detect Solarwinds Diagnostics, skipping log capture" -ForegroundColor Yellow    
        }    
    }
    ElseIf ($OSArch -like '*32*') {
        # 32-bit program files on 32-bit
        $SolarwindsDiagnosticsFolderPath = [Environment]::GetEnvironmentVariable("ProgramFiles")+"\SolarWinds MSP\PME\Diagnostics"
        $SolarwindsDiagnosticsExePath = [Environment]::GetEnvironmentVariable("ProgramFiles")+"\SolarWinds MSP\PME\Diagnostics\SolarwindsDiagnostics.exe"
        If (Test-Path $SolarwindsDiagnosticsExePath) {
            Write-Output "Solarwinds Diagnostics located at '$SolarwindsDiagnosticsExePath'"
            If (Test-Path "C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs") {
                Write-Output "Directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs', no need to create directory"
            }
            Else {
                Try {
                    Write-Output "Directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' does not exist, creating directory"
                    New-Item -ItemType Directory -Path "C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs" -Force | Out-Null
                }
                Catch {
                    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to create directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' required for saving log capture. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
                    Throw "Unable to create directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' required for saving log capture. Error: $($_.Exception.Message)"
                }
            }     
            Write-Output "Starting Solarwinds Diagnostics"
            #Write-Output "DEBUG: Solarwinds Diagnostics started with:- Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait"
            Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait
            Write-Output "Solarwinds Diagnostics completed, file saved to 'C:\ProgramData\SolarWinds MSP\Repair-PME\Diagnostic Logs'"   
        }
        Else {
            Write-Host "WARNING: Unable to detect Solarwinds Diagnostics, skipping log capture" -ForegroundColor Yellow  
        }   
    }
    Else {
        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to detect processor architecture, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
        Throw "Unable to detect processor architecture, aborting. Error: $($_.Exception.Message)"
    }      
}

Function Stop-PMESetup {
    # Kill any running instances of PMESetup.exe to ensure that we can download & install successfully
    Write-Host "Checking if PMESetup is currently running..." -ForegroundColor Cyan
    $PMESetupRunning = Get-Process PMESetup* -ErrorAction SilentlyContinue
        If ($PMESetupRunning) {
            Write-Host "WARNING: PMESetup is currently running, terminating" -ForegroundColor Yellow
            $PMESetupRunning | Stop-Process -Force
        }
        Else {
            Write-Host "OK: PMESetup is not currently running, proceeding" -ForegroundColor Green
        }
    
    Write-Host "Checking if CacheServiceSetup is currently running..." -ForegroundColor Cyan
    $PMESetupRunning = Get-Process CacheServiceSetup* -ErrorAction SilentlyContinue
        If ($PMESetupRunning) {
            Write-Host "WARNING: CacheServiceSetup is currently running, terminating" -ForegroundColor Yellow
            $PMESetupRunning | Stop-Process -Force
        }
        Else {
            Write-Host "OK: CacheServiceSetup is not currently running, proceeding" -ForegroundColor Green
        }
    
    Write-Host "Checking if RPCServerServiceSetup is currently running..." -ForegroundColor Cyan
    $PMESetupRunning = Get-Process RPCServerServiceSetup* -ErrorAction SilentlyContinue
        If ($PMESetupRunning) {
            Write-Host "WARNING: RPCServerServiceSetup is currently running, terminating" -ForegroundColor Yellow
            $PMESetupRunning | Stop-Process -Force
        }
        Else {
            Write-Host "OK: RPCServerServiceSetup is not currently running, proceeding" -ForegroundColor Green
        }
    
    Write-Host "Checking if _iu14D2N.tmp instances are currently running..." -ForegroundColor Cyan
    $PMESetupRunning = Get-Process _iu* -ErrorAction SilentlyContinue
        If ($PMESetupRunning) {
            Write-Host "WARNING: _iu14D2N.tmp instances are currently running, terminating" -ForegroundColor Yellow
            $PMESetupRunning | Stop-Process -Force
        }
        Else {
            Write-Host "OK: _iu14D2N.tmp instances are not currently running, proceeding" -ForegroundColor Green
        }      
}   

Function Stop-PMEServices {
    $Service = "SolarWinds.MSP.PME.Agent.PmeService"
    $ServiceStatus = (Get-Service $Service -ErrorAction SilentlyContinue).Status
    $Process = "SolarWinds.MSP.PME.Agent"
    If (($ServiceStatus -eq "Running") -or ($ServiceStatus -eq "Stopping") -or ($ServiceStatus -eq "Suspended")) {
        Write-Host "$Service is $ServiceStatus, attempting to stop..." -ForegroundColor Cyan
        Stop-Service -Name $Service -Force
        $ServiceStatus = (Get-Service $Service -ErrorAction SilentlyContinue).Status
        If ($ServiceStatus -eq "Stopped") {
            Write-Host "OK: $Service service successfully stopped" -ForegroundColor Green    
        }
        Else {
            Write-Host "WARNING: $Service still running, temporarily disabling recovery and terminating" -ForegroundColor Yellow   
            #Set-Service -Name $Service -StartupType Disabled
            sc.exe failure "$Service" reset= 0 actions= // >null
            Stop-Process -Name $Process* -Force
            sc.exe failure "$Service" actions= restart/0/restart/0//0 reset= 0 >null       
        }
    }
    
    $Service = "SolarWinds.MSP.RpcServerService"
    $ServiceStatus = (Get-Service $Service -ErrorAction SilentlyContinue).Status
    $Process = "SolarWinds.MSP.RpcServerService"
    If (($ServiceStatus -eq "Running") -or ($ServiceStatus -eq "Stopping") -or ($ServiceStatus -eq "Suspended")) {
        Write-Host "$Service is $ServiceStatus, attempting to stop..." -ForegroundColor Cyan
        Stop-Service -Name $Service -Force
        $ServiceStatus = (Get-Service $Service -ErrorAction SilentlyContinue).Status
        If ($ServiceStatus -eq "Stopped") {
            Write-Host "OK: $Service service successfully stopped" -ForegroundColor Green    
        }
        Else {
            Write-Host "WARNING: $Service still running, temporarily disabling recovery and terminating" -ForegroundColor Yellow   
            #Set-Service -Name $Service -StartupType Disabled
            sc.exe failure "$Service" reset= 0 actions= // >null
            Stop-Process -Name $Process* -Force
            sc.exe failure "$Service" actions= restart/0/restart/0//0 reset= 0 >null              
        }
    }
    
    $Service = "SolarWinds.MSP.CacheService"
    $ServiceStatus = (Get-Service $Service -ErrorAction SilentlyContinue).Status
    $Process = "SolarWinds.MSP.CacheService"
    If (($ServiceStatus -eq "Running") -or ($ServiceStatus -eq "Stopping") -or ($ServiceStatus -eq "Suspended")) {
        Write-Host "$Service is $ServiceStatus, attempting to stop..." -ForegroundColor Cyan
        Stop-Service -Name $Service -Force
        $ServiceStatus = (Get-Service $Service -ErrorAction SilentlyContinue).Status
        If ($ServiceStatus -eq "Stopped") {
            Write-Host "OK: $Service service successfully stopped" -ForegroundColor Green    
        }
        Else {
            Write-Host "WARNING: $Service still running, temporarily disabling recovery and terminating" -ForegroundColor Yellow   
            #Set-Service -Name $Service -StartupType Disabled
            sc.exe failure "$Service" reset= 0 actions= // >null
            Stop-Process -Name $Process* -Force
            sc.exe failure "$Service" actions= restart/0/restart/0//0 reset= 0 >null                    
        }
    }
}

Function Clear-PME {
    # Cleanup Solarwinds MSP Cache Service root folder
    If (Test-Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService") {
        Try {   
            Write-Output "Performing cleanup of Solarwinds MSP Cache Service root folder"
            Remove-Item "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\*.*" -Force -Confirm:$false | Out-Null
        }
        Catch {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to cleanup 'C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\*.*' aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
            Throw "Unable to cleanup 'C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\*.*' aborting. Error: $($_.Exception.Message)"
        }    
    } 
    Else {
        Write-Output "Cleanup not required as Solarwinds MSP Cache Service root folder does not already exist"
    }
    # Cleanup Solarwinds MSP Cache Service cache folder
    If (Test-Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache") {
        Try {
            Write-Output "Performing cleanup of Solarwinds MSP Cache Service cache folder"
            Remove-Item "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache\*.*" -Force -Confirm:$false | Out-Null
        }
        Catch {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to cleanup C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache\*.*' aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
            Throw "Unable to cleanup C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache\*.*' aborting. Error: $($_.Exception.Message)"            
        }  
    } 
    Else {
        Write-Output "Cleanup not required as Solarwinds MSP Cache Service cache folder does not already exist"
    }
}

Function Get-PMESetup {
    # Download Setup
    If ($Fallback -eq "Yes") {
        $FallbackDownloadURL = ($PMEDetails.DownloadURL).Replace('https','http')
        Write-Output "Begin download of current $($PMEDetails.FileName) version $($PMEDetails.Version) from sis.n-able.com"
        Try {
            (New-Object System.Net.WebClient).DownloadFile("$($FallbackDownloadURL)","C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe")
        }
        Catch {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
            Throw "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message)"
        }
    }
    Else {
        Write-Output "Begin download of current $($PMEDetails.FileName) version $($PMEDetails.Version) from sis.n-able.com"
        Try {
            (New-Object System.Net.WebClient).DownloadFile("$($PMEDetails.DownloadURL)","C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe")
        }
        Catch {
            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
            Throw "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message)"
        }
    }     
}

Function Read-PMEConfig {
    # Check PME Config and inform of possible misconfigurations
    $CacheServiceConfigFile = "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\config\CacheService.xml"

    If (Test-Path "$CacheServiceConfigFile") {
        $CacheServiceConfig = Get-Content -Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\config\CacheService.xml"

        If ($CacheServiceConfig -match '<CanBypassProxyCacheService>false</CanBypassProxyCacheService>') {
            Write-Host "WARNING: Patch profile doesn't allow PME to fallback to external sources, if probe is not reachable PME may not work!" -ForegroundColor Yellow
        }
        ElseIf ($CacheServiceConfig -match '<CanBypassProxyCacheService>true</CanBypassProxyCacheService>') {
            Write-Host "INFO: Patch profile allows PME to fallback to external sources" -ForegroundColor Cyan
        }
        Else {
        Write-Host "WARNING: Unable to determine if patch profile allows PME to fallback to external sources" -ForegroundColor Yellow   
        }

        $CacheSize = ($CacheServiceConfig -match '<CacheSizeInMB>')[-1].Trim()
        $CacheSize = $CacheSize.Trim('<CacheSizeInMB>,</CacheSizeInMB>')

        If ($CacheServiceConfig -match '<CacheSizeInMB>10240</CacheSizeInMB>') {
            Write-Host "INFO: Cache Service is set to default cache size of 10240 MB" -ForegroundColor Cyan
        }
        Else {
            Write-Host "WARNING: Cache Service is not set to default cache size of 10240 MB (currently $CacheSize MB), PME may not work at expected!" -ForegroundColor Yellow
        }
    }
    Else {
        Write-Host "WARNING: Cache Service config file does not exist, skipping Cache Service settings checks" -ForegroundColor Yellow
    }
}    

Function Set-PMEConfig {
    # NCPM-4407 – This solution covers a specific case that can trigger PME to be Misconfigured with error "System.OutOfMemoryException". This issue seems to be caused by not having enough continuous blocks of free memory, even if there is plenty of memory available. This is usually why PME will work for an amount of time after a reboot of the device. This solution will instead keep PME in memory rather than unloading, to ensure resources are made available to it. To trigger this feature, the value of "UnloadModuleAppDomainAfterEachRequest" must be changed from "true" to "false" in C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.RpcServerService\config\RpcServerConfiguration.xml
    $RPCServerServiceConfigFile = "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.RpcServerService\config\RpcServerConfiguration.xml"

    If (Test-Path "$RPCServerServiceConfigFile") {
        $RpcServerServiceConfig = Get-Content -Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.RpcServerService\config\RpcServerConfiguration.xml"

        If ($RpcServerServiceConfig -match '<UnloadModuleAppDomainAfterEachRequest>false</UnloadModuleAppDomainAfterEachRequest>') {
            Write-Host "INFO: RPC Server Service configuration to address NCPM-4407 (System.OutOfMemoryException) is already applied" -ForegroundColor Cyan
        }    

        If (($NCPM4407 -eq "Yes") -and ($RpcServerServiceConfig -match '<UnloadModuleAppDomainAfterEachRequest>true</UnloadModuleAppDomainAfterEachRequest>')) {
            Try {
                Write-Output "Changing RPC Server Service configuration to address NCPM-4407 (System.OutOfMemoryException)"
                $RpcServerServiceConfig -replace "<UnloadModuleAppDomainAfterEachRequest>true</UnloadModuleAppDomainAfterEachRequest>","<UnloadModuleAppDomainAfterEachRequest>false</UnloadModuleAppDomainAfterEachRequest>" | Set-Content -Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.RpcServerService\config\RpcServerConfiguration.xml"
                }
            Catch {
                Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to change RpcServerService configuration to address NCPM-4407 (System.OutOfMemoryException). Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
                Throw "Unable to change RPC Server Service configuration to address NCPM-4407 (System.OutOfMemoryException). Error: $($_.Exception.Message)"
            }						  
        }
    }
    Else {
    Write-Host "WARNING: RPC Server Service config file does not exist, fix for NCPM-4407 can't be applied" -ForegroundColor Yellow
    }
}

Function Install-PME {
    # Check Setup Exists in PME Archive Directory
    If (Test-Path "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe") {
        # Check Hash
        Write-Output "Checking hash of local file at 'C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe'"
        $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe"
            If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                # Install
                Write-Output "Local copy of $($PMEDetails.FileName) is current and hash is correct"
                Write-Host "Installing $($PMEDetails.FileName) - logs will be saved to 'C:\ProgramData\Solarwinds MSP\Repair-PME\'" -ForegroundColor Cyan
                $Install = Start-process -FilePath "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe" -Argumentlist "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"C:\ProgramData\Solarwinds MSP\Repair-PME\Setup Log $DateTime.txt`"" -Wait -Passthru
                    If ($Install.ExitCode -eq 0) {
                        Write-Host "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed" -ForegroundColor Green
                    }
                    ElseIf ($Install.ExitCode -eq 5) {
                        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode).`nScript: Repair-PME.ps1"  
                        Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode)"
                    }
                    Else {
                        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'.`nScript: Repair-PME.ps1"  
                        Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'"
                    }
            }
            Else {
                # Download
                Write-Output "Hash of local file ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, downloading the latest available version"
                . Get-PMESetup   
                # Check Hash
                Write-Output "Checking hash of local file at 'C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe'"
                $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe"
                    If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                        # Install
                        Write-Output "Hash of file is correct"
                        Write-Host "Installing $($PMEDetails.FileName) - logs will be saved to 'C:\ProgramData\Solarwinds MSP\Repair-PME\'" -ForegroundColor Cyan
                        $Install = Start-process -FilePath "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe" -Argumentlist "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"C:\ProgramData\Solarwinds MSP\Repair-PME\Setup Log $DateTime.txt`"" -Wait -Passthru
                        If ($Install.ExitCode -eq 0) {
                            Write-Host "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed" -ForegroundColor Green
                        }
                        ElseIf ($Install.ExitCode -eq 5) {
                            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode).`nScript: Repair-PME.ps1"  
                            Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode)"
                        }
                        Else {
                            Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'.`nScript: Repair-PME.ps1"  
                            Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'"
                        }
                    }
                    Else {
                        Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting.`nScript: Repair-PME.ps1"  
                        Throw "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting"
                    }
            }
    }
    Else {
        Write-Output "$($PMEDetails.FileName) does not exist, begin download and install phase"
            # Check for PME Archive Directory
            If (Test-Path "C:\ProgramData\SolarWinds MSP\PME\archives") {
                Write-Output "Directory 'C:\ProgramData\SolarWinds MSP\PME\archives' already exists, no need to create directory"
            }
            Else {
                Try {
                    Write-Output "Directory 'C:\ProgramData\SolarWinds MSP\PME\archives' does not exist, creating directory"
                    New-Item -ItemType Directory -Path "C:\ProgramData\SolarWinds MSP\PME\archives" -Force | Out-Null
                }
                Catch {
                    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Unable to create directory 'C:\ProgramData\SolarWinds MSP\PME\archives' required for download, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"  
                    Throw "Unable to create directory 'C:\ProgramData\SolarWinds MSP\PME\archives' required for download, aborting. Error: $($_.Exception.Message)"
                } 
            }
        # Download
        . Get-PMESetup
        # Check Hash
        Write-Output "Checking hash of local file at 'C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe'"
        $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe"
            If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                # Install
                Write-Output "Hash of file is correct"
                Write-Host "Installing $($PMEDetails.FileName) - logs will be saved to 'C:\ProgramData\Solarwinds MSP\Repair-PME\'" -ForegroundColor Cyan
                $Install = Start-process -FilePath "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe" -Argumentlist "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"C:\ProgramData\Solarwinds MSP\Repair-PME\Setup Log $DateTime.txt`"" -Wait -Passthru
                If ($Install.ExitCode -eq 0) {
                    Write-Host "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed" -ForegroundColor Green
                }
                ElseIf ($Install.ExitCode -eq 5) {
                    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode).`nScript: Repair-PME.ps1"  
                    Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode)"
                }
                Else {
                    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'.`nScript: Repair-PME.ps1"  
                    Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'"
                }
            }
            Else {
                Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting.`nScript: Repair-PME.ps1"  
                Throw "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting"
            }
    }
}

Function Set-End {
    New-EventLog -LogName Application -Source "Repair-PME" -ErrorAction SilentlyContinue
    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Repair-PME has finished.`nScript: Repair-PME.ps1"
}

. Set-Start
. Confirm-Elevation
. Get-OSVersion
. Get-OSArch
. Get-PSVersion
. Confirm-PMEInstalled
. Get-PMESetupDetails
. Get-PMEConfigurationDetails
. Confirm-PMEUpdatePending
. Test-Connectivity
. Invoke-SolarwindsDiagnostics 
. Stop-PMESetup
. Stop-PMEServices
. Clear-PME
. Read-PMEConfig
. Set-PMEConfig
. Install-PME
. Set-End
