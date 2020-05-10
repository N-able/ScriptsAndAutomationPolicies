<#    
   **************************************************************************************************************************
    Name:            Repair-PME.ps1
    Version:         0.1.4.1 (10/05/2020)
    Purpose:         Install/Reinstall Patch Management Engine (PME)
    Created by:      Ashley How
    Thanks to:       Jordan Ritz for initial Get-PMESetup function code. Thanks to Prejay Shah for input into script.
    Pre-Reqs:        Powershell 2.0
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
                     0.1.4.0   Update script to PSScriptAnalyzer best practices.
                               Rename Function 'Cleanup-PME' to 'Clear-PME' to conform to approved verbs for PS.
                               Rename Function 'Download-PME' to 'Get-PMESetup' to conform to approved verbs for PS.
                               Rename Function 'Get-PMESetup' to 'Get-PMESetupDetails' to release use for above function.
                               New function 'Stop-PMESetup' moving code from 'Clear-PME' function.
                               New function 'Set-Start' and 'Set-End' to write event log entries.
                     0.1.4.1   New function 'Invoke-SolarwindsDiagnostics' to capture logs for support prior to repair.                       
   **************************************************************************************************************************
#>
$Version = '0.1.4.1 (10/05/2020)'

Write-Output "Repair-PME $Version`n"

Function Set-Start {
    New-EventLog -LogName Application -Source "Repair-PME" -ErrorAction SilentlyContinue
    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Repair-PME has started, running version $Version.`nScript: Repair-PME.ps1"
}

Function Invoke-SolarwindsDiagnostics {
    # Invokes Solarwinds official diagnostics tool to capture logs for support
    $ZipPath = "/`"ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs/SWDiagnostics$(Get-Date -Format 'yyyyMMdd-hhmmss').zip`""
    $OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
    If ($OSArch -like '*64*') {
        # 32-bit program files on 64-bit
        $SolarwindsDiagnosticsFolderPath = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")+"\SolarWinds MSP\PME\Diagnostics"
        $SolarwindsDiagnosticsExePath = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")+"\SolarWinds MSP\PME\Diagnostics\SolarwindsDiagnostics.exe"
        If (Test-Path $SolarwindsDiagnosticsExePath) {
            Write-Output "Processor architecture is $OSArch, Solarwinds Diagnostics located at '$SolarwindsDiagnosticsExePath'"
            Write-Output "Starting Solarwinds Diagnostics"
            Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait
            Write-Output "Solarwinds Diagnostics completed, file saved to 'C:\ProgramData\SolarWinds MSP\Repair-PME\Diagnostic Logs'"    
        }
        Else {
            Write-Output "Unable to detect Solarwinds Diagnostics at '$SolarwindsDiagnosticsExePath', skipping log capture"    
        }    
    }
    ElseIf ($OSArch -like '*32*') {
        # 32-bit program files on 32-bit
        $SolarwindsDiagnosticsFolderPath = [Environment]::GetEnvironmentVariable("ProgramFiles")+"\SolarWinds MSP\PME\Diagnostics"
        $SolarwindsDiagnosticsExePath = [Environment]::GetEnvironmentVariable("ProgramFiles")+"\SolarWinds MSP\PME\Diagnostics\SolarwindsDiagnostics.exe"
        If (Test-Path $SolarwindsDiagnosticsExePath) {
            Write-Output "Processor architecture is $OSArch, Solarwinds Diagnostics located at '$SolarwindsDiagnosticsExePath'"
            Write-Output "Starting Solarwinds Diagnostics"
            Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait
            Write-Output "Solarwinds Diagnostics completed, file saved to 'C:\ProgramData\SolarWinds MSP\Repair-PME\Diagnostic Logs'"   
        }
        Else {
            Write-Output "Unable to detect Solarwinds Diagnostics at '$SolarwindsDiagnosticsExePath', skipping log capture"    
        }   
    }
    Else {
        Throw "Unable to detect processor architecture, aborting. Error: $($_.Exception.Message)"
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
        Throw "Unable to performing hashing, aborting. Error: $($_.Exception.Message)"
    }
}

Function Get-PMESetupDetails {
   # Declare static URI of PMESetup_details.xml
    $PMESetup_detailsURI = "https://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml"  
    Try {
        $request = $null
        [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$PMESetup_detailsURI") -split '<\?xml.*\?>')[-1]
        $PMEDetails = $request.ComponentDetails
    }
    Catch [System.Net.WebException] {
        Throw "Error fetching PMESetup_Details.xml, check your source URL, aborting. Error: $($_.Exception.Message)"
    }
    Catch [System.Management.Automation.MetadataException] {
        Throw "Error casting to XML, could not parse PMESetup_details.xml, aborting. Error: $($_.Exception.Message)"
    }
    Catch {
        Throw "Error occurred attempting to obtain PMESetup details, aborting. Error: $($_.Exception.Message)"
     }     
}

Function Stop-PMESetup {
    # Kill any running instances of PMESetup.exe to ensure that we can download & install successfully
    Write-Output "Checking if PMESetup is currently running"
    $PMESetupRunning = Get-Process PMESetup* -ErrorAction SilentlyContinue
        If ($PMESetupRunning) {
            Write-Output "PMESetup is currently running, forcefully terminating"
            $PMESetupRunning | Stop-Process -Force
        }
        Else {
            Write-Output "PMESetup is not currently running, proceeding"
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
            Throw "Unable to cleanup C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache\*.*' aborting. Error: $($_.Exception.Message)"            
        }  
    } 
    Else {
        Write-Output "Cleanup not required as Solarwinds MSP Cache Service cache folder does not already exist"
    }
}

Function Get-PMESetup {
    # Download Setup
    Write-Output "Begin download of current $($PMEDetails.FileName) version $($PMEDetails.Version) from sis.n-able.com"
    Try {
        (New-Object System.Net.WebClient).DownloadFile("$($PMEDetails.DownloadURL)","C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)")
    }
    Catch {
        Throw "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message)"
    } 
}

Function Install-PME {
    # Check Setup Exists in PME Archive Directory
    If (Test-Path "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)") {
        # Check Hash
        Write-Output "Checking hash of local file at 'C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)'"
        $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)"
            If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                # Install
                Write-Output "Local copy of $($PMEDetails.FileName) is current and hash is correct, installing"
                $Install = Start-process -FilePath "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)" -Argumentlist '/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Wait -Passthru
                    If ($Install.ExitCode -eq 0) {
                        Write-Output "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed"
                    }
                    Else {
                        Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode)"
                    }
            }
            Else {
                # Download
                Write-Output "Hash of local file ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, downloading the latest available version"
                . Get-PMESetup   
                # Check Hash
                Write-Output "Checking hash of local file at 'C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)'"
                $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)"
                    If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                        # Install
                        Write-Output "Hash of file is correct, installing $($PMEDetails.FileName)"
                        $Install = Start-process -FilePath "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)" -Argumentlist '/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Wait -Passthru
                            If ($Install.ExitCode -eq 0) {
                                Write-Output "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed"
                            }
                            Else {
                                Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode)"
                            }
                    }
                    Else {
                        Throw "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting"
                    }
            }
    }
    Else {
        Write-Output "$($PMEDetails.FileName) does not exist, begin download and install phase"
            # Check for PME Archive Directory
            If (Test-Path "C:\ProgramData\SolarWinds MSP\PME\archives") {
                Write-Output "Directory C:\ProgramData\SolarWinds MSP\PME\archives already exists, no need to create directory"
            }
            Else {
                Try {
                    Write-Output "Directory 'C:\ProgramData\SolarWinds MSP\PME\archives' does not exist, creating directory"
                    New-Item -ItemType Directory -Path "C:\ProgramData\SolarWinds MSP\PME\archives" -Force | Out-Null
                }
                Catch {
                    Throw "Unable to create directory 'C:\ProgramData\SolarWinds MSP\PME\archives' required for download, aborting. Error: $($_.Exception.Message)"
                } 
            }
        # Download
        . Get-PMESetup
        # Check Hash
        Write-Output "Checking hash of local file at 'C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)'"
        $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)"
            If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                # Install
                Write-Output "Hash of file is correct, installing $($PMEDetails.FileName)"
                $Install = Start-process -FilePath "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)" -Argumentlist '/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Wait -Passthru
                    If ($Install.ExitCode -eq 0) {
                        Write-Output "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed"
                    }
                    Else {
                        Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode)"
                    }
            }
            Else {
                Throw "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting"
            }
    }
}

Function Set-End {
    New-EventLog -LogName Application -Source "Repair-PME" -ErrorAction SilentlyContinue
    Write-EventLog -LogName Application -Source "Repair-PME" -EntryType Information -EventID 100 -Message "Repair-PME has finished.`nScript: Repair-PME.ps1"
}

. Set-Start
. Invoke-SolarwindsDiagnostics 
. Get-PMESetupDetails
. Stop-PMESetup
. Clear-PME
. Install-PME
. Set-End
