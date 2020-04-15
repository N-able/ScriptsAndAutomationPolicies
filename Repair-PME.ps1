<#    
   **********************************************************************************************************************
    Name:            Repair-PME.ps1
    Version:         0.1.2.0 (15/04/2020)
    Purpose:         Install/Reinstall Patch Management Enginge (PME)
    Created by:      Ashley How
    Thanks to:       Jordan Ritz for inital Get-PMESetup function code
    Pre-Reqs:        Powershell 2.0
    Version History: 0.1.0.0 - Initial Release.
                     0.1.1.0 - Update PMESetup_details.xml URL, update install arguments, create new cleanup function 
                               as suggested by Jan Tauwinkl at Solarwinds. Thanks for your input. Rename Function
                               Install-PMESetup to Install-PME.
                     0.1.1.1 - Update Get-PMESetup function for better error detection when unable to grab
                               PMESetup_Details.xml. Thanks to Jordan Ritz.       
                     0.1.2.0 - Update Get-PMESetup function for PS 2.0 support. Import of BitsTransfer Module for PS 2.0.               
   **********************************************************************************************************************
#>
$Version = '0.1.2.0 (15/04/2020)'

Write-Output "Repair-PME $Version`n"

function Get-LegacyHash {
    Param($Path)

    $csp = new-object -TypeName System.Security.Cryptography.SHA256CryptoServiceProvider
    $ComputedHash = @()
    $ComputedHash = $csp.ComputeHash([System.IO.File]::ReadAllBytes($Path))
    $ComputedHash = [System.BitConverter]::ToString($ComputedHash).Replace("-", "").ToLower()
    Return $ComputedHash
}

Function Get-PMESetup {
   # Delcare static URI of PMESetup_details.xml
    $PMESetup_detailsURI = "https://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml"  
    Try {
        $request = $null
        [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$PMESetup_detailsURI") -split '<\?xml.*\?>')[-1]
        $PMEDetails = $request.ComponentDetails
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

Function Cleanup-PME {
   # Cleanup Solarwinds MSP Cache Service root folder
   If (Test-Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService") {
      Write-Output "Performing cleanup of Solarwinds MSP Cache Service root folder"
      Remove-Item "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\*.*" -Force -Confirm:$false
   } 
   Else {
      Write-Output "Cleanup not required as Solarwinds MSP Cache Service root folder does not already exist"
   }
   
   # Cleanup Solarwinds MSP Cache Service cache folder
   If (Test-Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache") {
      Write-Output "Performing cleanup of Solarwinds MSP Cache Service cache folder"
      Remove-Item "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache\*.*" -Force -Confirm:$false
   } 
   Else {
      Write-Output "Cleanup not required as Solarwinds MSP Cache Service cache folder does not already exist"
   }
}

Function Install-PME {
    # Check Setup Exists in PME Archive Directory
    If (Test-Path "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)") {
        # Check Hash
        $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)"
            If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                # Install
                Write-Output "Local copy of $($PMEDetails.FileName) is current and hash is correct, installing"
                $Install = Start-process -filepath "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)" -argumentlist '/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -wait -passthru
                    If ($Install.ExitCode -eq 0) {
                        Write-Output "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed"
                    }
                    Else {
                        Write-Error "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode)"
                    }
            }
            Else {
                # Download Setup
                Write-Output "Hash of local file ($($Download.Hash)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.nable.com, downloading the latest available version"
                Write-Output "Begin download of current $($PMEDetails.FileName) version $($PMEDetails.Version) from sis.a-able.com"
                # Load module to ensure it works on Powershell 2.0 devices
                Import-Module BitsTransfer
                Start-BitsTransfer -Source "$($PMEDetails.DownloadURL)" -Destination "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)"
                
                # Check Hash
                $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)"
                    If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                        # Install
                        Write-Output "Hash of file is correct, installing $($PMEDetails.FileName)"
                        $Install = Start-process -filepath "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)" -argumentlist '/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -wait -passthru
                            If ($Install.ExitCode -eq 0) {
                                Write-Output "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed"
                            }
                            Else {
                                Write-Error "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode)"
                            }
                    }
                    Else {
                        Write-Error "Hash of file downloaded ($($Download.Hash)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.nable.com, aborting install"
                    }
            }
    }
    Else {
        Write-Output "$($PMEDetails.FileName) does not exist, begin download and install phase"

            # Check for PME Archive Directory
            If (Test-Path "C:\ProgramData\SolarWinds MSP\PME\archives") {
                Write-Output "C:\ProgramData\SolarWinds MSP\PME\archives already exists"
            }
            Else {
                Write-Output "Directory 'C:\ProgramData\SolarWinds MSP\PME\archives' does not exist, creating directory"
                New-Item -ItemType Directory -Path "C:\ProgramData\SolarWinds MSP\PME\archives" -Force | Out-Null
            }

        # Download Setup
        Write-Output "Begin download of current $($PMEDetails.FileName) version $($PMEDetails.Version) from sis.a-able.com"
        # Load module to ensure it works on Powershell 2.0 devices
        Import-Module BitsTransfer
        Start-BitsTransfer -Source "$($PMEDetails.DownloadURL)" -Destination "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)"
        
        # Check Hash
        $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)" -Algorithm SHA256
            If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                # Install
                Write-Output "Hash of file is correct, installing $($PMEDetails.FileName)"
                $Install = Start-process -filepath "C:\ProgramData\SolarWinds MSP\PME\archives\$($PMEDetails.FileName)" -argumentlist '/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -wait -passthru
                    If ($Install.ExitCode -eq 0) {
                        Write-Output "$($PMEDetails.Name) version $($PMEDetails.Version) successfully installed"
                    }
                    Else {
                        Write-Error "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode)"
                    }
            }
            Else {
                Write-Error "Hash of file downloaded ($($Download.Hash)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.nable.com, aborting install"
            }
    }
}

. Get-PMESetup
. Cleanup-PME
. Install-PME
