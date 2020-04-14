<#    
    ******************************************************************************************************************
    Name:        Repair-PME.ps1
    Version:     0.1.0.1 (14/04/2020)
    Purpose:     Install/Reinstall Patch Management Enginge (PME)
    Created by:  Ashley How
    Thanks:      Jordan Ritz for Get-PMESetup function code
    Pre-Reqs:    Powershell 3.0 (2.0 is possible - alternative to Invoke-RestMethod in Get-PMESetup function required)
    Version History: 0.1.0.0 - Initial Release
                     0.1.0.1 - Update PMESetup_details.xml URL, update install arguments 
                               as requested by Jan Tauwinkl at Solarwinds
******************************************************************************************************************
#>
$Version = '0.1.0.1 (14/04/2020)'

Write-Output "Repair-PME $Version`n"

function Get-LegacyHash {
    param($Path)

    $csp = new-object -TypeName System.Security.Cryptography.SHA256CryptoServiceProvider
    $ComputedHash = @()
    $ComputedHash = $csp.ComputeHash([System.IO.File]::ReadAllBytes($Path))
    $ComputedHash = [System.BitConverter]::ToString($ComputedHash).Replace("-", "").ToLower()
    return $ComputedHash
}

Function Get-PMESetup {
    [xml]$x = ((Invoke-RestMethod https://sis.n-able.com/ComponentData/RMM/1/PMESetup_details.xml) -split '<\?xml.*\?>')[-1]
    $PMEDetails = $x.ComponentDetails
}

Function Install-PMESetup {
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
                Write-Output "Hash of local file ($($Download.Hash)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.nable.com, downloading the latest available version"
                Write-Output "Begin download of current $($PMEDetails.FileName) version $($PMEDetails.Version) from sis.a-able.com"
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
. Install-PMESetup
