<#    
    ************************************************************************************************************
    Name: Install-TakeControl-Generic.ps1
    Version: 0.1 (30th September 2020)
    Purpose:    Download and Install TakeControl
    Pre-Reqs:    Powershell 2
    0.1 Initial Release of Generic version

    ************************************************************************************************************
#>

$SISXML='http://sis.n-able.com/GenericFiles.xml'
$LocalXML='C:\Program Files (x86)\N-able Technologies\Windows Agent\Temp\GenericFiles.xml'
$expectedversion = "7.0.17.1117"
$downloadrequired = $null
#$MSPInstallURL = 'http://sis.n-able.com/GenericFiles/MSPA/MSPAWINInstaller/6.77.26.0/MSPA4NCentral-7.00.17-20200715.exe'
$MSPInstallLocation = "$env:windir\temp\MSPA4NCentral.exe"

#region functions

# Detect TakeControl Version to download and URL
Function Get-TakeControlVersionURL {
    # Set Correct XML Location
    if (test-path 'C:\Program Files\N-able Technologies\Windows Agent') {
    $LocalXML='C:\Program Files\N-able Technologies\Windows Agent\Temp\GenericFiles.xml'
    }
            # Download XML File if it doesn't already exist
            if (!(test-path "$LocalXML")) {
                import-module bitstransfer
                Start-BitsTransfer -Source $SISXML -Destination $LocalXML
                }
            $mspa = get-content $LocalXML
            $location = "MSPA4NCentral-"
            # $array=($mspa -match $location).Split('"')
            $array0 = $mspa -match $location
            $array = ("$array0").split('"')
            $MSPInstallurl = $array[1]
            $MSPCVersion = $array[3]
            
            write-host "URL: " -nonewline; write-host "$MSPInstallurl" -foregroundcolor green
            write-host "Version: "-nonewline; write-host "$MSPCVersion" -foregroundcolor green 
    }

# Download TakeControl Installer
Function Get-TakeControlDownload {
if ($downloadrequired -eq $true) {

    Write-Host "Downloading TakeControl Installer"
    Import-Module BitsTransfer
    $start_time = Get-Date
    Write-Verbose -Message "Downloading: $($MSPInstallURL) to $($MSPInstallLocation)"
    Start-BitsTransfer -Source $MSPInstallURL -Destination $MSPInstallLocation
        If ($? -eq "True") {
        Write-Host "TakeControl Installer was succesfully downloaded in $((Get-Date).Subtract($start_time).Seconds) second(s)."
        }
        Else {
        Write-Host "There was an problem downloading the TakeControl Installer. Aborting Proceedings."
        }
    }
}

# Check if TakeControl Installer has already been downloaded
Function Check-DownloadedVersion {

        if (!(test-path $MSPInstallLocation)) {
            Write-Host "TakeControl Installer has not been detected. Prcoeeding to download..."
            $downloadrequired = $true
            }
            Else {
                Write-host "TakeControl Installer has been downloaded. Checking if up to date"
                $DownloadedVersion = (get-itemproperty $MSPInstallLocation).VersionInfo.fileVersion
                    if ($DownloadedVersion -eq $MSPCVersion) {
                    Write-Host "Downloaded Version is up to date..."
                    $downloadrequired = $false
                    }
                    else {
                    Write-Host "Downloaded Version is not up to date. Redownloading..."
                    remove-item $MSPInstallLocation -Force
                    $downloadrequired = $true
                    }

                }
}

# Install TakeControl
Function Install-TakeControl {
    # Running TakeControl Installer
    Write-host "Running TakeControl Installer..."
    Start-Process -filepath "$MSPInstallLocation" -argumentlist '/S /R' -Wait
}

# Check if TakeControl is installed now
function Confirm-Install {

        $TakeControlservice = Get-Service -Name BASupportExpressStandaloneService_N_Central -ErrorAction SilentlyContinue
        $TakeControlUpdateservice = Get-Service -Name BASupportExpressSrvcUpdater_N_Central -ErrorAction SilentlyContinue

        if (($TakeControlservice -eq $Null) -and ($TakeControlUpdateservice -eq $Null)) {
        Write-Host "There was a problem with the install. Check Install Log - C:\Windows\Temp\ServerAppletServiceInstaller-N-Central.install.log" -foregroundcolor Red
        }
        else {
        Write-Host "TakeControl has been succesfully installed." -foregroundcolor Green
        }
}

#endregion


# Check if TakeControl Service is present on device
$ServiceRunning = Get-Service -Name BASupportExpressStandaloneService_N_Central -ErrorAction SilentlyContinue
if ($ServiceRunning -eq $Null) {
    Write-Host "TakeControl is not installed..."

    . Get-TakeControlVersionURL
    . Check-DownloadedVersion
    . Get-TakeControlDownload
    . Install-TakeControl
    . Confirm-Install

    }
 
Else {
$fileversion = (get-item "C:\Program Files (x86)\BeAnywhere Support Express\GetSupportService_N-Central\BASupSrvc.exe").VersionInfo.ProductVersion
Write-Host "TakeControl "-nonewline; Write-Host ($fileversion) -foregroundcolor Green -nonewline; Write-Host " is already installed"
    if ($fileversion -ge $expectedversion){
    Write-Host "Upgrade is not required." -foregroundcolor Green
    }
    else {
    Write-Host "Upgrade is required." -foregroundcolor Red
    . Get-TakeControlVersionURL
    . Check-DownloadedVersion
    . Get-TakeControlDownload
    . Install-TakeControl
    . Confirm-Install
    }
}