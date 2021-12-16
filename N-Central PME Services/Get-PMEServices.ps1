<#*****************************************************************************************************

    Name: Get-PMEServices.ps1
    Version: 0.2.2.9 (15/11/2021)
    Author: Prejay Shah (Doherty Associates)
    Thanks To: Ashley How
    Purpose:    Get N-Central Patch Management Engine (PME) Status Details
    Use in combination with Ashley How's Repair-PME to diagnose/fix/upgrade PME when automatic updates won't work.

    Pre-Reqs:   PowerShell 2.0 (PowerShell 4.0+ for Connectivity Tests)

    Optional:   Repair-PME: https://github.com/N-able/ScriptsAndAutomationPolicies/tree/master/Repair-PME
                Repair-PME can be used in a standalone or self-healing capacity to perform diagnostics 
                on PME or upgrade PME outside of the autoamtic update function.

    Version History:    0.1.0.0 - Initial Release.
                                + Improved Detection for PME Services being missing on a device
                                + Improved Detection of Latest PME Version
                                + Improved Detection and error handling for latest Public PME Version when 
                                PME 1.1 is installed or PME is not installed on a device
                                + Improved Compatibility of PME Version Check
                                + Updated to include better PME 1.2.x Latest Versioncheck
                                + Updated for better PS 2.0 Compatibility
                        0.1.5.0 + Added PME Profile Detection, and Diagnostics information courtesy of 
                        Ashley How's Repair-PME
                        0.1.5.0a + N-Central AMP Variant using AMP Input Parameter to control Forcing of 
                        Diagnostics Mode (Disabled by Default)	
                        0.1.5.1 + Improved TLS Support, Updated Error Message for PME connectivity Test not 
                        working on Windows 7
                        0.1.5.2 + PME 1.2.4 has been made GA for the default stream so have had to alter detection methods
                        0.1.5.3 + Improved Compatibility with Server 2008 R2
                        0.1.5.4 + Updated 'Test-PMEConnectivity' function to fix a message typo. 
                        Thanks for Clint Conner for finding. 
                        0.1.6.0 + Have Added in PME Installer Log Analysis for use when PME is not up to date
                        0.1.6.1 + Have added date analysis of log file/detection/installation proceedings
                        0.1.6.2 + Fixed Detection Logic for when there has been no PME Scan on a device, 
                        and missing components
                        0.1.6.3 + Added Reading in of PME Config for Cache settings
                        0.1.6.4 + Added in better x86/x64 compatability because apparently there are still 32-bit OS devices out there.
                        0.1.6.5 + Fixed Typo
                        0.1.6.6 + [Ashley How] migrated code from my Repair-PME script giving it the abilty to no longer consider a pending update of PME a failure.
                                  Please Note: By default the grace period has been set at 2 days but can be changed by changing the $PendingUpdateDays in settings section.
                                + Various updates to functions and parts of the script to be closer in-line with Repair-PME script.
                                + Updated Validate-PME function to account for pending update, this will report a status of 0. Renamed status messages to make it eaiser to theshold in an AMP service.
                                + Fixed minor issues with date-time parsing, valid PS/OS detection, URL Querying 
                        0.1.6.7 + Added Offine Scanning Detection
                        0.1.6.8 + Updated Cache Size variable extraction from XML Config File (Thanks to Clayton Murphy for identifying this)
                        0.1.6.9 + Upated Cache Fallback detection from XML file as I found that some devices semeed to have corrupted XML config files.
                        0.1.6.10 + Updated Profile Options to Default/Insiders as it seems that SW have retired the alpha moniker.
                        0.1.6.11 + Improved Logging method for when PME installation Details cannot be found.
                        0.1.6.12 + Updated PME OS Requirement Checks to cater for older OS not being supported                       
                        0.1.7.0 + [Ashley How] Updated Get-PMESetupDetails function to be in line with latest Repair-PME script.
                                + [Ashley How] Removed Get-PMEConfigurationDetails function, code merged into Get-PMESetupDetails function.   
                                + [Ashley How] Updated Confirm-PMEInstalled function to be in line with latest Repair-PME script.
                                + [Ashley How] Updated Confirm-PMEUpdatePending function to be in line with latest Repair-PME script.
                                + [Ashley How] Updated Get-PMEConfigMisconfigurations function to be in line with latest Repair-PME script.
                                + [Ashley How] Fixed issue in Get-PMEAnalysis function where match comparision operators would not return $true or $false against the $PMEQueryLogContent variable.
                                + [Ashley How] Fixed some minor spacing issues in output. Updated script title formatting so it is more prominent. 
                                + [Ashley How] Changed date formating to dd/MM/yyyy for $Version variable and release notes.
                                + [Ashley How] Updated Get-PMEProfile function for more consistent formatting. Offline scanning enablement will no longer report if PME is not installed.         
                        0.1.7.1 + Updated PME Insider version detection string, Have changed 64bit OS detection method       
                        0.2.0.0 + Updated for Unexpected PME 2.0 release; Cleaned up registry application detection method
                        0.2.0.1 + Slight Tweaks for PMe 1.3.1 Compatibility
                        0.2.0.2 + Tweak Expectation for Insider Profile as versions no longer match up
                        0.2.0.3 + Tweak Cache XML Config parsing for PMe 2.0
                        0.2.1.0 + Modified Status Message Output to include timestamps
                        0.2.1.1 + Updated Placeholder for PME 2.0.1 
                        0.2.2.0 + Using Community XML as data source for PME Information instead of placeholder while we wait to see if anything can be done with official SW sources
                        0.2.2.1 + Cleanup Formatting and Typo's
                        0.2.2.2 + Update for compatibility with version expectation when using legacy PME
                        0.2.2.3 + Update OS Compatibility Output, Status/Version Output
                        0.2.2.4 + Changed Legacy PME Detection Method
                        0.2.2.5 + Converted from Throw to Write-Host for AMP compatibility, Hardcoded Legacy PME Release Date for devices that cannot access the website
                        0.2.2.6 + Updating 32-bit OS compatibility with Legacy and 2.x PME
                        0.2.2.7 + Updated for PME 2.1 Testing and minor information output improvements. 
                        Pending Update Fix courtesy of Ashley How
                        0.2.2.8 + Updated Parameter comment to help those who try to throw this script directly into AM's "Run Powershell Script" object
                        0.2.2.9 + Updated for better Windows 11 Detection/compatibility

    Examples: 
    Diagnostics Input: False
    Runs Normally and only engages diagnostcs for connectibity testing if PME is not up to date or missing a service.
    
    Diagnostics Input: True
    Force Diagnostics Mode to be enabled on the run regardless of PME Status

*****************************************************************************************************#>

# N-Able Automation Manager does not support the use of PS parameters within the "Run Powershell Script" object
# Comment out the paramter below if you're trying to run this PS script within AM.

Param (
        [Parameter(Mandatory=$false,Position=1)]
        [switch] $Diagnostics
    )

# Settings
# *****************************************************************************************************
# Change this variable to number of days (must be a number!) to consider a new version of PME as pending an update. Default is 2.
$PendingUpdateDays = "2"
# *****************************************************************************************************

#ddMMyy
$Version = '0.2.2.9 (15/11/2021)'
$EventLogCompanyName ="Doherty Associates"
$winbuild = $null
$osvalue = $null
$osbuildversion = $null
$RecheckStartup = $Null
$RecheckStatus = $Null
$request = $null
$Latestversion = $Null
$pmeprofile = $null
$diagnosticserrorint = $null
$pmeinstalllogcontent = $null
$PMEExpectationSetting = $False

$legacyPMEReleaseDate = "2021.01.27"
$legacyPME = $null

# $NAblePMESetup_detailsURIHTTPS = 'https://api.us-west-2.prd.patch.system-monitor.com/api/v1/pme/version/default'
# N-Able URL doens't include individual component versions or release data so we use own own URL instead

$CommunityPMESetup_detailsURIHTTP = "http://raw.githubusercontent.com/N-able/CustomMonitoring/master/N-Central%20PME%20Services/Community_PMESetup_details.xml"
$CommunityPMESetup_detailsURIHTTPS = "https://raw.githubusercontent.com/N-able/CustomMonitoring/master/N-Central%20PME%20Services/Community_PMESetup_details.xml"
$LegacyPMESetup_detailsURIHTTPS = "https://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml"
$LegacyPMESetup_detailsURIHTTP = "http://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml"

Write-Host ""
Write-Host "Get-PMEServices $Version" -ForegroundColor Cyan
Write-Host ""

if ($Diagnostics -eq 'True'){
    Write-Host "Diagnostics Mode Enabled" -foregroundcolor Yellow
}

#region Functions

Function Test-PMERequirement {
$winbuild = (Get-WmiObject -class Win32_OperatingSystem).Version
# [string]$WinBuild=[System.Environment]::OSVersion.Version
$UBR = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR).UBR
$OSBuildVersion = $winbuild + "." + $UBR 
write-Host "Windows Build Version: " -nonewline; Write-Host "$osbuildversion" -ForegroundColor Green

$OSName = (Get-WmiObject Win32_OperatingSystem).Caption
Write-Host "OS: " -nonewline; Write-Host "$OSName" -ForegroundColor Green
    if (($osname -match "XP") -or ($osname -match "Vista")  -or ($osname -match "Home") -or ($osname -match "2003") -or (($osname -match "2008") -and ($osname -notmatch "2008 R2")) ) {
        
        $Continue = $False
        $PMECacheStatus = "N/A - N-Central PME does not support OS"
        $PMEAgentStatus = "N/A - N-Central PME does not support OS"
        $PMERpcServerStatus = "N/A - N-Central PME does not support OS"
        $PMECacheVersion = '0.0'
        $PMEAgentVersion = '0.0'
        $PMERpcServerVersion = '0.0'
        pmeprofile = 'Default'
        $diagnosticserrorint = '2'
        $OverallStatus = '2'
        $StatusMessage = "$(Get-Date) - Error: The OS running on this device ($OSName $osbuildversion) is not supported by N-Central PME."
        $installernotes = $StatusMessage

        Write-Host "$StatusMessage" -ForegroundColor Red

    }
    else {
        $statusmessage = "$(Get-Date) - Information: The OS running on this device ($OSName $osbuildversion) is supported by N-Central PME"
        $installernotes = $statusmessage
        Write-Host "$statusmessage" -ForegroundColor Green
        Write-Host ""
        $Continue = $True

        # See: https://chocolatey.org/docs/installation#completely-offline-install
         # Attempt to set highest encryption available for SecurityProtocol.
        # PowerShell will not set this by default (until maybe .NET 4.6.x). This
        # will typically produce a message for PowerShell v2 (just an info message though)
        try {
            # Set TLS 1.2 (3072), then TLS 1.1 (768), then TLS 1.0 (192), finally SSL 3.0 (48)
            # Use integers because the enumeration values for TLS 1.2 and TLS 1.1 won't
            # exist in .NET 4.0, even though they are addressable if .NET 4.5+ is
            # installed (.NET 4.5 is an in-place upgrade).
            [System.Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768 -bor 192 -bor 48

            #Hardcoding usage of TLS 1.2
            #[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        } 
        catch {
            Write-Host 'Unable to set PowerShell to use TLS 1.2 and TLS 1.1 due to old .NET Framework installed. If you see underlying connection closed or trust errors, you may need to upgrade to .NET Framework 4.5+ and PowerShell v3+.' -ForegroundColor Red
        }

        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

        }  
}  

Function Set-PMEExpectations {

#$OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
$64bitOS = [System.Environment]::Is64BitOperatingSystem
if ($64bitOS -eq $true) {
    Write-Host "64-Bit OS Detected" -ForegroundColor Cyan
    $OSArch = "64-bit"
    $UninstallRegLocation = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
}
else {
    Write-Host "32-Bit OS Detected" -ForegroundColor Cyan
    $OSArch = "32-Bit"
    $UninstallRegLocation = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
}


    if (test-path $env:programdata\MspPlatform\PME\log) {
        $MSPPlatformCoreLogSize = (get-item "$env:programdata\MspPlatform\PME\log\core.log").length
        }
        
        if ($MSPPlatformCoreLogSize -gt '0') {
        $legacyPME = $false
        # N-Able have pre-announced PME via https://status.n-able.com/release-notes/ although there is no direct category for it
        Write-Host "Warning: PME 2.x Detected - Artificially setting version and release date expectation via Community XML" -ForegroundColor Yellow
        #$PMEExpectationSetting = $True

        #Write-Host "PME Latest Version: $PME20LatestVersionPlaceholder" -ForegroundColor Yellow
        #Write-Host "PME Release Date: $PME20ReleaseDatePlaceholder" -ForegroundColor Yellow

        $PMEProgramDataFolder = "$env:programdata\MspPlatform\PME"
        If ($64bitOS -eq $true) {
        $PMEProgramFilesFolder = "${Env:ProgramFiles(x86)}\MspPlatform"
        }
        else {
            $PMEProgramFilesFolder = "$Env:ProgramFiles\MspPlatform"
        }
        $PMEAgentExe = "PME\PME.Agent.exe"  
        $PMECacheExe = "FileCacheServiceAgent\FileCacheServiceAgent.exe"
        $PMERPCExe = "RequestHandlerAgent\RequestHandlerAgent.exe"
        $PMEAgentServiceName = "PME.Agent.PmeService"
        $PMECacheServiceName = "SolarWinds.MSP.CacheService"
        $PMERPCServiceName = "SolarWinds.MSP.RpcServerService"

        $PMEAgentAppName =  "Patch Management Service Controller"
        $PMECacheAppName = "File Cache Service Agent"
        $PMERPCAppName = "Request Handler Agent"

        $CacheServiceConfigFile = "$env:programdata\MspPlatform\Filecacheserviceagent\config\FileCacheServiceAgent.xml"
        $PMESetup_detailsURI = $CommunityPMESetup_detailsURIHTTPS
        
    }
    else {
        $legacyPME = $true
        $PMEProgramDataFolder = "$env:programdata\SolarWinds MSP\PME"
        If ($64bitOS -eq $true) {
        $PMEProgramFilesFolder = "${Env:ProgramFiles(x86)}\SolarWinds MSP"
        }
        else {
            $PMEProgramFilesFolder = "$Env:ProgramFiles\SolarWinds MSP"            
        }
        $PMEAgentExe = "PME\SolarWinds.MSP.PME.Agent.exe"
        $PMECacheExe = "CacheService\SolarWinds.MSP.CacheService.exe"
        $PMERPCExe = "RpcServer\SolarWinds.MSP.RpcServerService.exe"
        
        $PMEAgentServiceName = "SolarWinds.MSP.PME.Agent.PmeService"
        $PMECacheServiceName = "SolarWinds.MSP.CacheService"
        $PMERPCServiceName = "SolarWinds.MSP.RpcServerService"

        $PMEAgentAppName = "SolarWinds MSP Patch Management Engine"
        $PMECacheAppName = "SolarWinds MSP Cache Service"
        $PMERPCAppName = "Solarwinds MSP RPC Server"

        $CacheServiceConfigFile = "$env:programdata\SolarWinds MSP\PME\SolarWinds.MSP.CacheService\config\CacheService.xml"
        $PMESetup_detailsURI = $LegacyPMESetup_detailsURIHTTPS
    }


    
        If ($64bitOS -eq $true) {
            $SolarWindsMSPCacheLocation = "$PMEProgramFilesFolder\$PMECacheExe"
            $SolarWindsMSPPMEAgentLocation = "$PMEProgramFilesFolder\$PMEAgentExe"
            $SolarWindsMSPRpcServerLocation = "$PMEProgramFilesFolder\$PMERPCExe"
            $NCentralLog = "c:\Program Files (x86)\N-able Technologies\Windows Agent\log"
        }
        else {
            $SolarWindsMSPCacheLocation = "$PMEProgramFilesFolder\$PMECacheExe"
            $SolarWindsMSPPMEAgentLocation = "$PMEProgramFilesFolder\$PMEAgentExe"
            $SolarWindsMSPRpcServerLocation = "$PMEProgramFilesFolder\$PMERPCExe"
            $NCentralLog = "c:\Program Files\N-able Technologies\Windows Agent\log"
        }
}

Function Get-PMEServicesStatus {
$PMEAgentStatus = (get-service $PMEAgentServiceName -ErrorAction SilentlyContinue).Status
$PMECacheStatus = (get-service $PMECacheServiceName -ErrorAction SilentlyContinue).Status
$PMERpcServerStatus = (get-service $PMERPCServiceName -ErrorAction SilentlyContinue).status
}

Function Get-PMESetupDetails {
    <#
    if ($PMEExpectationSetting -eq $true) {
        $LatestVersion = $PME20LatestVersionPlaceholder
        $PMEReleaseDate = $PME20ReleaseDatePlaceholder
    }
    else {
    #>
    # Determine URI used for PMESetup_details.xml
        If ($Fallback -eq "Yes") {
            if ($legacyPME -eq $true) {
                $PMESetup_detailsURI = $LegacyPMESetup_detailsURIHTTP          
            }
            else {
                $PMESetup_detailsURI = $CommunityPMESetup_detailsURIHTTP
            }
        } Else {
            if ($legacyPME -eq $true) {
                $PMESetup_detailsURI = $LegacyPMESetup_detailsURIHTTPS
            }
            else {
                $PMESetup_detailsURI = $CommunityPMESetup_detailsURIHTTPS
            }
        }

        Try {
            $PMEDetails = $null
            $request = $null
            [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$PMESetup_detailsURI") -split '<\?xml.*\?>')[-1]
            $PMEDetails = $request.ComponentDetails
            $LatestVersion = $request.ComponentDetails.Version
            if ($legacyPME -eq $false) {
                $PMEReleaseDate = $request.ComponentDetails.ReleaseDate
                if ($? -eq $true) {
                    Write-Host "Success reading from Community XML!" -ForegroundColor Green
                }
                Write-Host "Setting PME Component Version Expectation to individual PME Component Versions:" -ForegroundColor Cyan
                $LatestPMEAgentVersion = $request.ComponentDetails.PatchManagementServiceControllerVersion
                $LatestCacheServiceVersion = $request.ComponentDetails.FileCacheServiceAgentVersion
                $LatestRPCServerVersion = $request.ComponentDetails.RequestHandlerAgentVersion
            }

        } Catch [System.Net.WebException] {
            $overallstatus = '2'
            $diagnosticserrorint = '2'
            $message = "$(Get-Date) ERROR: Error fetching PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message)"
            write-host $message
            $diagnosticsinfo = $diagnosticsinfo + "`n$message"
        } Catch [System.Management.Automation.MetadataException] {
            $overallstatus = '2'
            $diagnosticserrorint = '2'
            $message = "$(Get-Date) ERROR: Error casting to XML, could not parse PMESetup_details.xml, aborting. Error: $($_.Exception.Message)"
            write-host "$message"
            $diagnosticsinfo = $diagnosticsinfo + "`n$message"
        } Catch {
            $overallstatus = '2'
            $diagnosticserrorint = '2'
            $message = "$(Get-Date) ERROR: Error occurred attempting to obtain PMESetup_details.xml, aborting. Error: $($_.Exception.Message)"
            $diagnosticsinfo = $diagnosticsinfo + "`n$message"
        }

        if ($legacyPME -eq $true) {
            Write-Host "Setting PME Component Version Expectation to match overall PME Version" -ForegroundColor Cyan
            $LatestPMEAgentVersion = $request.ComponentDetails.Version
            $LatestCacheServiceVersion = $request.ComponentDetails.Version
            $LatestRPCServerVersion = $request.ComponentDetails.Version
        Try {
            $webRequest = $null; $webResponse = $null
            $webRequest = [System.Net.WebRequest]::Create($PMESetup_detailsURI)
            $webRequest.Method = "HEAD"
            $WebRequest.AllowAutoRedirect = $true
            $WebRequest.KeepAlive = $false
            $WebRequest.Timeout = 10000
            $webResponse = $webRequest.GetResponse()
            $remoteLastModified = ($webResponse.LastModified) -as [DateTime]
            $PMEReleaseDate = $remoteLastModified | Get-Date -Format "yyyy.MM.dd"
            $webResponse.Close()
        } Catch [System.Net.WebException] {
            $overallstatus = '2'
            $diagnosticserrorint = '2'
            write-host "Error fetching header for PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message)"
        } Catch {
            $overallstatus = '2'
            $diagnosticserrorint = '2'
            write-host "Error fetching header for PMESetup_Details.xml, aborting. Error: $($_.Exception.Message)"
        }
    }
  
    Write-Host "Latest PME Version: " -nonewline; Write-Host "$latestversion" -ForegroundColor Green
    Write-Host "Latest PME Release Date: " -nonewline; Write-Host "$PMEReleaseDate" -ForegroundColor Green
    Write-Host "Latest PME Agent Version: " -nonewline; Write-Host "$latestPMEAgentversion" -ForegroundColor Green
    Write-Host "Latest Cache Service Version: " -nonewline; Write-Host "$LatestCacheServiceVersion" -ForegroundColor Green
    Write-Host "Latest RPC Server Version: " -nonewline; Write-Host "$latestrpcserverversion" -ForegroundColor Green
    Write-Host ""
}

Function Get-LatestPMEVersion {
    
    if ($legacyPME -eq $true) {
        if (!($pmeprofile -eq 'insiders')) {
            . Get-PMESetupDetails
            . Confirm-PMEInstalled
            . Confirm-PMEUpdatePending 
        }
        else {
                Write-Host "PME Insiders Stream Detected" -ForegroundColor Yellow
                #$PMEWrapper = get-content "${Env:ProgramFiles(x86)}\N-able Technologies\Windows Agent\log\PMEWrapper.log"
                #$Latest = "Pme.GetLatestVersion result = LatestVersion"
                #$LatestVersion = $LatestMatch.Split(' ')[9].TrimEnd(',')
                $PMECore = get-content "$PMEProgramDataFolder\log\Core.log"
                $Latest = "Latest PMESetup Version is"
                $LatestMatch = ($PMECore -match $latest)[-1]
                #$LatestVersion = $LatestMatch.Split(' ')[10].Trim()
                $LatestVersion = ($LatestMatch -Split(' '))[10]
            }
    }

    if ($legacyPME -eq $false) {
        . Get-PMESetupDetails
        . Confirm-PMEInstalled
        . Confirm-PMEUpdatePending 
    }

}

Function Restore-Date {
    If ($InstallDate.Length -le 7) {
        $MMdd = $InstallDate.Substring(4, 3)
        $Year = $InstallDate.Substring(0, 4)
        $InstallDate = $($Year + "0" + $MMdd)
    }
}
    
Function Confirm-PMEInstalled {

# Check if PME Agent is currently installed
    # Write-Host "Checking if PME Agent is already installed..." -ForegroundColor Cyan
    $PATHS = @($UninstallRegLocation)
    $SOFTWARE = $PMEAgentAppName
    ForEach ($path in $PATHS) {
        $installed = Get-ChildItem -Path $path |
        ForEach-Object { Get-ItemProperty $_.PSPath } |
        Where-Object { $_.DisplayName -match $SOFTWARE } |
        Select-Object -Property DisplayName, DisplayVersion, InstallDate

        If ($null -ne $installed) {
            ForEach ($app in $installed) {
                If ($($app.DisplayName) -eq $PMEAgentAppName) {
                    $PMEAgentAppDisplayVersion = $($app.DisplayVersion)
                    $InstallDate = $($app.InstallDate)
                    If ($null -ne $InstallDate -and $InstallDate -ne "") {
                        . Restore-Date
                        $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                        $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                    }
                    $IsPMEAgentInstalled = "Yes"
                    Write-Host "PME Agent Already Installed: Yes" -ForegroundColor Green
                    # Write-Output "Installed PME Agent Version: $PMEAgentAppDisplayVersion"
                    # Write-Output "Installed PME Agent Date: $InstallDateFormatted"
                }
            }
        } Else {
            $IsPMEAgentInstalled = "No"
            Write-Host "PME Agent Already Installed: No" -ForegroundColor Yellow
        }
    }


# Check if PME RPC Service is currently installed

    # Write-Host "Checking if PME RPC Server Service is already installed..." -ForegroundColor Cyan
    $PATHS = @($UninstallRegLocation)
    $SOFTWARE = $PMERPCAppName
    ForEach ($path in $PATHS) {
        $installed = Get-ChildItem -Path $path |
        ForEach-Object { Get-ItemProperty $_.PSPath } |
        Where-Object { $_.DisplayName -match $SOFTWARE } |
        Select-Object -Property DisplayName, DisplayVersion, InstallDate

        If ($null -ne $installed) {
            ForEach ($app in $installed) {
                If ($($app.DisplayName) -eq $PMERPCAppName) {
                    $PMERPCServerAppDisplayVersion = $($app.DisplayVersion) 
                    $InstallDate = $($app.InstallDate)
                    If ($null -ne $InstallDate -and $InstallDate -ne "") {
                        . Restore-Date
                        $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                        $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                    }
                    $IsPMERPCServerServiceInstalled = "Yes"
                    Write-Host "PME RPC Server Service Already Installed: Yes" -ForegroundColor Green
                    # Write-Output "Installed PME RPC Server Service Version: $PMERPCServerAppDisplayVersion"
                    # Write-Output "Installed PME RPC Server Service Date: $InstallDateFormatted"
                }
            }
        } Else {
            $IsPMERPCServerServiceInstalled = "No"
            Write-Host "PME RPC Server Service Already Installed: No" -ForegroundColor Yellow
        }
    }
    

# Check if PME Cache Service is currently installed
    # Write-Host "Checking if PME RPC Server Service is already installed..." -ForegroundColor Cyan
    $PATHS = @($UninstallRegLocation)
    $SOFTWARE = $PMECacheAppName
    ForEach ($path in $PATHS) {
        $installed = Get-ChildItem -Path $path |
        ForEach-Object { Get-ItemProperty $_.PSPath } |
        Where-Object { $_.DisplayName -match $SOFTWARE } |
        Select-Object -Property DisplayName, DisplayVersion, InstallDate

        If ($null -ne $installed) {
            ForEach ($app in $installed) {
                If ($($app.DisplayName) -eq $PMECacheAppName) {
                    $PMECacheServiceAppDisplayVersion = $($app.DisplayVersion) 
                    $InstallDate = $($app.InstallDate)
                    If ($null -ne $InstallDate -and $InstallDate -ne "") {
                        . Restore-Date
                        $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                        $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                    }
                    $IsPMECacheServiceInstalled = "Yes"
                    Write-Host "PME Cache Service Already Installed: Yes" -ForegroundColor Green
                    # Write-Output "Installed PME Cache Service Version: $PMECacheServiceAppDisplayVersion"
                    # Write-Output "Installed PME Cache Service Date: $InstallDateFormatted"
                }
            }
        } Else {
            $IsPMECacheServiceInstalled = "No"
            Write-Host "PME Cache Service Already Installed: No" -ForegroundColor Yellow
        }
    }

}
         
Function Confirm-PMEUpdatePending {
    # Check if PME is awaiting update for new release but has not updated yet (normally within 48 hours)
    write-host ""
    If ($IsPMEAgentInstalled -eq "Yes") {
        $Date = Get-Date -Format 'yyyy.MM.dd'
        if ($PMEReleaseDate -ne $null) {
            $ConvertPMEReleaseDate = Get-Date "$PMEReleaseDate"
        } 
        if (($legacyPME -eq $true) -and ($PMEReleaseDate -eq $null)){
            $Message = "$(Get-Date) INFO: Script was unable to read PME Release Date from Webpage. Falling back to hardset release date in Script"
            Write-Host $Mssage -ForegroundColor Red
            $StatusMessage = $StatusMessage + "`n$message"
            $diagnosticsinfo = $diagnosticsinfo + "`n$message"
            $ConvertPMEReleaseDate = [datetime]$legacyPMEReleaseDate
        }
        $SelfHealingDate = $ConvertPMEReleaseDate.AddDays($PendingUpdateDays).ToString('yyyy.MM.dd')
        Write-Host "INFO: Script considers a PME update to be pending for ($PendingUpdateDays) days after a new version of PME has been released" -ForegroundColor Yellow -BackgroundColor Black
        $DaysElapsed = (New-TimeSpan -Start $SelfHealingDate -End $Date).Days
        $DaysElapsedReversed = (New-TimeSpan -Start $ConvertPMEReleaseDate -End $Date).Days

        # Only run if current $Date is greater than or equal to $SelfHealingDate and $LatestVersion is greater than $app.DisplayVersion
        If (($Date -ge $SelfHealingDate) -and ([version]$LatestVersion -ge [version]$PMEAgentAppDisplayVersion)) {
            $UpdatePending = "No"
            Write-Host "Update Pending: " -nonewline; Write-Host "No (Last Update was released [$DaysElapsed] days since the grace period)" -ForegroundColor Green    
        } Else {
            $UpdatePending = "Yes"
            Write-Host "Update Pending: " -nonewline; Write-Host "Yes (New Update has been released and [$DaysElapsedReversed] days has elapsed since the grace period)" -ForegroundColor Yellow
        }
    }
}

Function Get-PMEServiceVersions {

    $PMEAgentVersion = (get-item $SolarWindsMSPPMEAgentLocation -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
    $PMECacheVersion = (get-item $SolarWindsMSPCacheLocation -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
    $PMERpcServerVersion = (get-item $SolarWindsMSPRpcServerLocation -ErrorAction SilentlyContinue).VersionInfo.ProductVersion

    if ($PMEAgentStatus -eq $null) { 
        Write-Host "PME Agent service is missing" -ForegroundColor Red
        $PMEAgentStatus = 'Service is Missing'
        $PMEAgentVersion = '0.0'
    }

    if ($PMECacheStatus -eq $null) {
        Write-Host "PME Cache service is missing" -ForegroundColor Red
        $PMECacheStatus = 'Service is Missing'
        $PMECacheVersion = '0.0'
    }

    if ($PMERpcServerStatus -eq $null) {
        Write-Host "PME RPC Server service is missing" -ForegroundColor Red
        $PMERpcServerStatus = 'Service is Missing'
        $PMERpcServerVersion = '0.0'
    }

}

Function Get-PMEProfile {
    $PMEconfigXML = "$PMEProgramDataFolder\config\PmeConfig.xml"
if ($PMEAgentStatus -ne $null) {
    if (test-path $pmeconfigxml) {
    $xml = [xml](Get-Content "$PMEConfigXML")
    $pmeprofile = $xml.Configuration.Profile 
    $pmeofflinescan = $xml.Configuration.OfflineScan   
    }
    else {
        $pmeprofile = 'N/A'
        $pmeofflinescan = 'N/A'
    }
}
else {
    $pmeprofile = 'Error - Agent is running but config file could not be found'
}
Write-Host "PME Profile: " -nonewline; Write-Host "$pmeprofile" -foregroundcolor Green
Write-Host "PME Offline Scanning: " -nonewline; Write-Host "$pmeofflinescan" -foregroundcolor Green
    if ($pmeofflinescan -eq '1') {
        $pmeofflinescanbool = $True
        Write-Host "INFO: PME Offline Scanning is enabled" -ForegroundColor Yellow -BackgroundColor Black
    }
    elseif ($pmeofflinescan -eq "N/A")  {
    }
    else {
        $pmeofflinescanbool = $False
        Write-Host "INFO: PME Offline Scanning is not enabled" -ForegroundColor Yellow -BackgroundColor Black
    }
    write-host ""
}

Function Test-PMEConnectivity {
    $DiagnosticsError = $null
    $diagnosticsinfo = $null
    # Performs connectivity tests to destinations required for PME
    $OSVersion = (Get-WmiObject Win32_OperatingSystem).Caption
    If (($PSVersionTable.PSVersion -ge "4.0") -and (!($OSVersion -match 'Windows 7')) -and (!($OSVersion -match '2008 R2')) -and (!($OSVersion -match 'Small Business Server 2011 Standard'))) {
        Write-Host "Performing HTTPS connectivity tests for PME required destinations..." -ForegroundColor Cyan
        $List1= @("sis.n-able.com")
        $HTTPSError = @()
        $List1 | ForEach-Object {
            $Test1 = Test-NetConnection $_ -port 443
            If ($Test1.tcptestsucceeded -eq $True) {
                $Message = "OK: Connectivity to https://$_ ($(($Test1).RemoteAddress.IpAddressToString)) established"
                Write-Host "$Message" -ForegroundColor Green
                $HTTPSError += "No" 
                $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message
            }
            Else {
                $Message = "ERROR: Unable to establish connectivity to https://$_ ($(($Test1).RemoteAddress.IpAddressToString))"
                Write-Host "$Message" -ForegroundColor Red
                $StatusMessage = $StatusMessage + "`n$Message"
                $HTTPSError += "Yes"
                $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message
            }
        }

        Write-Host "Performing HTTP connectivity tests for PME required destinations..." -ForegroundColor Cyan
        $HTTPError = @()
        $List2= @("sis.n-able.com","download.windowsupdate.com","fg.ds.b1.download.windowsupdate.com")
        $List2 | ForEach-Object {
            $Test1 = Test-NetConnection $_ -port 80
            If ($Test1.tcptestsucceeded -eq $True) {
                $Message = "OK: Connectivity to http://$_ ($(($Test1).RemoteAddress.IpAddressToString)) established"
                Write-Host "$Message" -ForegroundColor Green
                $HTTPError += "No"
                $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message 
            }
            Else {
                $message = "ERROR: Unable to establish connectivity to http://$_ ($(($Test1).RemoteAddress.IpAddressToString))"
                Write-Host $message -ForegroundColor Red
                $StatusMessage = $StatusMessage + "`n$Message"
                $HTTPError += "Yes"
                $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message 
            }
        }

        If (($HTTPError[0] -like "*Yes*") -and ($HTTPSError[0] -like "*Yes*")) {
            $Message = "ERROR: No connectivity to $($List2[0]) can be established"
            Write-EventLog -LogName Application -Source "Get-PMEServices" -EntryType Information -EventID 100 -Message "$Message, aborting.`nScript: Get-PMEServices.ps1"  
            $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message
            write-host "ERROR: No connectivity to $($List2[0]) can be established, aborting"
        }
        ElseIf (($HTTPError[0] -like "*Yes*") -or ($HTTPSError[0] -like "*Yes*")) {
            $Message = "WARNING: Partial connectivity to $($List2[0]) established, falling back to HTTP."
            Write-EventLog -LogName Application -Source "Get-PMEServices" -EntryType Information -EventID 100 -Message "$Message`nScript: Get-PMEServices.ps1"  
            Write-Host "$Message" -ForegroundColor Yellow
            $Fallback = "Yes"
            $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message
        }

        If ($HTTPError[1] -like "*Yes*") {
            $Message = "WARNING: No connectivity to $($List2[1]) can be established"
            Write-EventLog -LogName Application -Source "Get-PMEServices" -EntryType Information -EventID 100 -Message "$Message, you will be unable to download Microsoft Updates!`nScript: Get-PMEServices.ps1"  
            Write-Host "$Message, you will be unable to download Microsoft Updates!" -ForegroundColor Red
            $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message
        }

        If ($HTTPError[2] -like "*Yes*") {
            $Message = "WARNING: No connectivity to $($List2[2]) can be established"
            Write-EventLog -LogName Application -Source "Get-PMEServices" -EntryType Information -EventID 100 -Message "$Message, you will be unable to download Windows Feature Updates!`nScript: Get-PMEServices.ps1"  
            Write-Host "$Message, you will be unable to download Windows Feature Updates!" -ForegroundColor Red  
            $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message
    }
}
    Else {
        $Message = "Windows: $OSVersion`nPowershell: $($PSVersionTable.PSVersion)`nSkipping connectivity tests for PME required destinations as OS is Windows 7/ Server 2008 (R2)/ SBS 2011 and/or Powershell 4.0 or above is not installed"
        Write-Host $Message
        $Fallback = "Yes" 
        $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message  
    }
    $DiagnosticsError = $HTTPSError + $HTTPError
    if ($diagnosticsError -contains 'Yes' ){
        $diagnosticserrorInt = '1'
    }
    else {
        $diagnosticserrorInt = '0'
    }
}

Function Write-Status {
Write-Host ""
Write-Host "SolarWinds MSP PME Agent Status: $PMEAgentStatus ($PMEAgentVersion)"
Write-Host "SolarWinds MSP Cache Service Status: $PMECacheStatus ($PMECacheVersion)"
Write-Host "SolarWinds MSP RPC Server Status: $PMERpcServerStatus ($PMERpcServerVersion)"
Write-Host ""
}

Function Start-Services {
    if (($PMEAgentStatus -eq 'Running') -and ($PMECacheStatus -eq 'Running') -and ($PMERpcServerStatus -eq 'Running')) {
            Write-Host "$(Get-Date) - OK - All PME Services are in a Running State" -foregroundcolor Green
    }
    else {
        $RecheckStatus = $True
        if ($PMEAgentStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP PME Agent" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP PME Agent...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds.MSP.PME.Agent.PmeService" 
        }

        if ($PMERpcServerStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP RPC Server" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP RPC Server Service...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds MSP RPC Server" 
        }

        if ($PMECacheStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP Cache Service Service" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP Cache Service...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds MSP Cache Service Service" 
        }
    }
}   

Function Set-AutomaticStartup {
    if (($SolarWinds.MSP.PME.Agent.PmeServiceStartup -eq 'Auto') -and ($SolarWinds.MSP.CacheServiceStartup -eq 'Auto') -and ($SolarWinds.MSP.RpcServerServiceStartup -eq 'Auto')) {
            Write-Host "$(Get-Date) - OK - All PME Services are set to Automatic Startup" -foregroundcolor Green
    }
    else {
        $RecheckStatus = $True
        if ($SolarWinds.MSP.PME.Agent.PmeServiceStartup -ne 'Auto') {
            Write-Host "Changing SolarWinds MSP PME Agent to Automatic" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 102 -Message "Setting SolarWinds MSP PME Agent to Automatic...`nSource: Get-PMEServices.ps1"
            set-service -Name "SolarWinds.MSP.PME.Agent.PmeService" -StartupType Automatic
        }

        if ($SolarWinds.MSP.RpcServerServiceStartup -ne 'Auto') {
            Write-Host "Changing SolarWinds MSP RPC Server to Automatic" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 102 -Message "Setting SolarWinds MSP RPC Server Service to Automatic...`nSource: Get-PMEServices.ps1"
            set-service -Name "SolarWinds MSP RPC Server" -StartupType Automatic
        }

        if ($SolarWinds.MSP.CacheServiceStartup -ne 'Auto') {
            Write-Host "Changing SolarWinds MSP Cache Service Service to Automatic" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 102 -Message "Setting SolarWinds MSP Cache Service to Automatic...`nSource: Get-PMEServices.ps1"
            set-service -Name "SolarWinds MSP Cache Service Service" -StartupType Automatic 
        }
    }
}   

Function Validate-PME {
Write-Host ""
If ([version]$PMEAgentVersion -ge [version]$latestpmeagentversion) {
    Write-Host "PME Agent Version: " -nonewline; Write-Host "Up To Date ($PMEAgentVersion)" -ForegroundColor Green
}
else {
    Write-Host "PME Agent Version: " -nonewline; Write-Host "Not Up To Date ($PMEAgentVersion)" -ForegroundColor Red
}

If ([version]$PMECacheVersion -ge [version]$LatestCacheServiceVersion) {
    Write-Host "PME Cache Service Version: " -nonewline; Write-Host "Up To Date ($PMECacheVersion)" -ForegroundColor Green
}
else {
    Write-Host "PME Cache Service Version: " -nonewline; Write-Host "Not Up To Date ($PMECacheVersion)" -ForegroundColor Red
}

If ([version]$PMERpcServerVersion -ge [version]$latestrpcserverversion) {
    Write-Host "PME RPC Server Version: " -nonewline; Write-Host "Up To Date ($PMERpcServerVersion)" -ForegroundColor Green
}
else {
    Write-Host "PME RPC Server Version: " -nonewline; Write-Host "Not Up To Date ($PMERpcServerVersion)" -ForegroundColor Red
}


if (($PMECacheVersion -eq '0.0') -or ($PMEAgentVersion -eq '0.0') -or ($PMERpcServerVersion -eq '0.0')) {
    $OverallStatus = 1
    $StatusMessage = "$(Get-Date) - WARNING: PME is missing one or more application installs"
    Write-Host ""
    Write-Host "$StatusMessage" -ForegroundColor Red
}

elseif (([version]$PMECacheVersion -ge [version]$latestcacheserviceversion) -and ([version]$PMEAgentVersion -ge [version]$latestpmeagentversion) -and ([version]$PMERpcServerVersion -ge [version]$latestrpcserverversion)) {
    $OverallStatus = 0  
    $StatusMessage = "$(Get-Date) - OK: All PME Services are running the latest version`n" + $StatusMessage  
    Write-Host ""
    Write-Host "$StatusMessage" -foregroundcolor Green
}
elseif ($UpdatePending -eq "Yes") {
    $OverallStatus = 0  
    $StatusMessage = "$(Get-Date) - OK: All PME Services are awaiting an update to the latest version`n" + $StatusMessage
    Write-Host ""   
    Write-Host "$StatusMessage" -foregroundcolor Green    
}
else {
    $OverallStatus = 2
    $StatusMessage = "$(Get-Date) - WARNING: One or more PME Services are not running the latest version`n" + $StatusMessage
    Write-Host "" 
    Write-Host "$StatusMessage" -foregroundcolor Yellow
    Write-Host ""
}
if ($OverallStatus -eq "0") {
    Write-Host "PME Status: " -nonewline; Write-Host "$OverallStatus" -ForegroundColor Green
}
else {
    Write-Host "PME Status: " -nonewline; Write-Host "$OverallStatus" -ForegroundColor Red
}
Write-Host ""
}

Function Get-PMEAnalysis {
if (test-path "$NCentralLog\PME_Install_*.log") {
    $pmeinstalllog = ((get-childitem "$NCentralLog\PME_Install_*.log" | where-object {$_.name -like "*[0-9].log"})[-1]).VersionInfo.FileName
    [datetime]$pmeinstalllogdate = (Get-Content -Path $pmeinstalllog | Select-Object -First 1).substring(0,10)

    #$pmeinstalllogcontent = get-content $pmeinstalllog

    if ($pmeinstalllogdate -ne $null) {
        $dateexecute = get-date
        $installtimedifference = ($dateexecute - $pmeinstalllogdate).Days
        $PMEInstallTimeData = "The last PME install was carried out during a detection $installtimedifference Days ago."
    }
    else {
        $PMEInstallTimeData = "Error: The last PME install date was not found in the PME install log."
    }

    $PMEQueryLogContent = get-content "$PMEProgramDataFolder\log\QueryManager.log"
    $PMEScanDateFound = $PMEQueryLogContent -contains '===============================>>>>> Start scan <<<<<========================================'
    if (($PMEQueryLogContent -eq $null) -or ($PMESCanDateFound -eq $false)) {
        Write-Host "No PME Scan data was found" -ForegroundColor Red
        $lastdetectionlogdate = $null
        $PMELastScanData = "There has been no recent patch detection scan." 
    }
    else {
        Write-Host "PME Scan data was found" -ForegroundColor Green
    [datetime]$lastdetectionlogdate = (($PMEQueryLogContent -contains '===============================>>>>> Start scan <<<<<========================================')[-1]).Split(" ")[1]
    $detectiontimedifference = ($dateexecute - $lastdetectionlogdate).Days
    $PMELastScanData = "The last patch detection scan took place $detectiontimedifference Days ago"
    }

    if (($installtimedifference -gt $detectiontimedifference) -and ($PMEAgentVersion -ne $latestversion)) {
        $InstallProblem = "There was a problem with the automatic upgrade process. Recommend using Repair-PME to force upgrade of PME Agent"
    }
    else {
        $installproblem = $null
    }

    $PMEAnalysisMessage = "Installed PME Version: $PMEAgentVersion`nLatest PME Version: $latestversion`n$PMEInstallTimeData`n$PMELastScanData`n$InstallProblem"
    Write-Host "$PMEAnalysisMessage"

    $PMEInstallerExefromLog = [Regex]::Matches(($pmeinstalllogcontent -match "Original Setup EXE"),'[A-Z]:\\(?:[^\\\/:*?"<>|\r\n]+\\)*[^\\\/:*?"<>|\r\n]*$').value
    $startinglinecontent = "Installing SolarWinds.MSP.PME.Agent.exe windows service"
    
    $TotalLinesInFile = ($pmeinstalllogcontent | Measure-Object).count
    $startingLineNumber = ($pmeinstalllogcontent | select-string -pattern $startinglinecontent | select-object -expandproperty 'LineNumber') -2
    if ($startinglinenumber -ne '-2'){ 
        $relevantlines = $TotalLinesInFile - $startinglinenumber
        $UpgradeError = get-content $pmeinstalllog -last $relevantlines
        Write-Host ""
        Write-Host "Installer EXE: " -ForegroundColor Green -nonewline; Write-Host "$PMEInstallerExefromLog"
        if ($pmeInstallerExefromLog -ne "$PMEProgramDataFolder\PME\Archives\PMESetup_$LatestVersion.exe") {
            Write-Host "Incorrect Setup EXE is being used" -foregroundcolor Red
        }
        else {
            Write-Host "Correct Setup EXE is being used" -foregroundcolor Green
        }
        Write-Host ""
        Write-Host "Last Upgrade Results: " -ForegroundColor Green
        get-content $pmeinstalllog -last $relevantlines
        }
        else {
            $pmeinstalllog = 'Error: There was no successful upgrades detected'
            Write-Host $pmeinstalllog -foregroundcolor Red        
        }
    }
    else {
        $pmeinstalllog = 'Error: There was no successful upgrades detected'
        Write-Host $pmeinstalllog -foregroundcolor Red    
    }
    
}
    
Function Get-PMEConfigMisconfigurations {
    # Check PME Config and inform of possible misconfigurations
    Write-Host "PME Config Details:" -ForegroundColor Cyan
     Try {

        If (Test-Path "$CacheServiceConfigFile") {
            $xml = New-Object XML
            $xml.Load($CacheServiceConfigFile)
            $CacheServiceConfig = $xml.Configuration

            If ($null -ne $CacheServiceConfig) {
                If ($CacheServiceConfig.CanBypassProxyCacheService -eq "False") {
                    $CacheConfigMessage = "$(Get-Date) - WARNING: Patch profile doesn't allow PME to fallback to external sources, if probe is not reachable PME may not work!"
                    Write-Warning "$CacheConfigMessage"
                } ElseIf ($CacheServiceConfig.CanBypassProxyCacheService -eq "True") {
                    $CacheConfigMessage = "$(Get-Date) - INFO: Patch profile allows PME to fallback to external sources"
                    Write-Host "$CacheConfigMessage" -ForegroundColor Yellow -BackgroundColor Black
                } Else {
                    $CacheConfigMessage = "$(Get-Date) - WARNING: Unable to determine if patch profile allows PME to fallback to external sources"
                    Write-Warning "$CacheConfigMessage"
                }


                If ($CacheServiceConfig.CacheSizeInMB -eq 10240) {
                    $CacheConfigSizeMessage = "$(Get-Date) - INFO: Cache Service is set to default cache size of 10240 MB"
                    Write-Host "$CacheConfigSizeMessage" -ForegroundColor Yellow -BackgroundColor Black
                } Else {
                    $CacheSize = $CacheServiceConfig.CacheSizeInMB
                    $CacheConfigSizeMessage = "$(Get-Date) - WARNING: Cache Service is not set to default cache size of 10240 MB (currently $CacheSize MB), PME may not work at expected!"
                    Write-Warning "$CacheConfigSizeMessage"
                }
            }   
        }
    }    
    Catch {
        $CacheConfigMessage = "$(Get-Date) - WARNING: Unable to read Cache Service config file as a valid xml file, default cache size can't be checked"
        Write-Warning "$CacheConfigMessage"
    }   

$StatusMessage = $StatusMessage + "`n" + $CacheConfigMessage + "`n" + $CacheConfigSizeMessage
Write-Host ""
}

#endregion

. Test-PMERequirement
. Set-PMEExpectations

if ($continue -eq $true) {
    . Get-PMEServicesStatus
    . Get-PMEServiceVersions
    . Get-PMEProfile
    . Get-LatestPMEVersion
    . Validate-PME
    . Get-PMEConfigMisconfigurations

    if ($RecheckStartup -eq $True) {
    . Get-PMEServicesStartup   
    . Write-Startup
    }

    if ($RecheckStatus -eq $True) {
    . Get-PMEServicesStatus   
    . Write-Status
    }

    if (($OverallStatus -ne '0') -or ($Diagnostics -eq 'True')) {
        Write-Host "Error Detected. Running diagnostics..." -ForegroundColor Red
    . Test-PMEConnectivity
    # Write-Host "$DiagnosticsInfo`n"
    # Write-Host "$DiagnosticsError"
    Write-Host "Diagnostics Error: " -nonewline; Write-Host "$DiagnosticsErrorInt" -ForegroundColor Green
    }

    if ($OverallStatus -ne '0') {
        . Get-PMEAnalysis
    }
}

$SolarWindsMSPPMEAgentStatus = $PMEAgentStatus
$SolarWindsMSPCacheStatus = $PMECacheStatus
$SolarWindsMSPRpcServerStatus = $PMERPCServerStatus