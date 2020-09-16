<#    
    ************************************************************************************************************
    Name: Get-PMEServices.ps1
    Version: 0.1.6.7 (04th September 2020)
    Author: Prejay Shah (Doherty Associates)
    Thanks To: Ashley How
    Purpose:    Get/Reset PME Service Details
    Pre-Reqs:    PowerShell 2.0 (PowerShell 4.0+ for Connectivity Tests)
    Version History:    0.1.0.0 - Initial Release.
                                + Improved Detection for PME Services being missing on a device
                                + Improved Detection of Latest PME Version
                                + Improved Detection and error handling for latest Public PME Version when PME 1.1 is installed or PME is not installed on a device
                                + Improved Compatibility of PME Version Check
                                + Updated to include better PME 1.2.x Latest Versioncheck
                                + Updated for better PS 2.0 Compatibility
                        0.1.5.0 + Added PME Profile Detection, and Diagnostics information courtesy of Ashley How's Repair-PME
                        0.1.5.0a + N-Central AMP Variant using AMP Input Parameter to control Forcing of Diagnostics Mode (Disabled by Default)	
                        0.1.5.1 + Improved TLS Support, Updated Error Message for PME connectivity Test not working on Windows 7
                        0.1.5.2 + PME 1.2.4 has been made GA for the default stream so have had to alter detection methods
                        0.1.5.3 + Improved Compatibility with Server 2008 R2
                        0.1.5.4 + Updated 'Test-PMEConnectivity' function to fix a message typo. Thanks for Clint Conner for finding. 
                        0.1.6.0 + Have Added in PME Installer Log Analysis for use when PME is not up to date
                        0.1.6.1 + Have added date analysis of log file/detection/installation proceedings
                        0.1.6.2 + Fixed Detection Logic for when there has been no PME Scan on a device, and missing components
                        0.1.6.3 + Added Reading in of PME Config for Cache settings
                        0.1.6.4 + Added in better x86/x64 compatability because apparently there are still 32-bit OS devices out there.
                        0.1.6.5 + Fixed Typo
                        0.1.6.6 + [Ashley How] migrated code from my Repair-PME script giving it the abilty to no longer consider a pending update of PME a failure.
                                  Please Note: By default the grace period has been set at 2 days but can be changed by changing the $PendingUpdateDays in settings section.
                                + Various updates to functions and parts of the script to be closer in-line with Repair-PME script.
                                + Updated Validate-PME function to account for pending update, this will report a status of 0. Renamed status messages to make it eaiser to theshold in an AMP service.
                                + Fixed minor issues with date-time parsing, valid PS/OS detection, URL Querying 
                        0.1.6.7 + Added Offine Scanning Detection

    Examples: 
    Diagnostics Input: False
    Runs Normally and only engages diagnostcs for connectibity testing if PME is not up to date or missing a service.
    
    Diagnostics Input: True
    Force Diagnostics Mode to be enabled on the run regardless of PME Status
    ************************************************************************************************************
#>

Param (
        [Parameter(Mandatory=$false,Position=1)]
        [switch] $Diagnostics
    )

# Settings
# *******************************************************************************************************************************
# Change this variable to number of days (must be a number!) to consider a new version of PME as pending an update. Default is 2.
$PendingUpdateDays = "2"
# *******************************************************************************************************************************

$Version = '0.1.6.7 (04th September 2020)'
$RecheckStartup = $Null
$RecheckStatus = $Null
$request = $null
$Latestversion = $Null
$pmeprofile = $null
$diagnosticserrorint = $null
$pmeinstalllogcontent = $null
$EventLogCompanyName ="Doherty Associates"

Write-Host "Get-PMEServices $Version"

if ($Diagnostics -eq 'True'){
    Write-Host "Diagnostics Mode Enabled" -foregroundcolor Yellow
}

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
  } catch {
    Write-Output 'Unable to set PowerShell to use TLS 1.2 and TLS 1.1 due to old .NET Framework installed. If you see underlying connection closed or trust errors, you may need to upgrade to .NET Framework 4.5+ and PowerShell v3+.'
  }

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}


#region Functions

Function Get-PMEServicesStatus {
$SolarWindsMSPCacheStatus = (get-service "SolarWinds.MSP.CacheService" -ErrorAction SilentlyContinue).Status
$SolarWindsMSPPMEAgentStatus = (get-service "SolarWinds.MSP.PME.Agent.PmeService" -ErrorAction SilentlyContinue).Status
$SolarWindsMSPRpcServerStatus = (get-service "SolarWinds.MSP.RpcServerService" -ErrorAction SilentlyContinue).status
}

Function Get-PMESetupDetails {
    # Declare static URI of PMESetup_details.xml
    If ($Fallback -eq "Yes") {
        $PMESetup_detailsURI = "http://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml"
    } Else {
        $PMESetup_detailsURI = "https://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml"
    }
    Try {
        $request = $null
        [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$PMESetup_detailsURI") -split '<\?xml.*\?>')[-1]
        $PMEDetails = $request.ComponentDetails
        $LatestVersion = $request.ComponentDetails.Version
    } Catch [System.Net.WebException] {
        $overallstatus = '2'
        $diagnosticserrorint = '2'
        Throw "Error fetching PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message)"
    } Catch [System.Management.Automation.MetadataException] {
        $overallstatus = '2'
        $diagnosticserrorint = '2'
        Throw "Error casting to XML, could not parse PMESetup_details.xml, aborting. Error: $($_.Exception.Message)"
    } Catch {
        $overallstatus = '2'
        $diagnosticserrorint = '2'
        Throw "Error occurred attempting to obtain PMESetup details, aborting. Error: $($_.Exception.Message)"
    }
}


Function Get-PMEConfigurationDetails {
    # Declare static URI of PmeConfiguration_details.xml
    $Fallback
    If ($Fallback -eq "Yes") {
        $PMEConfigurationDetailsURI = "http://sis.n-able.com/ComponentData/RMM/all/PmeConfiguration_details.xml"
    } Else {
        $PMEConfigurationDetailsURI = "https://sis.n-able.com/ComponentData/RMM/all/PmeConfiguration_details.xml"
    }
    Try {
        $request = $null
        [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$PMEConfigurationDetailsURI") -split '<\?xml.*\?>')[-1]
        $PMEConfigurationDetails = $request.ComponentDetails
        $PMEConfigurationDate = $PMEConfigurationDetails.Version
        $PMEConfigurationDate = $PMEConfigurationDate.Substring(0, $PMEConfigurationDate.Length - 3)
        Write-Output "Latest PME Version: $LatestVersion"
        Write-Output "Latest PME Release Date: $PMEConfigurationDate"
    } Catch [System.Net.WebException] {
        Write-Output "Error fetching PMESetup_Details.xml check your source URL!"
        Throw
    } Catch [System.Management.Automation.MetadataException] {
        Write-Output "Error casting to XML; could not parse PMESetup_details.xml"
        Throw
    }
}

Function Get-LatestPMEVersion {
 
    if (!($pmeprofile -eq 'alpha')) {
        . Get-PMESetupDetails
        . Get-PMEConfigurationDetails
        . Confirm-PMEInstalled
        . Confirm-PMEUpdatePending 
    }
    else {
            Write-Host "PME Alpha Stream Detected" -ForegroundColor Yellow
            #$PMEWrapper = get-content "${Env:ProgramFiles(x86)}\N-able Technologies\Windows Agent\log\PMEWrapper.log"
            #$Latest = "Pme.GetLatestVersion result = LatestVersion"
            #$LatestVersion = $LatestMatch.Split(' ')[9].TrimEnd(',')
            $PMECore = get-content "$env:programdata\SolarWinds MSP\PME\log\Core.log"
            $Latest = "Latest PME Version is"
            $LatestMatch = ($PMECore -match $latest)[-1]
            $LatestVersion = $LatestMatch.Split(' ')[10].Trim()
        }
        Write-Host "Latest Version: " -nonewline; Write-Host "$latestversion" -ForegroundColor Green
}

Function Restore-Date {
    If ($InstallDate.Length -le 7) {
        $MMdd = $InstallDate.Substring(4, 3)
        $Year = $InstallDate.Substring(0, 4)
        $InstallDate = $($Year + "0" + $MMdd)
    }
}
    
Function Confirm-PMEInstalled {
    # Check if PME is currently installed
    If ($OSArch -like '*64*') {
        $PATHS = @("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
        $SOFTWARE = "SolarWinds MSP Patch Management Engine"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName, DisplayVersion, InstallDate

            If ($null -ne $installed) {
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "SolarWinds MSP Patch Management Engine") {
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsPMEInstalled = "Yes"
                        Write-Host "PME Already Installed: " -NoNewline; Write-Host "Yes" -ForegroundColor Green
                        Write-Output "Installed PME Version: $($app.DisplayVersion)"
                        Write-Output "Installed PME Date: $InstallDateFormatted"
                    }
                }
            } Else {
                $IsPMEInstalled = "No"
                Write-Host "PME Already Installed: " -NoNewline; Write-Host "No" -ForegroundColor Yellow
            }
        }
    }

    If ($OSArch -like '*32*') {
        $PATHS = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
        $SOFTWARE = "SolarWinds MSP Patch Management Engine"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName, DisplayVersion, InstallDate

            If ($null -ne $installed) {
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "SolarWinds MSP Patch Management Engine") {
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsPMEInstalled = "Yes"
                        Write-Host "PME Already Installed: " -NoNewline; Write-Host "Yes" -ForegroundColor Green
                        Write-Output "Installed PME Version: $($app.DisplayVersion)"
                        Write-Output "Installed PME Date: $InstallDateFormatted"
                    }
                }
            } Else {
                $IsPMEInstalled = "No"
                Write-Host "PME Already Installed: " -NoNewline; Write-Host "No" -ForegroundColor Yellow
            }
        }
    }
}


Function Confirm-PMEUpdatePending {
    # Check if PME is awaiting update for new release but has not updated yet (normally within 48 hours)
    If ($IsPMEInstalled -eq "Yes") {
        $Date = Get-Date -Format 'yyyy.MM.dd'
        Write-Output "Current Date: $Date"
        $ConvertPMEConfigurationDate = Get-Date "$PMEConfigurationDate"
        $SelfHealingDate = $ConvertPMEConfigurationDate.AddDays($PendingUpdateDays).ToString('yyyy.MM.dd')
        Write-Host "Get-PMEServices considers a PME update to be pending for ($PendingUpdateDays) days after a new version of PME has been released" -ForegroundColor Cyan
        $DaysElapsed = (New-TimeSpan -Start $SelfHealingDate -End $Date).Days
        $DaysElapsedReversed = (New-TimeSpan -Start $PMEConfigurationDate -End $Date).Days

        # Only run if current $Date is greater than or equal to $SelfHealingDate and $LatestVersion is greater than $app.DisplayVersion
        If (($Date -ge $SelfHealingDate) -and ($LatestVersion -ge $($app.DisplayVersion))) {
            $UpdatePending = "No"
            Write-Host "Update Pending: " -nonewline; Write-Host "No (Last Update was released [$DaysElapsed] since the grace period)" -ForegroundColor Green    
        } Else {
            $UpdatePending = "Yes"
            Write-Host "Update Pending: " -nonewline; Write-Host "Yes (New Update has been released and [$DaysElapsedReversed] days has elapsed since the grace period)" -ForegroundColor Yellow
        }
    } Else {
        Write-Warning "Skipping update pending check as PME is not currently installed"
    }
}


Function Get-PMEServiceVersions {
    $OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
    If ($OSArch -like '*64*') {
        $SolarWindsMSPCacheLocation = 'C:\Program Files (x86)\SolarWinds MSP\CacheService\SolarWinds.MSP.CacheService.exe'
        $SolarWindsMSPPMEAgentLocation = 'C:\Program Files (x86)\SolarWinds MSP\PME\SolarWinds.MSP.PME.Agent.exe'
        $SolarWindsMSPRpcServerLocation = 'C:\Program Files (x86)\SolarWinds MSP\RpcServer\SolarWinds.MSP.RpcServerService.exe'
        $NCentralLog = "c:\Program Files (x86)\N-able Technologies\Windows Agent\log"
    }
    else {
        $SolarWindsMSPCacheLocation = 'C:\Program Files\SolarWinds MSP\CacheService\SolarWinds.MSP.CacheService.exe'
        $SolarWindsMSPPMEAgentLocation = 'C:\Program Files\SolarWinds MSP\PME\SolarWinds.MSP.PME.Agent.exe'
        $SolarWindsMSPRpcServerLocation = 'C:\Program Files\SolarWinds MSP\RpcServer\SolarWinds.MSP.RpcServerService.exe'
        $NCentralLog = "c:\Program Files\N-able Technologies\Windows Agent\log"
    }

    $PMECacheVersion = (get-item $SolarWindsMSPCacheLocation -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
    $PMEAgentVersion = (get-item $SolarWindsMSPPMEAgentLocation -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
    $PMERpcServerVersion = (get-item $SolarWindsMSPRpcServerLocation -ErrorAction SilentlyContinue).VersionInfo.ProductVersion

    if ($SolarWindsMSPCacheStatus -eq $null) {
        Write-Host "PME Cache service is missing" -ForegroundColor Red
        $SolarWindsMSPCacheStatus = 'Service is Missing'
        $PMECacheVersion = '0.0'
    }

    if ($SolarWindsMSPPMEAgentStatus -eq $null) { 
        Write-Host "PME Agent service is missing" -ForegroundColor Red
        $SolarWindsMSPPMEAgentStatus = 'Service is Missing'
        $PMEAgentVersion = '0.0'
    }

    if ($SolarWindsMSPRpcServerStatus -eq $null) {
        Write-Host "PME RPC Server service is missing" -ForegroundColor Red
        $SolarWindsMSPRpcServerStatus = 'Service is Missing'
        $PMERpcServerVersion = '0.0'
    }

}

Function Get-PMEProfile {
    $PMEconfigXML = "C:\ProgramData\SolarWinds MSP\PME\config\PmeConfig.xml"
if ($SolarWindsMSPPMEAgentStatus -ne $null) {
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
Write-Host "PME Profile: " -nonewline; Write-Host "$pmeprofile`n" -foregroundcolor Green
Write-Host "PME Offline Scanning: " -nonewline; Write-Host "$pmeofflinescan`n" -foregroundcolor Green
    if ($pmeofflinescan -eq '1') {
        $pmeofflinescanbool = $True
        Write-Host "PME Offline Scanning is enabled" -ForegroundColor Yellow
    }
    else {
        $pmeofflinescanbool = $False
        Write-Host "PME Offline Scanning is not enabled" -ForegroundColor Yellow
    }
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
                $HTTPError += "Yes"
                $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message 
            }
        }

        If (($HTTPError[0] -like "*Yes*") -and ($HTTPSError[0] -like "*Yes*")) {
            $Message = "ERROR: No connectivity to $($List2[0]) can be established"
            Write-EventLog -LogName Application -Source "Get-PMEServices" -EntryType Information -EventID 100 -Message "$Message, aborting.`nScript: Get-PMEServices.ps1"  
            $diagnosticsinfo = $diagnosticsinfo + '`n' + $Message
            Throw "ERROR: No connectivity to $($List2[0]) can be established, aborting"
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
        Write-Output $Message
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
Write-Host "`nSolarWinds MSP Cache Service Status: $SolarWindsMSPCacheStatus"
Write-Host "SolarWinds MSP PME Agent Status: $SolarWindsMSPPMEAgentStatus"
Write-Host "SolarWinds MSP RPC Server Status: $SolarWindsMSPRpcServerStatus`n"
}

Function Write-Version {
Write-Host "`nSolarWinds MSP Cache Service Version: $PMECacheVersion"
Write-Host "SolarWinds MSP PME Agent Version: $PMEAgentVersion"
Write-Host "SolarWinds MSP RPC Server Version: $PMERpcServerVersion`n"
}

Function Start-Services {
    if (($SolarWindsMSPPMEAgentStatus -eq 'Running') -and ($SolarWindsMSPCacheStatus -eq 'Running') -and ($SolarWindsMSPRpcServerStatus -eq 'Running')) {
            Write-Host "OK - All PME Services are in a Running State" -foregroundcolor Green
    }
    else {
        $RecheckStatus = $True
        if ($SolarWindsMSPPMEAgentStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP PME Agent" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP PME Agent...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds.MSP.PME.Agent.PmeService" 
        }

        if ($SolarWindsMSPRpcServerStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP RPC Server" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP RPC Server Service...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds MSP RPC Server" 
        }

        if ($SolarWindsMSPCacheStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP Cache Service Service" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source $EventLogCompanyName -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source $EventLogCompanyName -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP Cache Service...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds MSP Cache Service Service" 
        }
    }
}   

Function Set-AutomaticStartup {
    if (($SolarWinds.MSP.PME.Agent.PmeServiceStartup -eq 'Auto') -and ($SolarWinds.MSP.CacheServiceStartup -eq 'Auto') -and ($SolarWinds.MSP.RpcServerServiceStartup -eq 'Auto')) {
            Write-Host "OK - All PME Services are set to Automatic Startup" -foregroundcolor Green
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

If ([version]$PMECacheVersion -ge $latestversion) {
    Write-Host "PME Cache Version: " -nonewline; Write-Host "Up To Date ($PMECacheVersion)" -ForegroundColor Green
}
else {
    Write-Host "PME Cache Version: " -nonewline; Write-Host "Not Up To Date ($PMECacheVersion)" -ForegroundColor Red
}

If ([version]$PMEAgentVersion -ge $latestversion) {
    Write-Host "PME Agent Version: " -nonewline; Write-Host "Up To Date ($PMEAgentVersion)" -ForegroundColor Green
}
else {
    Write-Host "PME Agent Version: " -nonewline; Write-Host "Not Up To Date ($PMEAgentVersion)" -ForegroundColor Red
}

If ([version]$PMERpcServerVersion -ge $latestversion) {
    Write-Host "PME RPC Server Version: " -nonewline; Write-Host "Up To Date ($PMERpcServerVersion)" -ForegroundColor Green
}
else {
    Write-Host "PME RPC Server Version: " -nonewline; Write-Host "Not Up To Date ($PMERpcServerVersion)" -ForegroundColor Red
}


if (($PMECacheVersion -eq '0.0') -or ($PMEAgentVersion -eq '0.0') -or ($PMERpcServerVersion -eq '0.0')) {
    $OverallStatus = 1
    $StatusMessage = 'PME is missing one or more application installs'
    Write-Host "`n$StatusMessage" -ForegroundColor Red
}

elseif (([version]$PMECacheVersion -ge $latestversion) -and ([version]$PMEAgentVersion -ge $latestversion) -and ([version]$PMERpcServerVersion -ge $latestversion)) {
    $OverallStatus = 0  
    $StatusMessage = 'OK - All PME Services are running the latest version'    
    Write-Host "`n$StatusMessage" -foregroundcolor Green
}
elseif ($UpdatePending -eq "Yes") {
    $OverallStatus = 0  
    $StatusMessage = 'OK - All PME Services are awaiting an update to the latest version'    
    Write-Host "`n$StatusMessage" -foregroundcolor Green    
}
else {
    $OverallStatus = 2
    $StatusMessage = 'One or more PME Services are not running the latest version'
    Write-Host "`n$StatusMessage`n" -foregroundcolor Yellow
    
}
Write-Host "Status: $OverallStatus"
}

Function Get-PMEAnalysis {
if (test-path "$NCentralLog\PME_Install_*.log") {
    $pmeinstalllog = ((get-childitem "$NCentralLog\PME_Install_*.log" | where-object {$_.name -like "*[0-9].log"})[-1]).VersionInfo.FileName
    $pmeinstalllogcontent = get-content $pmeinstalllog
    [datetime]$pmeinstalllogdate = ($($pmeinstalllogcontent.SubString(0,10)[0] | out-string))
    $dateexecute = get-date
    $installtimedifference = ($dateexecute - $pmeinstalllogdate).Days
    $PMEInstallTimeData = "The last PME install was carried out during a detection $installtimedifference Days ago."

    $PMEQueryLogContent = get-content "C:\ProgramData\SolarWinds MSP\PME\log\QueryManager.log"
    $PMEScanDateFound = $PMEQueryLogContent -match '===============================>>>>> Start scan <<<<<========================================'
    if (($PMEQueryLogContent -eq $null) -or ($PMESCanDateFound -eq $false)) {
        Write-Host "No PME Scan data was found" -ForegroundColor Red
        $lastdetectionlogdate = $null
        $PMELastScanData = "There has been no recent patch detection scan." 
    }
    else {
        Write-Host "PME Scan data was found" -ForegroundColor Green
    [datetime]$lastdetectionlogdate = (($PMEQueryLogContent -match '===============================>>>>> Start scan <<<<<========================================')[-1]).Split(" ")[1]
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
        Write-Host "`nInstaller EXE: " -ForegroundColor Green -nonewline; Write-Host "$PMEInstallerExefromLog"
        if ($pmeInstallerExefromLog -ne "C:\ProgramData\SolarWinds MSP\PME\Archives\PMESetup_$LatestVersion.exe") {
            Write-Host "Incorrect Setup EXE is being used" -foregroundcolor Red
        }
        else {
            Write-Host "Correct Setup EXE is being used" -foregroundcolor Green
        }
        Write-Host "`nLast Upgrade Results: " -ForegroundColor Green
        get-content $pmeinstalllog -last $relevantlines
        }
        else {
            $pmeinstalllog = 'There was no successful upgrades detected'
            Write-Host $pmeinstalllog -foregroundcolor Red        
        }
    }
    else {
        $pmeinstalllog = 'There was no successful upgrades detected'
        Write-Host $pmeinstalllog -foregroundcolor Red    
    }

}

Function Get-PMEConfigMisconfigurations {
    # Check PME Config and inform of possible misconfigurations
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
#endregion

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
    Write-Host "Error Detected so running diagnostics" -ForegroundColor Red
. Test-PMEConnectivity
# Write-Host "$DiagnosticsInfo`n"
# Write-Host "$DiagnosticsError"
Write-Host "Diagnostics Error: " -nonewline; Write-Host "$DiagnosticsErrorInt" -ForegroundColor Green
}

if ($OverallStatus -ne '0') {
    . Get-PMEAnalysis
}


