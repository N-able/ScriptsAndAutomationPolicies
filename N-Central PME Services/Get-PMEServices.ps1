<#    
    ************************************************************************************************************
    Name: Get-PMEServices.ps1
    Version: 0.1.5.4 (22nd May 2020)
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
                        0.1.5.0 + Added PME Profile Detection, and Diagnostics information courtesy of Ashley How's Repair-PME via a disagnostics parameter
                        0.1.5.1 + Improved TLS Support, Updated Error Message for PME connectivity Test not working on Windows 7
                        0.1.5.2 + PME 1.2.4 has been made GA for the default stream so have had to alter detection methods
                        0.1.5.3 + Improved Compatibility with Server 2008 R2
                        0.1.5.4 + Updated 'Test-PMEConnectivity' function to fix a message typo. Thanks for Clint Conner for finding. 

    Examples: 
    .\get-pmeservices.ps1
    Runs Normally and only engages diagnostcs for connectibity testing if PME is not up to date or missing a service.
    
    .\get-pmeservices.ps1 -diagnostics
    Force Diagnostics Mode to be enabled on the run regardless of PME Status

    ************************************************************************************************************
#>
Param (
        [Parameter(Mandatory=$false,Position=1)]
        [switch] $Diagnostics
    )

$Version = '0.1.5.4 (22nd May 2020)'
$RecheckStartup = $Null
$RecheckStatus = $Null
$request = $null
$Latestversion = $Null
$pmeprofile = $null
$diagnosticserrorint = $null

Write-Host "Get-PMEServices $Version"

if ($Diagnostics){
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

Function Get-LatestPMEVersionfromURL {
# Declare static URI of PMESetup_details.xml
$PMESetup_detailsURI = "https://sis.n-able.com/Components/MSP-PME/latest/PMESetup_details.xml"

    Try {
        [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$PMESetup_detailsURI") -split '<\?xml.*\?>')[-1]
        $LatestVersion = $request.ComponentDetails.Version
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


Function Get-LatestPMEVersion {
 
    if (!($pmeprofile -eq 'alpha')) {
    . Get-LatestPMEVersionfromURL
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

Function Get-PMEServiceVersions {
    <#
    if ([version]$psversiontable.psversion -le '2.0') {
        Write-Host "PS 2.0 Compatible Version" -ForegroundColor Yellow
        $SolarWindsMSPCacheLocation = (get-wmiobject win32_service -filter "Name like 'SolarWinds.MSP.CacheService'" -ErrorAction SilentlyContinue).PathName.Replace('"','')
        $SolarWindsMSPPMEAgentLocation = (get-wmiobject win32_service -filter "Name like 'SolarWinds.MSP.PME.Agent.PmeService'" -ErrorAction SilentlyContinue).Pathname.Replace('"','')
        $SolarWindsMSPRpcServerLocation = (get-wmiobject win32_service -filter "Name like 'SolarWinds.MSP.RpcServerService'" -ErrorAction SilentlyContinue).Pathname.Replace('"','')
        
        }
        else {
            $SolarWindsMSPCacheLocation = (get-ciminstance win32_service -filter "Name like 'SolarWinds.MSP.CacheService'" -OperationTimeoutSec 5 -ErrorAction SilentlyContinue).PathName.Replace('"','')
            $SolarWindsMSPPMEAgentLocation = (get-ciminstance win32_service -filter "Name like 'SolarWinds.MSP.PME.Agent.PmeService'" -OperationTimeoutSec 5 -ErrorAction SilentlyContinue).Pathname.Replace('"','')
            $SolarWindsMSPRpcServerLocation = (get-ciminstance win32_service -filter "Name like 'SolarWinds.MSP.RpcServerService'" -OperationTimeoutSec 5 -ErrorAction SilentlyContinue).Pathname.Replace('"','')
        }
    #>
        $OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
        If ($OSArch -like '*64*') {
            $SolarWindsMSPCacheLocation = 'C:\Program Files (x86)\SolarWinds MSP\CacheService\SolarWinds.MSP.CacheService.exe'
            $SolarWindsMSPPMEAgentLocation = 'C:\Program Files (x86)\SolarWinds MSP\PME\SolarWinds.MSP.PME.Agent.exe'
            $SolarWindsMSPRpcServerLocation = 'C:\Program Files (x86)\SolarWinds MSP\RpcServer\SolarWinds.MSP.RpcServerService.exe'
        }
        else {
            $SolarWindsMSPCacheLocation = 'C:\Program Files\SolarWinds MSP\CacheService\SolarWinds.MSP.CacheService.exe'
            $SolarWindsMSPPMEAgentLocation = 'C:\Program Files\SolarWinds MSP\PME\SolarWinds.MSP.PME.Agent.exe'
            $SolarWindsMSPRpcServerLocation = 'C:\Program Files\SolarWinds MSP\RpcServer\SolarWinds.MSP.RpcServerService.exe'
        }

        $PMECacheVersion = (get-item $SolarWindsMSPCacheLocation).VersionInfo.ProductVersion
        $PMEAgentVersion = (get-item $SolarWindsMSPPMEAgentLocation).VersionInfo.ProductVersion
        $PMERpcServerVersion = (get-item $SolarWindsMSPRpcServerLocation).VersionInfo.ProductVersion

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
    }
    else {
        $pmeprofile = 'N/A'
    }
}
else {
    $pmeprofile = 'Error - Agent is running but config File could not be found'
}
Write-Host "PME Profile: " -nonewline; Write-Host "$pmeprofile`n" -foregroundcolor Green
}


Function Test-PMEConnectivity {
    $DiagnosticsError = $null
    $diagnosticsinfo = $null
    # Performs connectivity tests to destinations required for PME
    $OSVersion = (Get-WmiObject Win32_OperatingSystem).Caption
    If (($PSVersionTable.PSVersion -ge "4.0") -and (!($OSVersion -match 'Windows 7')) -and (!($OSVersion -match '2008 R2'))) {
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
        $Message = "Windows: $OSVersion`n>Powershell: $($PSVersionTable.PSVersion)`nSkipping connectivity tests for PME required destinations as OS is Windows 7/ Server 2008 R2 and/or Powershell 4.0 or above is not installed"
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
            Write-Host "All PME Services are in a Running State" -foregroundcolor Green
    }
    else {
        $RecheckStatus = $True
        if ($SolarWindsMSPPMEAgentStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP PME Agent" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source "Doherty Associates" -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source "Doherty Associates" -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP PME Agent...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds.MSP.PME.Agent.PmeService" 
        }

        if ($SolarWindsMSPRpcServerStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP RPC Server" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source "Doherty Associates" -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source "Doherty Associates" -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP RPC Server Service...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds MSP RPC Server" 
        }

        if ($SolarWindsMSPCacheStatus -ne 'Running') {
            Write-Host "Starting SolarWinds MSP Cache Service Service" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source "Doherty Associates" -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source "Doherty Associates" -EntryType Information -EventID 100 -Message "Starting SolarWinds MSP Cache Service...`nSource: Get-PMEServices.ps1"
            start-service -Name "SolarWinds MSP Cache Service Service" 
        }
    }
}   

Function Set-AutomaticStartup {
    if (($SolarWinds.MSP.PME.Agent.PmeServiceStartup -eq 'Auto') -and ($SolarWinds.MSP.CacheServiceStartup -eq 'Auto') -and ($SolarWinds.MSP.RpcServerServiceStartup -eq 'Auto')) {
            Write-Host "All PME Services are set to Automatic Startup" -foregroundcolor Green
    }
    else {
        $RecheckStatus = $True
        if ($SolarWinds.MSP.PME.Agent.PmeServiceStartup -ne 'Auto') {
            Write-Host "Changing SolarWinds MSP PME Agent to Automatic" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source "Doherty Associates" -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source "Doherty Associates" -EntryType Information -EventID 102 -Message "Setting SolarWinds MSP PME Agent to Automatic...`nSource: Get-PMEServices.ps1"
            set-service -Name "SolarWinds.MSP.PME.Agent.PmeService" -StartupType Automatic
        }

        if ($SolarWinds.MSP.RpcServerServiceStartup -ne 'Auto') {
            Write-Host "Changing SolarWinds MSP RPC Server to Automatic" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source "Doherty Associates" -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source "Doherty Associates" -EntryType Information -EventID 102 -Message "Setting SolarWinds MSP RPC Server Service to Automatic...`nSource: Get-PMEServices.ps1"
            set-service -Name "SolarWinds MSP RPC Server" -StartupType Automatic
        }

        if ($SolarWinds.MSP.CacheServiceStartup -ne 'Auto') {
            Write-Host "Changing SolarWinds MSP Cache Service Service to Automatic" -ForegroundColor Yellow
            New-EventLog -LogName Application -Source "Doherty Associates" -erroraction silentlycontinue
            Write-EventLog -LogName Application -Source "Doherty Associates" -EntryType Information -EventID 102 -Message "Setting SolarWinds MSP Cache Service to Automatic...`nSource: Get-PMEServices.ps1"
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
    $StatusMessage = 'All PME Services are running the latest version'    
    Write-Host "`n$StatusMessage" -foregroundcolor Green
}
else {
    $OverallStatus = 2
    $StatusMessage = 'One or more PME Services are not running the latest version'
    Write-Host "`n$StatusMessage`n" -foregroundcolor Yellow
    
}
Write-Host "Status: $OverallStatus"
}

#endregion

. Get-PMEServicesStatus
. Get-PMEServiceVersions
. Get-PMEProfile
. Get-LatestPMEVersion
. Validate-PME


if ($RecheckStartup -eq $True) {
 . Get-PMEServicesStartup   
 . Write-Startup
}

if ($RecheckStatus -eq $True) {
 . Get-PMEServicesStatus   
 . Write-Status
}

if (($OverallStatus -ne '0') -or ($Diagnostics)) {
. Test-PMEConnectivity
# Write-Host "$DiagnosticsInfo`n"
# Write-Host "$DiagnosticsError"
Write-Host "Diagnostics Error: " -nonewline; Write-Host "$DiagnosticsErrorInt" -ForegroundColor Green
}
