<#    
    ************************************************************************************************************
    Name: Get-PMEServices.ps1
    Version: 0.1.4.1 (16th March 2020)
    Purpose:    Get/Reset PME Service Details
    Pre-Reqs:    Powershell 2
    + Added Better Detection for PME Services being missing on a device
    ************************************************************************************************************
#>
$Version = '0.1.4.1 (16th March 2020)'
$RecheckStartup = $Null
$RecheckStatus = $Null
$latestVersion = '1.1.12'

Write-Host "Get-PMEServices $Version"

#region Functions

Function Get-PMEServicesStatus {
$SolarWindsMSPCacheStatus = (get-service "SolarWinds.MSP.CacheService" -ErrorAction SilentlyContinue).Status
$SolarWindsMSPPMEAgentStatus = (get-service "SolarWinds.MSP.PME.Agent.PmeService" -ErrorAction SilentlyContinue).Status
$SolarWindsMSPRpcServerStatus = (get-service "SolarWinds.MSP.RpcServerService" -ErrorAction SilentlyContinue).status
}

Function Get-PMEServicesVersions {

    if ($SolarWindsMSPCacheStatus -eq $null) {
        Write-Host "PME Cache service is missing" -ForegroundColor Red
        $SolarWindsMSPCacheStatus = 'Service is Missing'
        $PMECacheVersion = '0.0'
    }
    else {
        $SolarWindsMSPCacheLocation = (get-wmiobject win32_service -filter "Name like 'SolarWinds.MSP.CacheService'" -ErrorAction SilentlyContinue).PathName.Replace('"','')
        $PMECacheVersion = (get-item $SolarWindsMSPCacheLocation).VersionInfo.ProductVersion
    }

    if ($SolarWindsMSPPMEAgentStatus -eq $null) { 
        Write-Host "PME Agent service is missing" -ForegroundColor Red
        $SolarWindsMSPPMEAgentStatus = 'Service is Missing'
        $PMEAgentVersion = '0.0'
    }
    else {
        $SolarWindsMSPPMEAgentLocation = (get-wmiobject win32_service -filter "Name like 'SolarWinds.MSP.PME.Agent.PmeService'" -ErrorAction SilentlyContinue).Pathname.Replace('"','')
        $PMEAgentVersion = (get-item $SolarWindsMSPPMEAgentLocation).VersionInfo.ProductVersion
    }

    if ($SolarWindsMSPRpcServerStatus -eq $null) {
        Write-Host "PME RPC Server service is missing" -ForegroundColor Red
        $SolarWindsMSPRpcServerStatus = 'Service is Missing'
        $PMERpcServerVersion = '0.0'
    }
    else {
        $SolarWindsMSPRpcServerLocation = (get-wmiobject win32_service -filter "Name like 'SolarWinds.MSP.RpcServerService'" -ErrorAction SilentlyContinue).Pathname.Replace('"','')
        $PMERpcServerVersion = (get-item $SolarWindsMSPRpcServerLocation).VersionInfo.ProductVersion
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
. Get-PMEServicesVersions

# . Write-Status
# . Write-Version
. Validate-PME

# . Start-Services
# . Set-AutomaticStartup



if ($RecheckStartup -eq $True) {
 . Get-PMEServicesStartup   
 . Write-Startup
}

if ($RecheckStatus -eq $True) {
 . Get-PMEServicesStatus   
 . Write-Status
}
