<#
   *********************************************************************************************************************************
    Name:            Repair-PME.ps1
    Version:         0.2.0.0 (03/03/2021)
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
                               but has not updated yet. Change $RepairAfterUpdateDays = "2" variable to number of desired days
                               to force repair after new version of PME released. Change to $RepairAfterUpdateDays = "0" if you
                               want this script to run without safeguards. This new function means this script can finally be
                               used for self-healing.
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
                     0.1.7.2 - Updated functions 'Read-PMEConfig' and 'Set-PMEConfig' to handle situations where an exception
                               could be thrown stopping script execution when config files are empty.
                     0.1.7.3 - New function 'Get-NCAgentVersion' which checks and provides information on currently installed
                               N-Central agent. If agent is not installed or incompatible with PME script will be aborted.
                             - Move function 'Set-Start' to start after confirming elevation due to issue where event log write
                               could not be saved if no administrator rights.
                             - Updated function 'Confirm-Elevation' to remove event log write as no administrator rights will
                               cause an additional unwanted error.
                             - Updated event log writing to more accurately record the event level rather than defaulting to
                               information throughout the script.
                     0.1.7.4 - New function 'Restore-Date' to fix issue where the install month of a program is unusually
                               represented as a single digit in the registry. Thanks to Casper Stekelenburg for the provided code.
                             - Functions 'Get-NCAgentVersion' and 'Confirm-PMEInstalled' updated to call this function.
                     0.1.7.5 - Update 'Install-PME' function to fix issue where $datetime variable was missing which caused PME
                               install log files to not be timestamped.
                     0.1.8.0 - Moved 'Test-Connectivity' function to start earlier on in the script to prevent errors.
                             - New function 'Test-Port' to allow connectivity testing on legacy OS in 'Test-Connectivity' function.
                             - Updated Function 'Test-Connectivity' to support testing on Windows 7 and Server 2008 R2.
                             - New functions 'Get-Certificate' and 'Test-Certificate' to test certificate issues. Thanks to
                               David Brooks for feedback on this.
                             - Updated 'Get-PMEConfigurationDetails' function to support HTTPS to HTTP fallback.
                             - New Function 'Set-CryptoProtocol' to enable connections via TLS 1.2 ready for when/if Solarwinds
                               turn off older protocols such as TLS 1.0 on sis.n-able.com.
                             - Updated function 'Confirm-PMEUpdatePending' to fix issue where it would not correctly report
                               elapsed days since PME has been released in situations where an update has just been released.
                             - Fixed some minor grammatical errors throughout script.
                     0.1.8.1 - Fixed issue in 'Confirm-PMEUpdatePending' function to avoid using the PMEWrapper.log to detect if
                               install is pending as this is not reliable and caused the script to force the update rather than
                               gracefully wait until the $RepairAfterUpdateDays variable has passed.
                     0.1.8.2 - Incorporated changes made by Casper Stekelenburg of ICT-Concept B.V. :-
                               - Applied parameter splatting in a number of places to reduce length of commands that were too 
                                 long to fit on a 24 inch display.
                               - Replaced Write-Host "Warning: ...." with Write-Warning "..."
                               - Modified Get-PMEConfig and Set-PMEConfig to use [xml] type casting
                               - Replaced double full file paths with variables
                               - Turned NCPM-4407 explanation in Set-PMEConfig into a comment-block rather than a single line.
                             - Renamed functions 'Get-Certificate' and 'Test-Certificate' to 'Get-SWCertificate' and 
                               'Test-SWCertificate' to avoid warnings with cmdlets with the same name in PS 5.1.#
                             - Updated 'Get-NCAgentVersion' and 'Confirm-PMEInstalled' functions to account for edge case 
                               situation where install date may not be present causing the script to halt with an exception.
                     0.1.8.3 - New function 'Confirm-PMEServices' to check for edge cases where PME installer doesn't 
                               successfully install services but reports a successful exit code. Thanks to Prejay Shah
                               for reporting and code used from his Get-PMEService script for the new function.
                             - Updated functions 'Get-PMEConfig' and 'Set-PMEConfig' with a try/catch to allow the script
                               to continue if the xml files cannot be read i.e they are corrupt. The script will warn and
                               the reinstall should replace the files anyway. Thanks to Webster Massingham for reporting
                               and suggestion.
                     0.1.8.4 - Fixed issue in 'Confirm-PMEUpdatePending' function where it was not correctly comparing 
                               against older versions of PME causing the script to halt. [version] type casting now used.
                               Thanks to Webster Massingham for reporting and suggestion.
                     0.1.9.0 - New function 'Confirm-PMERecentInstall' allows successful repair/self-healing when an auto-update
                               fails during the update pending window. It will now force repair if installed within the last 2 days.
                               This is controlled via the $ForceRepairRecentInstallDays variable in the settings section.
                             - Updated 'Confirm-PMEUpdatePending', 'Confirm-PMEInstalled' and 'Get-OSArch' functions to support the 
                               new 'Confirm-PMERecentInstall' function.
                             - 'Confirm-PMEInstalled' function now also displays PME Cache Service and PME RPC Server Service info.
                             - Updated 'Restore-Date' function to account for situations where install dates are in YYYYMd format.
                               Thanks to David Brooks for reporting this issue.
                             - Minor output adjustments for readability.
                             - Depreciated 'Set-PMEConfig' function and removed $NCPM4407 variable as fix no longer required.
                     0.2.0.0 - Updated 'Get-PMESetupDetails' function to fix critical issue where the PME release date was incorrect.
                               It will now obtain this by getting the LastModified time of the PMESetup_details.xml instead.
                             - Removed 'Get-PMEConfigurationDetails' function, code merged into 'Get-PMESetupDetails' function.
                             - Updated 'Confirm-PMEUpdatePending' function to support changes in 'Get-PMESetupDetails' function.
                             - New function 'Get-RepairPMEUpdate' to perform an update check of the Repair-PME script. By default
                               this is turned on and is controlled via the $UpdateCheck variable in the settings section. If you
                               are using self healing/scripting from within N-Central, notifications should be setup to notify on
                               'Failure' and 'Send task output file in to email recipients' to alert if it is out of date. 
                               Connectivity to https://raw.githubusercontent.com is required for this feature.
                             - Renamed 'Get-PMEConfig' function to 'Get-PMEConfigMisconfigurations'.         
   *********************************************************************************************************************************
#>
$Version = '0.2.0.0'
$VersionDate = '(03/03/2021)'

# Settings
# ********************************************************************************************************************************
# Change this variable to number of days (must be a number!) to allow repair after new version of PME is released. 
# This is used for the update pending check. Default is 2.
$RepairAfterUpdateDays = "2"

# Change this variable to number of days (must be a number!) within a recent install to allow a force repair. 
# This will bypass the update pending check. Default is 2. Ensure this is equal to $RepairAfterUpdateDays.
$ForceRepairRecentInstallDays = "2"

# Change this variable to turn off/on update check of the Repair-PME script. Default is Yes. To turn this off set it to No.
$UpdateCheck = "Yes"
# ********************************************************************************************************************************

Write-Host "Repair-PME $Version $VersionDate`n" -ForegroundColor Yellow

$WriteEventLogInformationParams = @{
    LogName   = "Application"
    Source    = "Repair-PME"
    EntryType = "Information"
    EventID   = 100
}
$WriteEventLogErrorParams = @{
    LogName   = "Application"
    Source    = "Repair-PME"
    EntryType = "Error"
    EventID   = 100
}
$WriteEventLogWarningParams = @{
    LogName   = "Application"
    Source    = "Repair-PME"
    EntryType = "Warning"
    EventID   = 100
}
Function Confirm-Elevation {
    # Confirms script is running as an administrator
    Write-Host "Checking for elevated permissions" -ForegroundColor Cyan
    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
                [Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Throw "Insufficient permissions to run this script. Run PowerShell as an administrator and run this script again."
    } Else {
        Write-Host "OK: Script is running as administrator" -ForegroundColor Green
    }
}
Function Set-Start {
    New-EventLog -LogName Application -Source "Repair-PME" -ErrorAction SilentlyContinue
    Write-EventLog @WriteEventLogInformationParams -Message "Repair-PME has started, running version $Version.`nScript: Repair-PME.ps1"
}

function Get-LegacyHash {
    Param($Path)
    # Performs hashing functionality with compatibility for older OS
    Try {
        Add-Type -AssemblyName System.Security
        $csp = New-Object -TypeName System.Security.Cryptography.SHA256CryptoServiceProvider
        $ComputedHash = @()
        $ComputedHash = $csp.ComputeHash([System.IO.File]::ReadAllBytes($Path))
        $ComputedHash = [System.BitConverter]::ToString($ComputedHash).Replace("-", "").ToLower()
        Return $ComputedHash
    } Catch {
        Write-EventLog @WriteEventLogErrorParams -Message "Unable to performing hashing, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
        Throw "Unable to performing hashing, aborting. Error: $($_.Exception.Message)"
    }
}
Function Get-OSVersion {
    # Get OS version
    $OSVersion = (Get-WmiObject Win32_OperatingSystem).Caption
    #Workaround for WMI timeout or WMI returning no data
    If (($null -eq $OSVersion) -or ($OSVersion -like "*OS - Alias not found*")) {
        $OSVersion = (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('ProductName')
    }
    Write-Output "OS: $OSVersion"
}
Function Get-OSArch {
    $OSArch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
    Write-Output "OS Architecture: $OSArch"
    If ($OSArch -like "*64*") {
        $NCentralLog = "C:\Program Files (x86)\N-able Technologies\Windows Agent\log"
        $PMEAgentUninstall = "C:\Program Files (x86)\SolarWinds MSP\PME\unins000.exe"
        $PMERPCUninstall = "C:\Program Files (x86)\SolarWinds MSP\RpcServer\unins000.exe"
        $PMECacheUninstall = "C:\Program Files (x86)\SolarWinds MSP\RpcServer\unins000.exe"
    } Else {
        $NCentralLog = "C:\Program Files\N-able Technologies\Windows Agent\log"
        $PMEAgentUninstall = "C:\Program Files\SolarWinds MSP\PME\unins000.exe"
        $PMERPCUninstall = "C:\Program Files\SolarWinds MSP\RpcServer\unins000.exe"
        $PMECacheUninstall = "C:\Program Files\SolarWinds MSP\RpcServer\unins000.exe"    
    }
}

Function Get-PSVersion {
    $PSVersion = $($PSVersionTable.PSVersion)
    Write-Output "PowerShell: $($PSVersionTable.PSVersion)"
}

Function Test-Port ($server, $port) {
    $client = New-Object Net.Sockets.TcpClient
    try {
        $client.Connect($server, $port)
        $true
    } catch {
        $false
    }
}

Function Set-CryptoProtocol {
    # Enable TLS 1.2 - this should work on Windows 7 with PowerShell 2.0 and above.
    $tls12 = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
    [Net.ServicePointManager]::SecurityProtocol = $tls12
}

Function Get-RepairPMEUpdate {
    If ($UpdateCheck -eq "Yes") {
        Write-Host "Checking if update is available for Repair-PME script..." -ForegroundColor Cyan    
        $RepairPMEVersionURI = "http://raw.githubusercontent.com/N-able/ScriptsAndAutomationPolicies/master/Repair-PME/LatestVersion.xml"
        Try { 
            $Request = $null; $LatestPMEVersion = $null           
            [xml]$request = ((New-Object System.Net.WebClient).DownloadString("$RepairPMEVersionURI") -split '<\?xml.*\?>')
            $LatestPMEVersion  = $request.LatestVersion.Version
            Write-Output "Current Repair-PME Version: $Version `nLatest Repair-PME Version: $LatestPMEVersion"
            If ([version]$Version -ge [version]$LatestPMEVersion) {
                Write-Host "OK: Repair-PME is up to date" -ForegroundColor Green
            }
            ElseIf ([version]$Version -lt [version]$LatestPMEVersion) {
                Write-EventLog @WriteEventLogWarningParams -Message "WARNING: Repair-PME is not up to date! please download the latest version from https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/Repair-PME/Repair-PME.ps1`nScript: Repair-PME.ps1"
                Write-Error "Repair-PME is not up to date! please download the latest version from https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/Repair-PME/Repair-PME.ps1"
            }
            Else {
                Write-EventLog @WriteEventLogWarningParams -Message "ERROR: Unable to detect if Repair-PME is up to date!`nScript: Repair-PME.ps1"
                Write-Error "Unable to detect if Repair-PME is up to date!"
            }
        }
        Catch [System.Net.WebException] {
            Write-EventLog @WriteEventLogWarningParams -Message "ERROR: Unable to fetch version file to check if Repair-PME is up to date, possibly due to limited or no connectivity to https://raw.githubusercontent.com`nScript: Repair-PME.ps1"
            Write-Error "Unable to fetch version file to check if Repair-PME is up to date, possibly due to limited or no connectivity to https://raw.githubusercontent.com"
        }
    }
} 

Function Test-Connectivity {
    # Performs connectivity tests to destinations required for PME
    If (($PSVersionTable.PSVersion -ge "4.0") -and (!($OSVersion -match 'Windows 7')) -and (!($OSVersion -match '2008 R2'))) {
        Write-Host "Performing HTTPS connectivity tests for PME required destinations..." -ForegroundColor Cyan
        $List1 = @("sis.n-able.com")
        $HTTPSError = @()
        $List1 | ForEach-Object {
            $Test1 = Test-NetConnection $_ -Port 443
            If ($Test1.tcptestsucceeded -eq $True) {
                Write-Host "OK: Connectivity to https://$_ ($(($Test1).RemoteAddress.IpAddressToString)) established" -ForegroundColor Green
                $HTTPSError += "No"
            } Else {
                Write-Host "ERROR: Unable to establish connectivity to https://$_ ($(($Test1).RemoteAddress.IpAddressToString))" -ForegroundColor Red
                $HTTPSError += "Yes"
            }
        }

        Write-Host "Performing HTTP connectivity tests for PME required destinations..." -ForegroundColor Cyan
        $HTTPError = @()
        $List2 = @("sis.n-able.com", "download.windowsupdate.com", "fg.ds.b1.download.windowsupdate.com")
        $List2 | ForEach-Object {
            $Test1 = Test-NetConnection $_ -Port 80
            If ($Test1.tcptestsucceeded -eq $True) {
                Write-Host "OK: Connectivity to http://$_ ($(($Test1).RemoteAddress.IpAddressToString)) established" -ForegroundColor Green
                $HTTPError += "No"
            } Else {
                Write-Host "ERROR: Unable to establish connectivity to http://$_ ($(($Test1).RemoteAddress.IpAddressToString))" -ForegroundColor Red
                $HTTPError += "Yes"
            }
        }

        If (($HTTPError[0] -like "*Yes*") -and ($HTTPSError[0] -like "*Yes*")) {
            Write-EventLog @WriteEventLogErrorParams -Message "ERROR: No connectivity to $($List2[0]) can be established, aborting.`nScript: Repair-PME.ps1"
            Throw "ERROR: No connectivity to $($List2[0]) can be established, aborting."
        } ElseIf (($HTTPError[0] -like "*Yes*") -or ($HTTPSError[0] -like "*Yes*")) {
            Write-EventLog @WriteEventLogWarningParams -Message "WARNING: Partial connectivity to $($List2[0]) established, falling back to HTTP.`nScript: Repair-PME.ps1"
            Write-Warning "Partial connectivity to $($List2[0]) established, falling back to HTTP"
            $Fallback = "Yes"
        }

        If ($HTTPError[1] -like "*Yes*") {
            Write-EventLog @WriteEventLogWarningParams -Message "WARNING: No connectivity to $($List2[1]) can be established, you will be unable to download Microsoft Updates!`nScript: Repair-PME.ps1"
            Write-Warning "No connectivity to $($List2[1]) can be established, you will be unable to download Microsoft Updates!"
        }

        If ($HTTPError[2] -like "*Yes*") {
            Write-EventLog @WriteEventLogWarningParams -Message "WARNING: No connectivity to $($List2[2]) can be established, you will be unable to download Windows Feature Updates!`nScript: Repair-PME.ps1"
            Write-Warning "No connectivity to $($List2[2]) can be established, you will be unable to download Windows Feature Updates!"
        }
    } Else {
        Write-Host "Performing HTTPS connectivity tests for PME required destinations using legacy method..." -ForegroundColor Cyan
        $List1 = @("sis.n-able.com")
        $HTTPSError = @()
        $List1 | ForEach-Object {
            $Test1 = Test-Port $_ 443
            If ($Test1 -eq $True) {
                Write-Host "OK: Connectivity to https://$_ established" -ForegroundColor Green
                $HTTPSError += "No"
            } Else {
                Write-Host "ERROR: Unable to establish connectivity to https://$_ established" -ForegroundColor Red
                $HTTPSError += "Yes"
            }
        }

        Write-Host "Performing HTTP connectivity tests for PME required destinations using legacy method..." -ForegroundColor Cyan
        $HTTPError = @()
        $List2 = @("sis.n-able.com", "download.windowsupdate.com", "fg.ds.b1.download.windowsupdate.com")
        $List2 | ForEach-Object {
            $Test1 = Test-Port $_ 80
            If ($Test1 -eq $True) {
                Write-Host "OK: Connectivity to http://$_ established" -ForegroundColor Green
                $HTTPError += "No"
            } Else {
                Write-Host "ERROR: Unable to establish connectivity to http://$_ established" -ForegroundColor Red
                $HTTPError += "Yes"
            }
        }

        If (($HTTPError[0] -like "*Yes*") -and ($HTTPSError[0] -like "*Yes*")) {
            Write-EventLog @WriteEventLogErrorParams -Message "ERROR: No connectivity to $($List2[0]) can be established, aborting.`nScript: Repair-PME.ps1"
            Throw "ERROR: No connectivity to $($List2[0]) can be established, aborting."
        } ElseIf (($HTTPError[0] -like "*Yes*") -or ($HTTPSError[0] -like "*Yes*")) {
            Write-EventLog @WriteEventLogWarningParams -Message "WARNING: Partial connectivity to $($List2[0]) established, falling back to HTTP.`nScript: Repair-PME.ps1"
            Write-Warning "Partial connectivity to $($List2[0]) established, falling back to HTTP"
            $Fallback = "Yes"
        }

        If ($HTTPError[1] -like "*Yes*") {
            Write-EventLog @WriteEventLogWarningParams -Message "WARNING: No connectivity to $($List2[1]) can be established, you will be unable to download Microsoft Updates!`nScript: Repair-PME.ps1"
            Write-Warning "No connectivity to $($List2[1]) can be established, you will be unable to download Microsoft Updates!"
        }

        If ($HTTPError[2] -like "*Yes*") {
            Write-EventLog @WriteEventLogWarningParams -Message "WARNING: No connectivity to $($List2[2]) can be established, you will be unable to download Windows Feature Updates!`nScript: Repair-PME.ps1"
            Write-Warning "No connectivity to $($List2[2]) can be established, you will be unable to download Windows Feature Updates!"
        }
    }
}

Function Get-SWCertificate ($url) {
    [net.httpWebRequest] $WebRequest = [Net.WebRequest]::Create($url)
    $WebRequest.AllowAutoRedirect = $true
    $WebRequest.KeepAlive = $false
    $WebRequest.Timeout = 10000
    $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

    #Request website
    try {
        $Response = $WebRequest.GetResponse()
        $Response.close()
    } catch {
    }

    #Creates Certificate
    $Certificate = $WebRequest.ServicePoint.Certificate.Handle
    $Issuer = $WebRequest.ServicePoint.Certificate.Issuer
    $Subject = $WebRequest.ServicePoint.Certificate.Subject

    #Build chain
    $chain.Build($Certificate) | Out-Null
    #write-host $chain.ChainElements.Count #This returns "1" meaning none of the CA certs are included.
    #write-host $chain.ChainElements[0].Certificate.IssuerName.Name
    [Net.ServicePointManager]::ServerCertificateValidationCallback = $null

    $CertificateChain = $chain.ChainElements.Certificate | Select-Object -Property DnsNameList, NotAfter
    $CertificateChain = $chain.ChainElements | Select-Object -ExpandProperty Certificate | Select-Object Subject, NotAfter
}

Function Test-SWCertificate {
    If ($null -eq $Fallback) {
        Write-Host "Checking certificate chain for sis.n-able.com..." -ForegroundColor Cyan
        . Get-SWCertificate https://sis.n-able.com
        $Date = Get-Date
        $CertificateChain | ForEach-Object {
            If ($null -eq $($_.NotAfter)) {
                Write-EventLog @WriteEventLogErrorParams -Message "Unable to obtain certificate chain, PME may have trouble downloading from https://sis.n-able.com, aborting.`nScript: Repair-PME.ps1"
                Throw "Unable to obtain certificate chain, PME may have trouble downloading from https://sis.n-able.com, aborting."
            } ElseIf ($($_.NotAfter) -le $Date) {
                Write-Host "$($_.NotAfter)"
                Write-EventLog @WriteEventLogErrorParams -Message "Certificate for ($($_.Subject)) expired on $($_.NotAfter) PME may have trouble downloading from https://sis.n-able.com, aborting.`nScript: Repair-PME.ps1"
                Throw "Certificate for ($($_Subject)) expired on $($_.NotAfter) PME may have trouble downloading from https://sis.n-able.com, aborting."
            } Else {
                Write-Host "OK: Certificate for ($($_.Subject)) is valid"  -ForegroundColor Green
            }
        }
    }
}

Function Restore-Date {
    If ($InstallDate.Length -eq 6) {
        $M = $InstallDate.Substring(4, 1)
        $d = $InstallDate.Substring(5, 1)
        $Year = $InstallDate.Substring(0, 4)
        $InstallDate = $($Year + "0" + $M +"0" + $d )
    }
    If ($InstallDate.Length -eq 7) {
        $MMdd = $InstallDate.Substring(4, 3)
        $Year = $InstallDate.Substring(0, 4)
        $InstallDate = $($Year + "0" + $MMdd)
    }
}

Function Get-NCAgentVersion {
    # Check if N-Central Agent is currently installed
    If ($OSArch -like '*64*') {
        Write-Host "Checking if N-Central Agent is already installed..." -ForegroundColor Cyan
        $PATHS = @("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
        $SOFTWARE = "Windows Agent"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName, DisplayVersion, Publisher, InstallDate

            If ($null -ne $installed) {
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "Windows Agent" -and $($app.Publisher) -eq "N-able Technologies") {
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsNCAgentInstalled = "Yes"
                        Write-Host "N-Central Agent Installed: Yes" -ForegroundColor Green
                        Write-Output "N-Central Agent Version: $($app.DisplayVersion)"
                        Write-Output "N-Central Agent Install Date: $InstallDateFormatted"
                        If ($($app.DisplayVersion) -ge "12.2.0.274") {
                            Write-Host "N-Central Agent PME Compatible: Yes" -ForegroundColor Green
                        } Else {
                            Write-Host "N-Central Agent PME Compatible: No" -ForegroundColor Red
                            Write-EventLog @WriteEventLogErrorParams -Message "Installed N-Central Agent ($($app.DisplayVersion)) is not compatible with PME, aborting.`nScript: Repair-PME.ps1"
                            Throw "Installed N-Central Agent ($($app.DisplayVersion)) is not compatible with PME, aborting."
                        }
                    }
                }
            } Else {
                $IsNCAgentInstalled = "No"
                Write-Host "N-Central Agent Installed: No" -ForegroundColor Red
                Write-EventLog @WriteEventLogErrorParams -Message "N-Central Agent is not installed, PME requires an agent, aborting.`nScript: Repair-PME.ps1"
                Throw "N-Central Agent is not installed, PME requires an agent, aborting."
            }
        }
    }

    If ($OSArch -like '*32*') {
        Write-Host "Checking if N-Central Agent is already installed..." -ForegroundColor Cyan
        $PATHS = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
        $SOFTWARE = "Windows Agent"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName, DisplayVersion, Publisher, InstallDate

            If ($null -ne $installed) {
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "Windows Agent" -and $($app.Publisher) -eq "N-able Technologies") {
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsNCAgentInstalled = "Yes"
                        Write-Host "N-Central Agent Installed: Yes" -ForegroundColor Green
                        Write-Output "N-Central Agent Version: $($app.DisplayVersion)"
                        Write-Output "N-Central Agent Install Date: $InstallDateFormatted"
                        If ($($app.DisplayVersion) -ge "12.2.0.274") {
                            Write-Host "N-Central Agent PME Compatible: Yes" -ForegroundColor Green
                        } Else {
                            Write-Host "N-Central Agent PME Compatible: No" -ForegroundColor Red
                            Write-EventLog @WriteEventLogErrorParams -Message "Installed N-Central Agent ($($app.DisplayVersion)) is not compatible with PME, aborting.`nScript: Repair-PME.ps1"
                            Throw "Installed N-Central Agent ($($app.DisplayVersion)) is not compatible with PME, aborting."
                        }
                    }
                }
            } Else {
                $IsNCAgentInstalled = "No"
                Write-Host "N-Central Agent Installed: No" -ForegroundColor Red
                Write-EventLog @WriteEventLogErrorParams -Message "N-Central Agent is not installed, PME requires an agent, aborting.`nScript: Repair-PME.ps1"
                Throw "N-Central Agent is not installed, PME requires an agent, aborting."
            }
        }
    }
}

Function Confirm-PMEInstalled {
    # Check if PME Agent is currently installed
    If ($OSArch -like '*64*') {
        Write-Host "Checking if PME Agent is already installed..." -ForegroundColor Cyan
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
                        $PMEAgentAppDisplayVersion = $($app.DisplayVersion)
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsPMEAgentInstalled = "Yes"
                        Write-Host "PME Agent Already Installed: Yes" -ForegroundColor Green
                        Write-Output "Installed PME Agent Version: $PMEAgentAppDisplayVersion"
                        Write-Output "Installed PME Agent Date: $InstallDateFormatted"
                    }
                }
            } Else {
                $IsPMEAgentInstalled = "No"
                Write-Host "PME Agent Already Installed: No" -ForegroundColor Yellow
            }
        }
    }

    If ($OSArch -like '*32*') {
        Write-Host "Checking if PME Agent is already installed..." -ForegroundColor Cyan
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
                        $PMEAgentAppDisplayVersion = $($app.DisplayVersion) 
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsPMEAgentInstalled = "Yes"
                        Write-Host "PME Agent Already Installed: Yes" -ForegroundColor Green
                        Write-Output "Installed PME Agent Version: $PMEAgentAppDisplayVersion"
                        Write-Output "Installed PME Agent Date: $InstallDateFormatted"
                    }
                }
            } Else {
                $IsPMEAgentInstalled = "No"
                Write-Host "PME Agent Already Installed: No" -ForegroundColor Yellow
            }
        }
    }

    # Check if PME RPC Service is currently installed
    If ($OSArch -like '*64*') {
        Write-Host "Checking if PME RPC Server Service is already installed..." -ForegroundColor Cyan
        $PATHS = @("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
        $SOFTWARE = "SolarWinds MSP RPC Server"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName, DisplayVersion, InstallDate

            If ($null -ne $installed) {
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "Solarwinds MSP RPC Server") {
                        $PMERPCServerAppDisplayVersion = $($app.DisplayVersion) 
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsPMERPCServerServiceInstalled = "Yes"
                        Write-Host "PME RPC Server Service Already Installed: Yes" -ForegroundColor Green
                        Write-Output "Installed PME RPC Server Service Version: $PMERPCServerAppDisplayVersion"
                        Write-Output "Installed PME RPC Server Service Date: $InstallDateFormatted"
                    }
                }
            } Else {
                $IsPMERPCServerServiceInstalled = "No"
                Write-Host "PME RPC Server Service Already Installed: No" -ForegroundColor Yellow
            }
        }
    }

    If ($OSArch -like '*32*') {
        Write-Host "Checking if PME is already installed..." -ForegroundColor Cyan
        $PATHS = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
        $SOFTWARE = "SolarWinds MSP RPC Server"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName, DisplayVersion, InstallDate

            If ($null -ne $installed) {
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "SolarWinds MSP RPC Server") {
                        $PMERPCServerAppDisplayVersion = $($app.DisplayVersion) 
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsPMERPCServerServiceInstalled = "Yes"
                        Write-Host "PME RPC Server Service Already Installed: Yes" -ForegroundColor Green
                        Write-Output "Installed PME RPC Server Service Version: $PMERPCServerAppDisplayVersion"
                        Write-Output "Installed PME RPC Server Service Date: $InstallDateFormatted"
                    }
                }
            } Else {
                $IsPMERPCServerServiceInstalled = "No"
                Write-Host "PME RPC Server Service Already Installed: No" -ForegroundColor Yellow
            }
        }
    }

    # Check if PME Cache Service is currently installed
    If ($OSArch -like '*64*') {
        Write-Host "Checking if PME RPC Server Service is already installed..." -ForegroundColor Cyan
        $PATHS = @("HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
        $SOFTWARE = "SolarWinds MSP Cache Service"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName, DisplayVersion, InstallDate

            If ($null -ne $installed) {
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "SolarWinds MSP Cache Service") {
                        $PMECacheServiceAppDisplayVersion = $($app.DisplayVersion) 
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsPMECacheServiceInstalled = "Yes"
                        Write-Host "PME Cache Service Already Installed: Yes" -ForegroundColor Green
                        Write-Output "Installed PME Cache Service Version: $PMECacheServiceAppDisplayVersion"
                        Write-Output "Installed PME Cache Service Date: $InstallDateFormatted"
                    }
                }
            } Else {
                $IsPMECacheServiceInstalled = "No"
                Write-Host "PME Cache Service Already Installed: No" -ForegroundColor Yellow
            }
        }
    }

    If ($OSArch -like '*32*') {
        Write-Host "Checking if PME RPC Server Service is already installed..." -ForegroundColor Cyan
        $PATHS = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
        $SOFTWARE = "SolarWinds MSP Cache Service"
        ForEach ($path in $PATHS) {
            $installed = Get-ChildItem -Path $path |
            ForEach-Object { Get-ItemProperty $_.PSPath } |
            Where-Object { $_.DisplayName -match $SOFTWARE } |
            Select-Object -Property DisplayName, DisplayVersion, InstallDate

            If ($null -ne $installed) {
                ForEach ($app in $installed) {
                    If ($($app.DisplayName) -eq "SolarWinds MSP Cache Service") {
                        $PMECacheServiceAppDisplayVersion = $($app.DisplayVersion)
                        $InstallDate = $($app.InstallDate)
                        If ($null -ne $InstallDate -and $InstallDate -ne "") {
                            . Restore-Date
                            $ConvertDateTime = [DateTime]::ParseExact($InstallDate, "yyyyMMdd", $null)
                            $InstallDateFormatted = $ConvertDateTime | Get-Date -Format "yyyy.MM.dd"
                        }
                        $IsPMECacheServiceInstalled = "Yes"
                        Write-Host "PME Cache Service Already Installed: Yes" -ForegroundColor Green
                        Write-Output "Installed PME Cache Service Version: $PMECacheServiceAppDisplayVersion"
                        Write-Output "Installed PME Cache Service Date: $InstallDateFormatted"
                    }
                }
            } Else {
                $IsPMECacheServiceInstalled = "No"
                Write-Host "PME Cache Service Already Installed: No" -ForegroundColor Yellow
            }
        }
    }            
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
        Write-EventLog @WriteEventLogErrorParams -Message "Error fetching PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
        Throw "Error fetching PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message)"
    } Catch [System.Management.Automation.MetadataException] {
        Write-EventLog @WriteEventLogErrorParams -Message "Error casting to XML, could not parse PMESetup_details.xml, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
        Throw "Error casting to XML, could not parse PMESetup_details.xml, aborting. Error: $($_.Exception.Message)"
    } Catch {
        Write-EventLog @WriteEventLogErrorParams -Message "Error occurred attempting to obtain PMESetup_details.xml, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
        Throw "Error occurred attempting to obtain PMESetup_details.xml, aborting. Error: $($_.Exception.Message)"
    }

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
        Write-EventLog @WriteEventLogErrorParams -Message "Error fetching header for PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
        Throw "Error fetching header for PMESetup_Details.xml, check the source URL $($PMESetup_detailsURI), aborting. Error: $($_.Exception.Message)"
    } Catch {
        Write-EventLog @WriteEventLogErrorParams -Message "Error fetching header for PMESetup_Details.xml, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
        Throw "Error fetching header for PMESetup_Details.xml, aborting. Error: $($_.Exception.Message)"
    }

    Write-Host "Checking Latest PME version..." -ForegroundColor Cyan
    Write-Output "Latest PME Version: $LatestVersion"
    Write-Output "Latest PME Release Date: $PMEReleaseDate"    
}

Function Confirm-PMERecentInstall {
    If (($IsPMEAgentInstalled -eq "Yes") -or ($IsPMERPCServerServiceInstalled -eq "Yes") -or ($IsPMECacheServiceInstalled -eq "Yes")) {
        $Date = Get-Date -Format 'yyyy.MM.dd'
        $InstallDatePMEAgent = (Get-Item $PMEAgentUninstall).LastWriteTime
        $InstallDatePMERPC = (Get-Item $PMERPCUninstall).LastWriteTime
        $InstallDatePMECache = (Get-Item $PMECacheUninstall).LastWriteTime
        $DaysInstalledPMEAgent = (New-TimeSpan -Start $InstallDatePMEAgent -End $Date).Days
        $DaysInstalledPMERPC = (New-TimeSpan -Start $InstallDatePMERPC -End $Date).Days
        $DaysInstalledPMECache  = (New-TimeSpan -Start $InstallDatePMECache -End $Date).Days
        
        Write-Host "INFO: Repair-PME will force repair without update pending check if PME was installed in the last ($ForceRepairRecentInstallDays) days" -ForegroundColor Yellow -BackgroundColor Black
        If (($DaysInstalledPMEAgent -le $ForceRepairRecentInstallDays) -or ($DaysInstalledPMERPC -le $ForceRepairRecentInstallDays) -or ($DaysInstalledPMECache -le $ForceRepairRecentInstallDays)) {
            Write-Output "Less than ($ForceRepairRecentInstallDays) days has elapsed since PME has been installed. No update pending check required."
            $BypassUpdatePendingCheck = "Yes"
        }
        Else {
            Write-Output "More than ($ForceRepairRecentInstallDays) days has elapsed since PME has been installed. Update pending check required."
            $BypassUpdatePendingCheck = "No"
        }    
    }
}

Function Confirm-PMEUpdatePending {
    # Check if PME is awaiting update for new release but has not updated yet (normally within 48 hours)
    If (($IsPMEAgentInstalled -eq "Yes") -and ($BypassUpdatePendingCheck -eq "No")) {
        $Date = Get-Date -Format 'yyyy.MM.dd'
        $ConvertPMEReleaseDate = Get-Date "$PMEReleaseDate"
        $SelfHealingDate = $ConvertPMEReleaseDate.AddDays($RepairAfterUpdateDays).ToString('yyyy.MM.dd')
        Write-Host "Checking if PME update pending..." -ForegroundColor Cyan
        Write-Host "INFO: Script will proceed ($RepairAfterUpdateDays) days after a new version of PME has been released" -ForegroundColor Yellow -BackgroundColor Black
        $DaysElapsed = (New-TimeSpan -Start $SelfHealingDate -End $Date).Days
        $DaysElapsedReversed = (New-TimeSpan -Start $ConvertPMEReleaseDate -End $Date).Days

        # Only run if current $Date is greater than or equal to $SelfHealingDate and $LatestVersion is greater than or equal to $app.DisplayVersion
        If (($Date -ge $SelfHealingDate) -and ([version]$LatestVersion -ge [version]$PMEAgentAppDisplayVersion)) {
            Write-Output "($DaysElapsed) days has elapsed since a new version of PME has been released and is allowed to be installed, script will proceed."
        } Else {
            Write-EventLog @WriteEventLogWarningParams -Message "($DaysElapsedReversed) days has elapsed since a new version of PME has been released, PME will only install after ($RepairAfterUpdateDays) days, aborting.`nScript: Repair-PME.ps1"
            Throw "($DaysElapsedReversed) days has elapsed since a new version of PME has been released, PME will only install after ($RepairAfterUpdateDays) days, aborting."
            Break
        }
    } ElseIf ($BypassUpdatePendingCheck -eq "Yes"){
        Write-Warning "Skipping update pending check as PME has recently been installed"
    } Else {
        Write-Warning "Skipping update pending check as PME is not currently installed"
    }
}

Function Invoke-SolarwindsDiagnostics {
    # Invokes Solarwinds official diagnostics tool to capture logs for support
    $ZipPath = "/`"ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs/SWDiagnostics$(Get-Date -Format 'yyyyMMdd-hhmmss').zip`""
    If ($OSArch -like '*64*') {
        # 32-bit program files on 64-bit
        $SolarwindsDiagnosticsFolderPath = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)") + "\SolarWinds MSP\PME\Diagnostics"
        $SolarwindsDiagnosticsExePath = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)") + "\SolarWinds MSP\PME\Diagnostics\SolarwindsDiagnostics.exe"
        If (Test-Path $SolarwindsDiagnosticsExePath) {
            Write-Host "Checking Diagnostics..." -ForegroundColor Cyan
            Write-Output "Solarwinds Diagnostics located at '$SolarwindsDiagnosticsExePath'"
            If (Test-Path "C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs") {
                Write-Output "Directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' already exists, no need to create directory"
            } Else {
                Try {
                    Write-Output "Directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' does not exist, creating directory"
                    New-Item -ItemType Directory -Path "C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs" -Force | Out-Null
                } Catch {
                    Write-EventLog @WriteEventLogErrorParams -Message "Unable to create directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' required for saving log capture. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
                    Throw "Unable to create directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' required for saving log capture. Error: $($_.Exception.Message)"
                }
            }
            Write-Output "Starting Solarwinds Diagnostics"
            #Write-Output "DEBUG: Solarwinds Diagnostics started with:- Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait"
            Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait
            Write-Output "Solarwinds Diagnostics completed, file saved to 'C:\ProgramData\SolarWinds MSP\Repair-PME\Diagnostic Logs'"
        } Else {
            Write-Warning "Unable to detect Solarwinds Diagnostics, skipping log capture"
        }
    } ElseIf ($OSArch -like '*32*') {
        # 32-bit program files on 32-bit
        $SolarwindsDiagnosticsFolderPath = [Environment]::GetEnvironmentVariable("ProgramFiles") + "\SolarWinds MSP\PME\Diagnostics"
        $SolarwindsDiagnosticsExePath = [Environment]::GetEnvironmentVariable("ProgramFiles") + "\SolarWinds MSP\PME\Diagnostics\SolarwindsDiagnostics.exe"
        If (Test-Path $SolarwindsDiagnosticsExePath) {
            Write-Output "Solarwinds Diagnostics located at '$SolarwindsDiagnosticsExePath'"
            If (Test-Path "C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs") {
                Write-Output "Directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs', no need to create directory"
            } Else {
                Try {
                    Write-Output "Directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' does not exist, creating directory"
                    New-Item -ItemType Directory -Path "C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs" -Force | Out-Null
                } Catch {
                    Write-EventLog @WriteEventLogErrorParams -Message "Unable to create directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' required for saving log capture. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
                    Throw "Unable to create directory 'C:\ProgramData/SolarWinds MSP/Repair-PME/Diagnostic Logs' required for saving log capture. Error: $($_.Exception.Message)"
                }
            }
            Write-Output "Starting Solarwinds Diagnostics"
            #Write-Output "DEBUG: Solarwinds Diagnostics started with:- Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait"
            Start-Process -FilePath "$SolarwindsDiagnosticsExePath" -ArgumentList "$ZipPath" -WorkingDirectory "$SolarwindsDiagnosticsFolderPath" -Verb RunAs -Wait
            Write-Output "Solarwinds Diagnostics completed, file saved to 'C:\ProgramData\SolarWinds MSP\Repair-PME\Diagnostic Logs'"
        } Else {
            Write-Warning "Unable to detect Solarwinds Diagnostics, skipping log capture"
        }
    } Else {
        Write-EventLog @WriteEventLogErrorParams -Message "Unable to detect processor architecture, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
        Throw "Unable to detect processor architecture, aborting. Error: $($_.Exception.Message)"
    }
}

Function Stop-PMESetup {
    # Kill any running instances of PMESetup.exe to ensure that we can download & install successfully
    Write-Host "Checking if PMESetup is currently running..." -ForegroundColor Cyan
    $PMESetupRunning = Get-Process PMESetup* -ErrorAction SilentlyContinue
    If ($PMESetupRunning) {
        Write-Warning "PMESetup is currently running, terminating"
        $PMESetupRunning | Stop-Process -Force
    } Else {
        Write-Host "OK: PMESetup is not currently running, proceeding" -ForegroundColor Green
    }

    Write-Host "Checking if CacheServiceSetup is currently running..." -ForegroundColor Cyan
    $PMESetupRunning = Get-Process CacheServiceSetup* -ErrorAction SilentlyContinue
    If ($PMESetupRunning) {
        Write-Warning "CacheServiceSetup is currently running, terminating"
        $PMESetupRunning | Stop-Process -Force
    } Else {
        Write-Host "OK: CacheServiceSetup is not currently running, proceeding" -ForegroundColor Green
    }

    Write-Host "Checking if RPCServerServiceSetup is currently running..." -ForegroundColor Cyan
    $PMESetupRunning = Get-Process RPCServerServiceSetup* -ErrorAction SilentlyContinue
    If ($PMESetupRunning) {
        Write-Warning "RPCServerServiceSetup is currently running, terminating"
        $PMESetupRunning | Stop-Process -Force
    } Else {
        Write-Host "OK: RPCServerServiceSetup is not currently running, proceeding" -ForegroundColor Green
    }

    Write-Host "Checking if _iu14D2N.tmp instances are currently running..." -ForegroundColor Cyan
    $PMESetupRunning = Get-Process _iu* -ErrorAction SilentlyContinue
    If ($PMESetupRunning) {
        Write-Warning "_iu14D2N.tmp instances are currently running, terminating"
        $PMESetupRunning | Stop-Process -Force
    } Else {
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
        } Else {
            Write-Warning "$Service still running, temporarily disabling recovery and terminating"
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
        } Else {
            Write-Warning "$Service still running, temporarily disabling recovery and terminating"
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
        } Else {
            Write-Warning "$Service still running, temporarily disabling recovery and terminating"
            #Set-Service -Name $Service -StartupType Disabled
            sc.exe failure "$Service" reset= 0 actions= // >null
            Stop-Process -Name $Process* -Force
            sc.exe failure "$Service" actions= restart/0/restart/0//0 reset= 0 >null
        }
    }
}

Function Clear-PME {
    Write-Host "Cleanup..." -ForegroundColor Cyan
    # Cleanup Solarwinds MSP Cache Service root folder
    If (Test-Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService") {
        Try {
            Write-Output "Performing cleanup of Solarwinds MSP Cache Service root folder"
            Remove-Item "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\*.*" -Force -Confirm:$false | Out-Null
        } Catch {
            Write-EventLog @WriteEventLogErrorParams -Message "Unable to cleanup 'C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\*.*' aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
            Throw "Unable to cleanup 'C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\*.*' aborting. Error: $($_.Exception.Message)"
        }
    } Else {
        Write-Output "Cleanup not required as Solarwinds MSP Cache Service root folder does not already exist"
    }
    # Cleanup Solarwinds MSP Cache Service cache folder
    If (Test-Path "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache") {
        Try {
            Write-Output "Performing cleanup of Solarwinds MSP Cache Service cache folder"
            Remove-Item "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache\*.*" -Force -Confirm:$false | Out-Null
        } Catch {
            Write-EventLog @WriteEventLogErrorParams -Message "Unable to cleanup C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache\*.*' aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
            Throw "Unable to cleanup C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache\*.*' aborting. Error: $($_.Exception.Message)"
        }
    } Else {
        Write-Output "Cleanup not required as Solarwinds MSP Cache Service cache folder does not already exist"
    }
}

Function Get-PMESetup {
    # Download Setup
    If ($Fallback -eq "Yes") {
        $FallbackDownloadURL = ($PMEDetails.DownloadURL).Replace('https', 'http')
        Write-Output "Begin download of current $($PMEDetails.FileName) version $($PMEDetails.Version) from sis.n-able.com"
        Try {
            (New-Object System.Net.WebClient).DownloadFile("$($FallbackDownloadURL)", "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe")
        } Catch {
            Write-EventLog @WriteEventLogErrorParams -Message "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
            Throw "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message)"
        }
    } Else {
        Write-Output "Begin download of current $($PMEDetails.FileName) version $($PMEDetails.Version) from sis.n-able.com"
        Try {
            (New-Object System.Net.WebClient).DownloadFile("$($PMEDetails.DownloadURL)", "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe")
        } Catch {
            Write-EventLog @WriteEventLogErrorParams -Message "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
            Throw "Unable to download $($PMEDetails.FileName) from sis.n-able.com, aborting. Error: $($_.Exception.Message)"
        }
    }
}

Function Get-PMEConfigMisconfigurations {
    # Check PME Config and inform of possible misconfigurations
    Write-Host "Checking PME Configuration..." -ForegroundColor Cyan
    Try {    
        $CacheServiceConfigFile = "C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\config\CacheService.xml"

        If (Test-Path "$CacheServiceConfigFile") {
            $xml = New-Object XML
            $xml.Load($CacheServiceConfigFile)
            $CacheServiceConfig = $xml.Configuration

            If ($null -ne $CacheServiceConfig) {
                If ($CacheServiceConfig.CanBypassProxyCacheService -eq "False") {
                    Write-Warning "Patch profile doesn't allow PME to fallback to external sources, if probe is not reachable PME may not work!"
                } ElseIf ($CacheServiceConfig.CanBypassProxyCacheService -eq "True") {
                    Write-Host "INFO: Patch profile allows PME to fallback to external sources" -ForegroundColor Yellow -BackgroundColor Black
                } Else {
                    Write-Warning "Unable to determine if patch profile allows PME to fallback to external sources"
                }

                If ($CacheServiceConfig.CacheSizeInMB -eq 10240) {
                    Write-Host "INFO: Cache Service is set to default cache size of 10240 MB" -ForegroundColor Yellow -BackgroundColor Black
                } Else {
                    $CacheSize = $CacheServiceConfig.CacheSizeInMB
                    Write-Warning "Cache Service is not set to default cache size of 10240 MB (currently $CacheSize MB), PME may not work at expected!"
                }
            } Else {
                Write-Warning "Cache Service config file is empty, skipping Cache Service settings checks"
            }
        } Else {
            Write-Warning "Cache Service config file does not exist, skipping Cache Service settings checks"
        }
    }    
    Catch {
        Write-Warning "Unable to read Cache Service config file as a valid xml file, default cache size can't be checked"
    }    
}

Function Set-PMEConfig {
    # Reserved for future use
    Write-Host "Setting PME Configuration..." -ForegroundColor Cyan
}

Function Install-PME {
    Write-Host "Install PME..." -ForegroundColor Cyan
    # Check Setup Exists in PME Archive Directory
    If (Test-Path "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe") {
        # Check Hash
        Write-Output "Checking hash of local file at 'C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe'"
        $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe"
        If ($Download -eq $($PMEDetails.SHA256Checksum)) {
            # Install
            Write-Output "Local copy of $($PMEDetails.FileName) is current and hash is correct"
            Write-Output "Installing $($PMEDetails.FileName) - logs will be saved to 'C:\ProgramData\Solarwinds MSP\Repair-PME\'"
            $StartProcessParams = @{
                FilePath     = "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe"
                ArgumentList = "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"C:\ProgramData\Solarwinds MSP\Repair-PME\Setup Log $DateTime.txt`""
                Wait         = $true
                PassThru     = $true
            }
            $Install = Start-Process @StartProcessParams
            If ($Install.ExitCode -eq 0) {
                Write-Host "OK: $($PMEDetails.Name) version $($PMEDetails.Version) successfully installed" -ForegroundColor Green
            } ElseIf ($Install.ExitCode -eq 5) {
                Write-EventLog @WriteEventLogErrorParams -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode).`nScript: Repair-PME.ps1"
                Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode)"
            } Else {
                Write-EventLog @WriteEventLogErrorParams -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'.`nScript: Repair-PME.ps1"
                Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'"
            }
        } Else {
            # Download
            Write-Output "Hash of local file ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, downloading the latest available version"
            . Get-PMESetup
            # Check Hash
            Write-Output "Checking hash of local file at 'C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe'"
            $Download = Get-LegacyHash -Path "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe"
            If ($Download -eq $($PMEDetails.SHA256Checksum)) {
                # Install
                Write-Output "Hash of file is correct"
                Write-Output "Installing $($PMEDetails.FileName) - logs will be saved to 'C:\ProgramData\Solarwinds MSP\Repair-PME\'"
                $DateTime = Get-Date -Format 'yyyy-MM-dd HH-mm-ss'
                $Install = Start-Process -FilePath "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe" -ArgumentList "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"C:\ProgramData\Solarwinds MSP\Repair-PME\Setup Log $DateTime.txt`"" -Wait -PassThru
                If ($Install.ExitCode -eq 0) {
                    Write-Host "OK: $($PMEDetails.Name) version $($PMEDetails.Version) successfully installed" -ForegroundColor Green
                } ElseIf ($Install.ExitCode -eq 5) {
                    Write-EventLog @WriteEventLogErrorParams -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode).`nScript: Repair-PME.ps1"
                    Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode)"
                } Else {
                    Write-EventLog @WriteEventLogErrorParams -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'.`nScript: Repair-PME.ps1"
                    Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'"
                }
            } Else {
                Write-EventLog @WriteEventLogErrorParams -Message "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting.`nScript: Repair-PME.ps1"
                Throw "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting"
            }
        }
    } Else {
        Write-Output "$($PMEDetails.FileName) does not exist, begin download and install phase"
        # Check for PME Archive Directory
        If (Test-Path "C:\ProgramData\SolarWinds MSP\PME\archives") {
            Write-Output "Directory 'C:\ProgramData\SolarWinds MSP\PME\archives' already exists, no need to create directory"
        } Else {
            Try {
                Write-Output "Directory 'C:\ProgramData\SolarWinds MSP\PME\archives' does not exist, creating directory"
                New-Item -ItemType Directory -Path "C:\ProgramData\SolarWinds MSP\PME\archives" -Force | Out-Null
            } Catch {
                Write-EventLog @WriteEventLogErrorParams -Message "Unable to create directory 'C:\ProgramData\SolarWinds MSP\PME\archives' required for download, aborting. Error: $($_.Exception.Message).`nScript: Repair-PME.ps1"
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
            Write-Output "Installing $($PMEDetails.FileName) - logs will be saved to 'C:\ProgramData\Solarwinds MSP\Repair-PME\'"
            $DateTime = Get-Date -Format 'yyyy-MM-dd HH-mm-ss'
            $StartProcessParams = @{
                FilePath     = "C:\ProgramData\SolarWinds MSP\PME\archives\PMESetup_$($PMEDetails.Version).exe"
                ArgumentList = "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"C:\ProgramData\Solarwinds MSP\Repair-PME\Setup Log $DateTime.txt`""
                Wait         = $true
                PassThru     = $true
            }
            $Install = Start-Process @StartProcessParams
            If ($Install.ExitCode -eq 0) {
                Write-Host "OK: $($PMEDetails.Name) version $($PMEDetails.Version) successfully installed" -ForegroundColor Green
            } ElseIf ($Install.ExitCode -eq 5) {
                Write-EventLog @WriteEventLogErrorParams -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode).`nScript: Repair-PME.ps1"
                Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed because access is denied, exit code $($Install.ExitCode)"
            } Else {
                Write-EventLog @WriteEventLogErrorParams -Message "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'.`nScript: Repair-PME.ps1"
                Throw "$($PMEDetails.Name) version $($PMEDetails.Version) was unable to be successfully installed, exit code $($Install.ExitCode) see 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-'"
            }
        } Else {
            Write-EventLog @WriteEventLogErrorParams -Message "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting.`nScript: Repair-PME.ps1"
            Throw "Hash of file downloaded ($($Download.SHA256Checksum)) does not equal hash ($($PMEDetails.SHA256Checksum)) from sis.n-able.com, aborting"
        }
    }
}

Function Confirm-PMEServices {
    If ($Install.ExitCode -eq 0) {
        Write-Host "Checking PME services post-installation..." -ForegroundColor Cyan
        $SolarWindsMSPCacheStatus = (get-service "SolarWinds.MSP.CacheService" -ErrorAction SilentlyContinue).Status
        $SolarWindsMSPPMEAgentStatus = (get-service "SolarWinds.MSP.PME.Agent.PmeService" -ErrorAction SilentlyContinue).Status
        $SolarWindsMSPRpcServerStatus = (get-service "SolarWinds.MSP.RpcServerService" -ErrorAction SilentlyContinue).status

        Write-Output "SolarWinds MSP Cache Service Status: $SolarWindsMSPCacheStatus"
        Write-Output "SolarWinds MSP PME Agent Status: $SolarWindsMSPPMEAgentStatus"
        Write-Output "SolarWinds MSP RPC Server Status: $SolarWindsMSPRpcServerStatus"
    
        If (($SolarWindsMSPPMEAgentStatus -eq 'Running') -and ($SolarWindsMSPCacheStatus -eq 'Running') -and ($SolarWindsMSPRpcServerStatus -eq 'Running')) {
            Write-Host "OK: All PME services are installed and running following installation" -Foregroundcolor Green
        }
        Else {
            Write-EventLog @WriteEventLogErrorParams -Message "One or more of the PME services are not installed or running, investigation required.`nScript: Repair-PME.ps1"
            Throw "One or more of the PME services are not installed or running, investigation required"
        }
    }
}    

Function Set-End {
    Write-EventLog @WriteEventLogInformationParams -Message "Repair-PME has finished.`nScript: Repair-PME.ps1"
}

. Confirm-Elevation
. Set-Start
. Get-OSVersion
. Get-OSArch
. Get-PSVersion
. Set-CryptoProtocol
. Get-RepairPMEUpdate
. Test-Connectivity
. Test-SWCertificate
. Get-NCAgentVersion
. Confirm-PMEInstalled
. Get-PMESetupDetails
. Confirm-PMERecentInstall
. Confirm-PMEUpdatePending
. Invoke-SolarwindsDiagnostics
. Stop-PMESetup
. Stop-PMEServices
. Clear-PME
. Get-PMEConfigMisconfigurations
#. Set-PMEConfig
. Install-PME
. Confirm-PMEServices
. Set-End
