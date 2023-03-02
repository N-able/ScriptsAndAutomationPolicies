<#
Version = "0.3.1.1 (09th December 2022)"
#>

#region functions
Function Get-TakeControlVersionURL {

    $SISXML='https://sis.n-able.com/GenericFiles.xml'
    
    # Set Correct XML Location
    if (test-path 'C:\Program Files\N-able Technologies\Windows Agent') {
        Write-Host "32-Bit N-Central Agent Detected" -ForegroundColor Cyan
    $LocalXML='C:\Program Files\N-able Technologies\Windows Agent\Temp\GenericFiles.xml'
    }
    else {
        Write-Host "64-Bit N-Central Agent Detected" -ForegroundColor Cyan
        $LocalXML='C:\Program Files (x86)\N-able Technologies\Windows Agent\Temp\GenericFiles.xml'
    }
            # Download XML File if it doesn't already exist
            if (!(test-path "$LocalXML") -or ($localxml.length -le '78')) {
                Write-Host "GenericFiles.xml was not found locally/is not valid. Re-Downloading it " -nonewline -ForegroundColor Yellow;
                try {
                    Write-Host "using WebClient" -ForegroundColor Yellow
                    [System.Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768 -bor 192 -bor 48; (New-Object System.Net.WebClient).DownloadFile($sisxml, $localxml)  
                }
                catch {
                    Write-Host "Problem transferring file with Webclient. Trying Bitstransfer instead..." -ForegroundColor Red
                    import-module bitstransfer
                    Start-BitsTransfer -Source $SISXML -Destination $LocalXML -ErrorAction SilentlyContinue     
                }
            }
            if (test-path "$localxml") {
                $mspa = get-content $LocalXML
                $location = "MSPA4NCentral-"
                # $array=($mspa -match $location).Split('"')
                $array0 = $mspa -match $location
                $array = ("$array0").split('"')
                $MSPInstallurl = $array[1]
                $MSPCVersion = $array[3]
                write-host "Latest Take Control Agent Version: "-nonewline; write-host "$MSPCVersion" -foregroundcolor green 
            }
            else {
                Write-Host "XML file could not be downloaded" -ForegroundColor Red
                $MSPInstallURL = "N/A - Unable to detect from XML File"
                $MSPCVersion = '0.0'
                $LatestAgentInstalled = "Unknown"
            }
            write-host "Take Control Installer URL: " -nonewline; write-host "$MSPInstallurl" -foregroundcolor green
            
}

Function Get-InstalledTakeControlVersion {

    $TakeControlUpdateLog = "C:\ProgramData\GetSupportService_N-Central\Logs\BASUpSRvc*.log"
    $searchpattern = "SupportExpress Version"
    if (test-path "C:\ProgramData\GetSupportService_N-Central\Logs") {
    $takeControlUpdateLogData = get-content $takecontrolUpdateLog

    $TakeControlVersionfromLog = (($takecontrolupdatelogData -match $searchpattern) -split (" "))[-1]
        if ($TakeControlVersionfromLog -ne $null) {
        Write-Host "Take Control Agent Version (Logged): " -nonewline; Write-Host "$TakeControlVersionfromLog" -ForegroundColor Green
        }
    }
    else {
        $TakeControlVersionfromLog = "N/A - Log Folder not found"  
        Write-Host "Take Control Agent Version (Logged): " -nonewline; Write-Host "$TakeControlVersionfromLog" -ForegroundColor Yellow  
    }

    if (test-path "C:\Program Files\BeAnywhere Support Express\GetSupportService_N-Central\BASupSrvc.exe") {
        Write-Host "32-Bit Take Control Agent Detected" - -foregroundcolor Cyan
        $TakeControlNCentralAgentLocation = "C:\Program Files\BeAnywhere Support Express\GetSupportService_N-Central\BASupSrvc.exe"
    }
    else {
        Write-Host "64-Bit Take Control Agent Detected" -ForegroundColor Cyan
        $TakeControlNCentralAgentLocation = "C:\Program Files (x86)\BeAnywhere Support Express\GetSupportService_N-Central\BASupSrvc.exe"
    }
    
    $TakeControlAgentVersion = (get-item $TakeControlNCentralAgentLocation -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
    Write-Host "Take Control Agent Version (Installed): " -nonewline; Write-Host "$TakeControlAgentVersion" -ForegroundColor Green

    if ($MSPCVersion -eq $null) {
        Write-Host "Unable to confirm latest version of Take Control Agent" -ForegroundColor Red
    }
    else {
        if ([version]('{0}.{1}.{2}' -f $TakeControlAgentVersion.split('.')) -ge [version]('{0}.{1}.{2}' -f $mspcversion.split('.'))) {
        #if ($TakeControlVersionFromLog -match "$MSPCVersion") {
            $LatestAgentInstalled = $True
            Write-Host "Latest Agent Installed: " -nonewline; Write-Host "$LatestAgentInstalled" -ForegroundColor Green
        }
        else {
            $LatestAgentInstalled = $False
            Write-Host "Latest Agent Installed: " -nonewline; Write-Host "$LatestAgentInstalled" -ForegroundColor Red 
        }
    }
}

Function Get-LastTakeControlAgentHeartbeat {
    # Set up the suffix for the TC Support Service Log File
    $tcLogDate = Get-Date -Format "yyyyMMdd"

    # (OK I could do this by modify dates, but these get gzipped daily)
    $tcLogPath = 'C:\ProgramData\GetSupportService_N-Central\Logs\BASupSrvc_' + $tcLogDate + '.log'

    # Note the time
    $currentTime = Get-Date -Format "HH:mm:ss"
    if ((test-path "C:\ProgramData\GetSupportService_N-Central\Logs") -and (test-path $tcLogPath)) {

    Write-Host "INFO: Creating backup of log file to read content from" -foregroundcolor Yellow
    copy-item $tclogpath "$tclogpath.backup"
    if ($? -eq $true) {
        $tclogpath = "$tclogpath.backup"
    }
    # Open the last Log File, find the last confirmed heartbeat
    $tcLogData = Get-Content -Path $tcLogPath

    $searchPattern = "[HandleDoCallNCentralHeartBeatSOAPAPI] - SOAP Heartbeat response received..."

    foreach($tcLogLine in $tcLogData)
    {
        if($tcLogLine.LastIndexOf($searchPattern) -ne -1)
        {
            # Copy the last line with a heartbeat reply.
            $lastBeat = $tcLogLine.Substring(11, 8)
        }
    }
        if ($lastbeat -ne $null){
            $timeDelta = New-TimeSpan -Start $lastBeat -End $currentTime

            $lastBeatSeconds = $timeDelta.TotalSeconds
        }
        else {
            Write-Host "Heartbeat was not detected in Log" -ForegroundColor Red
        }
    }
    else {
        Write-Host "No Agent Install Log was found, so no heartbeat was found" -ForegroundColor Yellow
        $timedelta = $null
    }
    if ($lastbeatseconds -lt '3600') {
    write-host "Last Take Control Agent Heartbeat: " -nonewline; Write-Host "$timedelta (Normal)" -ForegroundColor Green
    $LastHeartbeat = "Normal"
    }
    else {
    $timedelta = '-1'
    write-host "Last Take Control Agent Heartbeat: " -nonewline; Write-Host "$timedelta (Failed)" -ForegroundColor Red
    $LastHeartbeat = "Failed"
    }
}

Function Get-TakeControlAgentInstallerDate {

    $TakeControlAgentInstallLog = "C:\Windows\Temp\ServerAppletServiceInstaller-N-Central.install.log"
    if (test-path $TakeControlAgentInstallLog) {
        $TakeControlAgentInstallDate = (get-itemproperty $TakeControlAgentInstallLog).creationtime 
    }
    else {
        Write-Host "Installer Log not Found. Taking Creation Date of exe instead" -ForegroundColor Yellow
        $TakeControlServiceLocation = (get-wmiobject win32_service -filter "Name like 'BASupportExpressStandaloneService_N_Central'")
        if ($TakeControlServiceLocation -ne $null) {
            $TakeControlServiceLocation = (get-wmiobject win32_service -filter "Name like 'BASupportExpressStandaloneService_N_Central'").PathName.replace('"','')
            $TakeControlAgentInstallDate = (get-item $TakeControlServiceLocation).CreationTime
        }
    }
    if ($TakeControlAgentInstallDate -eq $null) {
        $TakeControlAgentInstallDate = 'N/A (Take Control exe was not found)'
    Write-Host "Last Agent Install Date: " -nonewline; Write-Host "N/A (Take Control exe was not found)" -ForegroundColor Red
    }
    else {
        Write-Host "Last Agent Install Date: " -nonewline; Write-Host "$TakeControlAgentInstallDate" -ForegroundColor Yellow    
    }
}

Function Get-TakeControlServiceDetails {
    $TakeControlServiceRunning = (Get-Service "BASupportExpressStandaloneService_N_Central" -erroraction silentlycontinue).Status
    if ($TakeControlServiceRunning -eq $null) {
        $TakeControlServiceRunning = 'N/A - Service is not detected'
    }
    if ($TakeControlServiceRunning -eq 'Running') {
    Write-Host "Take Control Service Status: " -nonewline; Write-Host "$TakeControlServiceRunning" -foregroundcolor Green
    }
    else {
        Write-Host "Take Control Service Status: " -nonewline; Write-Host "$TakeControlServiceRunning" -foregroundcolor Red
    }
}

Function Get-TakeControlAgentVersionini {
$TCAgentIni = get-content "$env:ProgramData\GetSupportService_N-Central\BASupSrvc.ini"
$TCAgentVersion = (($TCAgentIni -match "Version")[0].split('='))[1]
}
#endregion

. Get-TakeControlVersionURL
. Get-InstalledTakeControlVersion
. Get-LastTakeControlAgentHeartbeat
. Get-TakeControlAgentInstallerDate
. Get-TakeControlServiceDetails