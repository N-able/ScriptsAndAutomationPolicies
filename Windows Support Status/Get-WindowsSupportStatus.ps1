<#
Get-WindowsSupportStatus 0.2.0 (21st July 2023)

.SYNOPSIS
    Takes the windows build details from the device and determines the active support status and security support status
.DESCRIPTION
    The script will take the OS Build and then attempt to look it up against endoflife.date to gather the current statuses re active support and ssecurity support.
    It will fail over to copies of those JSON files stored on your N-Central Server.
    Windows Insider builds get looked up against the MS Flight Hub Site
.NOTES
    For the failover to work, you will need to download the JSON files from endoflife.date and place those copies on your N-Central Server Repository.
    Then Update the links below/input parameters in the AMP accordingly.

    Thanks to Homotechsual as this was adapted from his base idea: https://homotechsual.dev/2022/12/22/NinjaOne-custom-fields-endless-possibilities#windows-os-support-status
    Additional thanks to Darren Chapman for some needed tweaks to the Server Lookup

.EXAMPLE

Device In Support:

OS: Microsoft Windows Server 2019 Standard (10.0.17763.4645)                    
Was able to retrieve Support Data from https://endoflife.date                   
Release Date: 2018-11-13                                                        
2023-07-21 17:51:24 | Windows Active Support: TRUE | End Date: 2024-01-09 | 171 Days Left | Microsoft Windows Server 2019 Standard | 10.0.17763.4645            
2023-07-21 17:51:24 | Windows Security Support: TRUE | End Date: 2029-01-09 | 1998 Days Left | Microsoft Windows Server 2019 Standard | 10.0.17763.4645         
Support URL: https://learn.microsoft.com/windows/release-health/windows-server-release-info  

Device Out of Support:

OS: Microsoft Windows 10 Pro (10.0.19044.3086)                                  
Was able to retrieve Support Data from https://endoflife.date                   
Release Date: 2021-11-16                                                        
2023-07-21 17:53:18 | Windows Active Support: FALSE | End Date: 2023-06-13 | 39 Days Over | Microsoft Windows 10 Pro | 10.0.19044.3086                          
2023-07-21 17:53:18 | Windows Security Support: FALSE | End Date: 2023-06-13 | 39 Days Over | Microsoft Windows 10 Pro | 10.0.19044.3086                        
Support URL: https://learn.microsoft.com/windows/release-health/release-information         

Insider Build:

OS: Microsoft Windows 11 Enterprise (10.0.25375.1)
Support information will not be looked up against https://endoflife.date as you appear to be running an insider build.
Checking Release date against Flight Hub.
Release Date: 2023-05-25
2023-07-21 17:50:15 | Windows Active Support: N/A - Insider Build | Microsoft Windows 11 Enterprise | 10.0.25375.1
2023-07-21 17:50:15 | Windows Security Support: N/A - Insider Build | Microsoft Windows 11 Enterprise | 10.0.25375.1
Support URL: https://learn.microsoft.com/en-us/windows-insider/flight-hub/

#>

$Version = "0.2.0 (21st July 2023)"

$Source = "https://endoflife.date"
$EndOfLifeUriWindows = 'https://endoflife.date/api/windows.json'
$EndOfLifeUriServer = 'https://endoflife.date/api/windowsserver.json'
$backupEndOfLifeUriWindows = "https://<NC SERVER URL>/download/repository/<GUID>/windows.json"
$backupEndOfLifeUriServer = "https://<NC SERVER URL>/download/repository/<GUID>/windowsserver.json"


Write-Host "Get-WindowsSupportStatus $Version" -ForegroundColor Green

#region functions

Function Get-InsiderBuildData {

    # Get the current build number
    $buildNumber = $windowsversion.Build
    
    # Define the URL for the Windows Insider Flight Hub website
    $url = "https://learn.microsoft.com/en-us/windows-insider/flight-hub/"
    
    # Send a request to the website and get the response
    $response = Invoke-WebRequest -Uri $url
    
    $buildDatestring = ($response.Content | Select-String -pattern "(?i)<a href=`".*?$buildNumber.*?`" data-linktype.*?>(.*?)<\/a>" | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Groups | Select-Object -ExpandProperty Value)[1]
    $buildDate = [datetime]::ParseExact($buildDateString, "M/d/yyyy", $null)
    
    # Format the release date in the desired format
    $releaseDate = $buildDate.ToString("yyyy-MM-dd")
}

Function Get-WindowsSupportStatus {

$WindowsVersion = [System.Environment]::OSVersion.Version
$OSVersion = ($WindowsVersion.Major, $WindowsVersion.Minor, $WindowsVersion.Build -Join '.')

# Old OS's may not have a UBR Value
$UBR = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR -ErrorAction SilentlyContinue).UBR
if ($null -eq $UBR) {
    $OSFullVersion = $OSVersion    
}
else {
    $OSFullVersion = ($OSVersion, $UBR -Join '.')
}

$OSCaption = (Get-WmiObject Win32_OperatingSystem).Caption
$OSDetails = "$OSCaption ($OSFullVersion)"
Write-Host "`nOS: $OSDetails" -ForegroundColor Cyan

    if (($OSCaption -match 'Insider') -or ([version]$OSFullVersion -ge "10.0.25000.0")) {
        Write-Host "Support information will not be looked up against https://endoflife.date as you appear to be running an insider build.`nChecking Release date against Flight Hub." -ForegroundColor Yellow
        . Get-InsiderBuildData
        $windowsActiveSupport = "N/A - Insider Build"
        $OSActiveSupportDays = "N/A - Insider Build"
        $windowsSecuritySupport = "N/A - Insider Build"
        $OSSecuritySupportDays = "N/A - Insider Build"
        $supporturl = $url

    }
    else {
        # Set SSL Support 
        try{
            $protocols = 3072, 768, 192, 48
            $validProtocols = $protocols | Where-Object { [Enum]::IsDefined([System.Net.SecurityProtocolType], $_) }
            [System.Net.ServicePointManager]::SecurityProtocol = $validProtocols
        } 
        catch {
            try {
                Write-Host "Enumerating the Security Protocols did not work. Trying an alternative method..." -ForegroundColor yellow
                $Tls12 = 0x00000C00
                $SecurityProtocolType = [Enum]::ToObject([System.Net.SecurityProtocolType], $Tls12)
                [System.Net.ServicePointManager]::SecurityProtocol = $SecurityProtocolType
            }
            catch {
                Write-Host "Unable to change the Security Protocols being utilized" -ForegroundColor yellow
            }

        }
        

        $EoLRequestParams = @{
            Method = 'GET'
        }
        #$ProductName = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name ProductName).ProductName
        $ProductName = $OSCaption                    
        if ($ProductName -match 'Home' -or $ProductName -match 'Pro') {
            $Edition = '(W)'
        } 
        else {
            $Edition = '(E)'
        }
        if ($ProductName -like '*Server*') {
            $EoLRequestParams.Uri = $EndOfLifeUriServer
            $IsServerOS = $True
        } 
        else {
            $EoLRequestParams.Uri = $EndOfLifeUriWindows
        }

        start-sleep -Milliseconds $(Get-random -Maximum 10000)

        if ([version]$psversiontable.psversion -lt '3.0') {
            $Source = "N-Central Server"
            Write-Host "Attempting PS 2.0 Compatible lookup using webclient against $source" -ForegroundColor Yellow
            if ($ProductName -like '*Server*') {
                $EoLRequestParamsUri = $backupEndOfLifeUriServer
            } 
            else {
                $EoLRequestParamsUri = $backupEndOfLifeUriWindows
            }
            $webClient = New-Object System.Net.WebClient
            $lifecycles = $webClient.DownloadString($EoLRequestParamsuri)
            
        }
        else {
            try {
                $LifeCycles = Invoke-RestMethod @EoLRequestParams
            }
            catch {
                try {
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ERROR: Unable to access $source." -nonewline -ForegroundColor Yellow;
                    $Source = "N-Central Server"
                    Write-Host "Falling back to $Source" -ForegroundColor Yellow
                    if ($ProductName -like '*Server*') {
                        $EoLRequestParams.Uri = $backupEndOfLifeUriServer
                    } else {
                        $EoLRequestParams.Uri = $backupEndOfLifeUriWindows
                    }
                    $LifeCycles = Invoke-RestMethod @EoLRequestParams
                }
                catch {
                    $Source = "N-Central Server"
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ERROR: Unable to use invoke-restmethod. Falling back to webclient request to $source" -ForegroundColor Yellow
 
                    if ($ProductName -like '*Server*') {
                        $EoLRequestParams.Uri = $backupEndOfLifeUriServer
                    } 
                    else {
                        $EoLRequestParams.Uri = $backupEndOfLifeUriWindows
                    }
                    $webClient = New-Object System.Net.WebClient
                    $lifecycles = $webClient.DownloadString($EoLRequestParams.uri)

                }
            }
        }

        if ($null -ne $lifecycles) {
            Write-Host "Was able to retrieve Support Data from $Source" -ForegroundColor Green
        }

        if ($isserveros -eq $true) {
            # Original Code
            # $LifeCycle = $LifeCycles | Where-Object { ($_.latest -eq $OSVersion -or $_.buildId -eq $OSVersion) -and (($_.cycle -match "$Edition") -or ($IsServerOS)) }
            #$LifeCycle = $LifeCycles | Where-Object { ($_.latest -eq $OSVersion -or $_.buildId -eq $OSVersion) -and (($_.cycle -match [regex]::Escape($Edition)) -or ($IsServerOS)) }

            #Added extra code to narrow down server differences between Perpetual and Subscription (e.g. Server 2019 and 1809)
            $LifeCycle = $LifeCycles | Where-Object { $_.latest -eq $OSVersion -and (($_.cycle -match [regex]::Escape($Edition)) -or ($IsServerOS -and ("$($ProductName -replace " ","-")" -like "*$($_.cycle)*"))) }

        }
        else {
            #$LifeCycle = $LifeCycles | Where-Object { $_.latest -eq $OSVersion -and (($_.cycle -like "*$Edition*") -or ($IsServerOS)) }
            #$LifeCycle = $LifeCycles | Where-Object { $_.latest -eq $OSVersion -and ($_.cycle -match "$Edition*") }
            $LifeCycle = $LifeCycles | Where-Object { $_.latest -eq $OSVersion -and $_.cycle -match [regex]::Escape($Edition) }

        }

            if ($LifeCycle) {
                $datenow = Get-Date
                $ReleaseDate = $lifecycle.releasedate
                $supporturl = $lifecycle.link
                $OSActiveSupportDate = $lifecycle.support
                $OSSecuritySupportDate = $lifecycle.eol
                $OSActiveSupport = ($LifeCycle.support -ge (Get-Date -Format 'yyyy-MM-dd'))
                $OSSecuritySupport = ($LifeCycle.eol -ge (Get-Date -Format 'yyyy-MM-dd'))
                
                
                if ($OSActiveSupport) {
                    $OSActiveSupportDays = [math]::round(([datetime]$OSActivesupportDate - $datenow).TotalDays,0)
                    $windowsActiveSupport = "TRUE | End Date: $($LifeCycle.support) | $OSActiveSupportDays Days Left"
                    
                } 
                else {
                    $OSActiveSupportDays =  [math]::round(($datenow - [datetime]$OSActivesupportDate).TotalDays,0)
                    $windowsActiveSupport = "FALSE | End Date: $($LifeCycle.support) | $OSActiveSupportDays Days Over"
                    
                }
                if ($OSSecuritySupport) {
                    $OSSecuritySupportDays = [math]::round(([datetime]$OSSecuritySupportDate - $datenow).TotalDays,0)
                    $windowsSecuritySupport = "TRUE | End Date: $($LifeCycle.eol) | $OSSecuritySupportDays Days Left"
                    
                } 
                else {
                    $OSSecuritySupportDays = [math]::round(($datenow - [datetime]$OSSecuritySupportDate).TotalDays,0)
                    $windowsSecuritySupport = "FALSE | End Date: $($LifeCycle.eol) | $OSSecuritySupportDays Days Over"
                    
                }

            } 
            else {
                if ($null -eq $lifecycles) {
                    Write-Host "Support information for $ProductName $OSVersion not found from $Source." -ForegroundColor Yellow
                }
                if (($null -ne $lifecycles) -and ($null -eq $lifecycle)) {
                    Write-Host "Support information for $ProductName $OSVersion could not be parsed from $source data. Please investigate further." -ForegroundColor Yellow
                }
                $windowsActiveSupport = '-1'
                $windowsSecuritySupport = '-1'
                $ReleaseDate = "N/A"
                $supporturl = "N/A"
                $OSActiveSupportDays = "-1"
                $OSSecuritySupportDays = "-1"
            }
    }

Write-Host "Release Date: $ReleaseDate" -foregroundcolor cyan
Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Windows Active Support: " -foregroundcolor Cyan -nonewline; if ($windowsActiveSupport -match 'TRUE') {Write-Host "$windowsActiveSupport" -ForegroundColor Green -nonewline}else {Write-Host "$windowsActiveSupport" -ForegroundColor Red -nonewline}; Write-Host " | $OSCaption | $OSFullVersion" -ForegroundColor Cyan
Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Windows Security Support: " -foregroundcolor Cyan -nonewline; if ($windowsSecuritySupport -match 'TRUE') {Write-Host "$windowsSecuritySupport" -ForegroundColor Green -nonewline}else {Write-Host "$windowsSecuritySupport" -ForegroundColor Red -nonewline}; Write-Host " | $OSCaption | $OSFullVersion" -ForegroundColor Cyan
Write-Host "Support URL: $supporturl" -ForegroundColor Cyan
}

#endregion

. Get-WindowsSupportStatus