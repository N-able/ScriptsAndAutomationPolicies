#Requires -Version 2.0

<#    
    ************************************************************************************************************
    Name: Get-InternetSpeedTest-Ookla
    Version: 0.1.5 (15th April 2020)
    Purpose:    Use Ookla SpeedTest to measure available bandwidth on a device
    Pre-Reqs:    Powershell 2
    Changes:
    + Restructured Functions to be more generic
    + updated calculations, changed switches being used
    + Download and Upload results converted to Mbps rather than bytes and rounded to 3 decimal places
    + Latency and Jitter results are rounded to 2 decimal places
    + Add % to packet loss output
    ************************************************************************************************************
#>
$company = "Doherty Associates"
$url = "https://bintray.com/ookla/download/download_file?file_path=ookla-speedtest-1.0.0-win64.zip"
$filename = $url.split('=')[-1]
$location = "$env:programdata\$company\$filename"
$expectedziphash='1B62CF2B020E4BD6ECDF9D36DEF04F00CB538BD0'
$exelocation="$env:programdata\$company\speedtest\speedtest.exe"
$expectedexehash='F818B285CE5A66EFC3C897745B2B3526111759BF'
$companylogsfolder = "$env:systemdrive\programdata\$company\logs" 
$DownloadRequired = $null
$exehash = $null
$ziphash = $null

#region functions
function set-folders {

    if (!(test-path "$env:systemdrive\programdata\$company")) {
        Write-Host "Creating ProgramData\$company Directory"
        new-item -itemtype "directory" -path "$env:systemdrive\programdata\$company" | out-null
    }
    else {
        write-host "ProgramData\$company Already exists"
    }
   
    if (!(test-path "$CompanyLogsFolder")) {
        Write-Host "Creating ProgramData\$company\Logs Directory"
        new-item -itemtype "directory" -path "$companyLogsFolder" | out-null
    }
    else {
        write-host "ProgramData\$company\Logs Already exists"
    }

}

Function Test-Speedtest {
    if (!(test-path $exelocation -ErrorAction SilentlyContinue)) {
    Write-Host "Local Speedtest.exe does not exist" -ForegroundColor Yellow
    $DownloadRequired = "True"
    }
    else {
        Write-Host "Local Speedtest.exe exists " -nonewline -ForegroundColor Green
    
        $exehash = (CertUtil -hashfile $exelocation SHA1)[1].Replace(" ", "")
    
        if ("$expectedexehash" -match "$exehash") {
        Write-Host "and is Valid" -ForegroundColor Green
        $DownloadRequired = "False"
        }
        Else {
        Write-Host "but is corrupt" -ForegroundColor Red
        $DownloadRequired = "True"    
        }
    }
    # Write-Host "Download Required: $DownloadRequired"
}
   
Function Get-SpeedTest {
    if ($DownloadRequired -eq "True") {
                Import-Module BitsTransfer
                $start_time = Get-Date
                Write-Host "Downloading: $($url) to $($location)"
                
                [System.Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768 -bor 192 -bor 48; (New-Object System.Net.WebClient -erroraction stop).DownloadFile($url, $location)

                    If ($? -eq "True") {
                    Write-Host "Ookla SpeedTest Executable was succesfully downloaded in $((Get-Date).Subtract($start_time).Seconds) second(s)."
                    $ziphash = (CertUtil -hashfile $location SHA1)[1].Replace(" ", "")

                        if ($expectedziphash -eq $ziphash) {
                        Write-Host "Ookla SpeedTest.zip download was succesful. Extracting Zip File"
                        extract-speedtestzip
                        #Exit 
                        }
                    }
                    Else {
                        Write-Host "There was an problem downloading the Ookla SpeedTest.zip. Aborting Proceedings."
                        #Exit
                        }

    }
}

function extract-speedtestzip {
    Add-Type -Assembly "System.IO.Compression.Filesystem"
    [System.IO.Compression.ZipFile]::ExtractToDirectory("$location","$env:systemdrive\programdata\$company\speedtest")
}

Function execute-speedtest {
    if (test-path $exelocation) {
    $argument = "--accept-license --accept-gdpr --format=json-pretty"
    Write-Host "`nPerforming SpeedTest..." -ForegroundColor Green
    Write-Host "$exelocation $argument" -ForegroundColor Yellow
    start-process -filepath $exelocation -argumentlist $argument -RedirectStandardOutput $companylogsfolder\ooklaspeedtest.json -Wait
    $json = convertfrom-json -inputobject (get-content $companylogsfolder\ooklaspeedtest.json -raw)

    $timestamp = $json.timestamp
    $ISP = $json.isp
    $serverid = $json.server.id
    $servername = $json.server.name
    $serverhost = $json.server.host
    $latency = [math]::Round($json.ping.latency, 2)                  # Round to 2 decimal places
    $jitter = [math]::Round($json.ping.jitter, 2)                    # Round to 2 decimal places
    $download= [math]::Round($json.download.bandwidth / 125000, 3)   # Convert raw bytes value to Mbps and round to 3 decimal places
    $upload= [math]::Round($json.upload.bandwidth / 125000, 3)       # Convert raw bytes value to Mbps and round to 3 decimal places
    $packetloss = $json.packetloss
    $externalip = $json.interface.externalIp
    $internalip = $json.interface.internalIp
    $macaddress = $json.interface.macAddr
    $url = $json.result.url
    }
}
#endregion

. set-folders
. test-speedtest
. get-speedtest
. execute-speedtest

# Output Results
Write-Host "Time: " -nonewline; Write-Host "$timestamp" -ForegroundColor Green
Write-Host "Ping: " -nonewline; Write-Host "$latency ms" -ForegroundColor Green
Write-Host "Jitter: " -nonewline; Write-Host "$jitter ms" -ForegroundColor Green
Write-Host "Download: " -nonewline; Write-Host "$download Mbps" -ForegroundColor Green
Write-Host "Upload: " -nonewline; Write-Host "$upload Mbps" -ForegroundColor Green
Write-Host "Packet Loss: " -nonewline; Write-Host "$packetloss %" -ForegroundColor Green
Write-Host "Server: " -nonewline; Write-Host "$serverid - $servername - $serverhost" -ForegroundColor Green
Write-Host "ISP: " -nonewline; Write-Host "$ISP" -ForegroundColor Green
Write-Host "External IP: " -nonewline; Write-Host "$externalip" -ForegroundColor Green
Write-Host "Internal IP: " -nonewline; Write-Host "$internalip" -ForegroundColor Green
Write-Host "MAC Address: " -nonewline; Write-Host "$macaddress" -ForegroundColor Green
Write-Host "URL: " -nonewline; Write-Host "$url" -ForegroundColor Green
