<#    
    ************************************************************************************************************
    Name: Install-WindowsUpdates-WUSA
    Version: 0.1.5.6 (24th Jan 2022)
    Purpose:    Install Windows Updates via WUSA
    Pre-Reqs:    Powershell 2, URL for downloads
    0.1 + Initial Verson
    0.1.5 Updated OS Build support
    0.1.5.1 - updated for latest 2020-01 Cumulative Updates
    0.1.5.2 - Updated for 2012 R2, latest 2020-02 Cumulative Updates
    0.1.5.3 - Refresher to October 2020 CU, adn adding in provisional support for 20H2
    0.1.5.4 + Updated for new 21H1 Build
    0.1.5.5 + Updated for 21H2, Windows 11 Builds
    0.1.5.6 + Updated KB's to be installed ot cover latest 2022-01 CU 
    ************************************************************************************************************
#>

$Version = "0.1.5.6 (24th Jan 2022)"
Write-Host "Install-WindowsUpdates " -nonewline; Write-Host "$Version" -ForegroundColor Green
$winbuild = $null
$osvalue = $null

#region functions
Function Get-OSVersion {
    $winbuild = (Get-WmiObject -class Win32_OperatingSystem).Version
    # [string]$WinBuild=[System.Environment]::OSVersion.Version
    $osproduct = (Get-WmiObject -class Win32_OperatingSystem).Producttype
    # Work Station (1)
    # Domain Controller (2)
    # Server (3)
    



    Switch -wildcard ($WinBuild) {
        "5.1.2600*" { $osvalue = "XP" }
        "5.1.3790*" { $osvalue = "2003" }
        "6.0.6001*" {
                     If($osproduct -eq 1) {
                        $osvalue = "Vista"
                       } #end if
                     Else {
                        $osvalue = "Server 2008"
                       } #end else
                   } #end 6001
         "6.0.6002*" {
                     If($osproduct -eq 1) {
                        $osvalue = "Vista SP2"
                       } #end if
                     Else {
                        $osvalue = "Server 2008 SP2"
                       } #end else
                   } #end 6002
        "6.1.7600*" {
                     If($osproduct -eq 1) {
                        $osvalue = "Windows 7"
                       } #end if
                     Else {
                        $osvalue = "Server 2008 R2"
                       } #end else
                   } #end 7600
        "6.1.7601*"   { # SP1
                     If($osproduct -eq 1) {
                        $osvalue = "Windows 7 SP1"
                       } #end if
                     Else {
                        $osvalue = "Server 2008 R2"
                       } #end else
                   } #end 7601
       "6.2.9200*" {
                     If($osproduct -eq 1) {
                        $osvalue = "Windows 8"
                       }
                     Else {
                        $osvalue = "Server 2012"
                       }
                   }
        "6.3.9600*" {
                     If($osproduct -eq 1) {
                        $osvalue = "Windows 8.1"
                       }
                     Else {
                        $osvalue = "Server 2012 R2"
                       }
                   }
       "10.0.10130*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Windows 10 (RTM)"
                       }
                    Else {
                        $osvalue = "Server 2016"
                        }
                    }
       "10.0.10240*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Windows 10 (1507)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1507)"
                        }  
                    }                
       "10.0.10586*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Windows 10 (1511)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1511)"
                        }
                    }
       "10.0.14393*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Windows 10 AU (1607)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1607)"
                        }
                    }
       "10.0.15063*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Windows 10 CU (1703)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1703)"
                        }
                    }
       "10.0.16*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Windows 10 FCU (1709)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1709)"
                        }
                    }
       "10.0.171*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Windows 10 April 2018 (1803)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1803)"
                        }
                    }
         "10.0.177*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Windows 10 October 2018 (1809)"
                    }
                Else {
                    $osvalue = "Server 2019 (1809)"
                    }
                }      
         "10.0.18362*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Windows 10 May 2019 (1903)"
                    }
                Else {
                    $osvalue = "Server 2019 (1903)"
                    }
                }
        "10.0.18363*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Windows 10 November 2019 (1909)"
                    }
                Else {
                    $osvalue = "Server 2019 (1909)"
                    }
                }
         "10.0.19041*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Windows 10 May 2020 (2004)"
                    }
                Else {
                    $osvalue = "Server 2019 (2004)"
                    }
                }    
        "10.0.19042*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Windows 10 October 2020 (20H2)"
                    }
                Else {
                    $osvalue = "Server 2019 (20H2)"
                    }
                }
        "10.0.19043*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Windows 10 May 2021 (21H1)"
                    }
                Else {
                    $osvalue = "Server 2019 (21H1)"
                    }
        }
        "10.0.19044*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Windows 10 Oct 2021 (21H2)"
                    }
                Else {
                    $osvalue = "Server 2019 (21H2)"
                    }
        }
        "10.0.220*" {
            If($osproduct -eq 1) {
            $osvalue = "Windows 11"
            }
            Else {
            $osvalue = "Server 2019"
            }
        }    
    
        "10.0.225*" {
            If($osproduct -eq 1) {
            $osvalue = "Windows 11 Insider (21H2)"
            }
            Else {
            $osvalue = "Server 2019 Insider (21H2)"
            }
                }           		   		   
        DEFAULT {$osvalue = "Unknown ($osversion)" }
      } #end switch

      Write-Host "Windows Build: " -nonewline; Write-Host "$WinBuild" -ForegroundColor Green
      write-Host "WinBuild: " -nonewline; Write-Host "$osvalue" -ForegroundColor Green

    . Get-OSArch
}

Function Get-OSArch {
    
    $OSname = (Get-WmiObject Win32_OperatingSystem).Caption
    $OSBfull = (Get-WmiObject Win32_OperatingSystem).Version
    if ($OSBfull -eq $null) { 
        $OSBfull = (Get-CimInstance Win32_OperatingSystem).Version 
    }
    $OSB = $OSBfull.Split(".")
    if (($OSname -like "*2003*") -or ($OSname -like "*XP*")) {
        if ($OSname -like "*x64*"){ 
            $arch = "64-bit" 
        }
        else{ 
            $arch = "32-bit" 
        }
    }
    else {
        $64bitOS = [System.Environment]::Is64BitOperatingSystem
        if ($64bitOS -eq $true){
            $arch = "64-bit"
        }
        else {
            $arch = "32-bit"
        }
        # $arch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
    }
    Write-Host "OS Architecture: " -nonewline; Write-Host "$arch`n" -ForegroundColor Green
}

Function Test-KBInstalled {
    Write-Host "$Title" -ForegroundColor Yellow
    $KB = $URL.split('-')[1]
    if (Get-HotFix $KB -ErrorAction SilentlyContinue) {
        Write-Host "$KB is already installed" -ForegroundColor Green
        $InstallRequired = $False
    }
    else {
        $InstallRequired = $True
    }
}

Function Download-KB {
    $filename = $URL.split('/')[-1]
    $KB = $URL.split('-')[1]
    $filepath = "$env:windir\temp\$filename"
    if (test-path $filepath) {
        Write-Host "$KB has already been downloaded"
    }
    else {
        Write-Host "$KB not found. Proceeding to Download..." -ForegroundColor Yellow
        Try {
            [System.Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768 -bor 192 -bor 48; (New-Object System.Net.WebClient -erroraction stop).DownloadFile($url, $filepath)
            if ($? -eq 'True') { Write-Host "$KB Download was successful" -foregroundcolor Green }
        }
        Catch {
            Write-Warning "There was an error with the download"
            Write-Host "Attempting Download via Bits-Transfer"
            [System.Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768 -bor 192 -bor 48; start-bitstransfer $url $filepath
            if ($? -eq 'True') { Write-Host "$KB Download was successful" -foregroundcolor Green }
        }

    }
}

Function Install-KB {
    . Test-KBInstalled
    if ($installrequired -eq $true) {
        . Download-KB
        Write-Host "Installing $KB - `n$env:windir\system32\wusa.exe -ArgumentList "$FilePath /quiet /norestart""
        $exitcode = (Start-Process -FilePath wusa.exe -ArgumentList "$FilePath /quiet /norestart" -Wait -passthru).Exitcode
        Write-Host "WUSA.exe exited with ErrorCode ($exitcode)"
        start-sleep 5
        Get-HotFix -id $KB -ErrorAction SilentlyContinue
        If ($? -eq "True") {
            Write-Host "$KB was successfully installed."
        }
        Else {
            Write-Host "There was a problem installing $KB." -ForegroundColor Red
        }
        if ($exitcode -eq '3010') {
            Write-Host "Reboot Required" -ForegroundColor Red
        }
    }
}
#endregion

. Get-OSVersion


# Install Server 2008 Updates
if (($Winbuild -like "6.0.600*") -and ($arch -like "64*")) {
    $title = "2020-10 Servicing Stack Update for Windows Server 2008 for x64-based Systems (KB4580971)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/10/windows6.0-kb4580971-x64_619424830431f3fba3c6d086b5dbe3e1ddf42f1f.msu"
    . Install-KB
    
    $title = "2020-10 Update for Windows Server 2008 for x64-based Systems (KB4578623)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/uprl/2020/10/windows6.0-kb4578623-x64_2c4ea22e26ace0c6d966196e581d15f06fa45c8d.msu"
    . Install-KB
}  

# Install Server 2008 R2 Updates
if (($Winbuild -like "6.1.7601*") -and ($arch -like "64*")) {
    $title = "2020-10 Servicing Stack Update for Windows Server 2008 R2 for x64-based Systems (KB4580970)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/10/windows6.1-kb4580970-x64_2616e0b0014f1b7db14ce7a8b0603dfdd0a9bf50.msu"
    . Install-KB
    
    $title = "2020-10 Update for Windows Server 2008 R2 for x64-based Systems (KB4578623)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/uprl/2020/10/windows6.1-kb4578623-x64_dcbc342c60cc1c6c4ca8559430008a8191e64455.msu"
    . Install-KB
}  

# Install Windows 7/ Server 2012 Updates
if (($Winbuild -like "6.1.7601*") -and ($arch -like "64*")) {
    $title = "2021-10 Servicing Stack Update for Windows 7 for x64-based Systems (KB5006749)"
    $url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/10/windows6.1-kb5006749-x64_4934673a816c4c0e5d6c380e0a61428da8aab4ac.msu"
    . Install-KB
    
    $title = "2021-10 Cumulative Security Update for Internet Explorer 11 for Windows 7 for x64-based systems (KB5006671)"
    $url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/09/windows6.1-kb5006671-x64_bdbf59b72f64161ebb89c9b7646e84e002ba033d.msu"
    . Install-KB
}  

# Install Windows 8/ Server 2012 Updates
if (($Winbuild -like "6.2.9200*") -and ($arch -like "64*")) {
    $title = "2021-04 Servicing Stack Update for Windows Server 2012 for x64-based Systems (KB5001401)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/04/windows8-rt-kb5001401-x64_1027ae2c9888c2dfe0caadeafc506b3012789c56.msu"
    . Install-KB
    
    $title = "2022-01 Security Monthly Quality Rollup for Windows Server 2012 for x64-based Systems (KB5009586)"
    $url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2022/01/windows8-rt-kb5009586-x64_fbaa8020aaaddbe28eceef67e03fc27f79294853.msu"
    . Install-KB
}

# Install Windows 8.1 Updates
if (($Winbuild -like "6.2.9600*") -and ($arch -like "64*")) {
    $title = "2021-08 Cumulative Security Update for Internet Explorer 11 for Windows 8.1 for x64-based systems (KB5005036)"
    $url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows8.1-kb5005036-x64_de4d5d3f6a2bc29954a6d578fa0fcdea1178176e.msu"
    . Install-KB
}  

# Install Server 2012 R2 Updates
if (($Winbuild -like "6.3.9600*") -and ($arch -like "64*")) {
$title = "2021-04 Servicing Stack Update for Windows Server 2012 R2 for x64-based Systems (KB5001403)"
$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/04/windows8.1-kb5001403-x64_7f15c4b281f38d43475abb785a32dbaf0355bad5.msu"
. Install-KB

$title = "2022-01 Security Monthly Quality Rollup for Windows Server 2012 R2 for x64-based Systems (KB5009624)"
$url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2022/01/windows8.1-kb5009624-x64_ae9f21e6bcae6274ea54ed380ab0a961aa7d6377.msu"
. Install-KB
}

# Install Windows 10/Server 2016 (1607) Updates
if (($Winbuild -like "10.0.14393*") -and ($arch -like "64*")) {
$title = "2021-09 Servicing Stack Update for Windows 10 Version 1607 for x64-based Systems (KB5005698)"
$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/09/windows10.0-kb5005698-x64_ff882b0a9dccc0c3f52673ba3ecf4a2a3b2386ca.msu"
. Install-KB

$title = "2022-01 Cumulative Update for Windows 10 Version 1607 for x64-based Systems (KB5010790)"
$url = "http://download.windowsupdate.com/c/msdownload/update/software/updt/2022/01/windows10.0-kb5010790-x64_df441589ce8556dc94d051df30be306266b25828.msu"
. Install-KB
}

# Install Windows 10/Server 2016 (1703) Updates
elseif (($Winbuild -like "10.0.15063*") -and ($arch -like "64*")) {
$title = "2020-07 Servicing Stack Update for Windows 10 Version 1703 for x86-based Systems (KB4565551)"
$url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2020/07/windows10.0-kb4565551-x86_4f5689e10f0f653b3b1a418b304563d68774bf8d.msu"
. Install-KB    

$title = "2021-03 Cumulative Update for Windows 10 Version 1703 for x64-based Systems (KB5000812)"
$url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/03/windows10.0-kb5000812-x64_40fa4a2bc24183c18efa21dd13056c6711a988d4.msu"
. Install-KB
}

# Install Windows 10/Server 2016 (1709) Updates
elseif (($Winbuild -like "10.0.16*") -and ($arch -like "64*")){
$title = "2020-07 Servicing Stack Update for Windows 10 Version 1709 for x64-based Systems (KB4565553)"
$url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2020/07/windows10.0-kb4565553-x64_37666386a4858aed971874a72cb8c07155c26a87.msu"
. Install-KB

$title = "2020-10 Cumulative Update for Windows 10 Version 1709 for x64-based Systems (KB4580328)"
$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/10/windows10.0-kb4580328-x64_b0565c3b26e77c3ce230a19121e30dccfb756235.msu"
. Install-KB    
}

# Install Windows 10/Server 2016 (1803) Updates
elseif ($Winbuild -like "10.0.171*") {
    if ($arch -like "64*") {
    $title = "2021-05 Servicing Stack Update for Windows 10 Version 1803 for x64-based Systems (KB5003364)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/05/windows10.0-kb5003364-x64_2d27672753b341e3f25305640320fbaedf2ee30e.msu"
    . Install-KB

    $title = "2021-05 Cumulative Update for Windows 10 Version 1803 for x64-based Systems (KB5003174)"
    $url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/05/windows10.0-kb5003174-x64_54fb6b8349dbaf6d5a9fff96c1cfe5f3663bf660.msu"
    . Install-KB   
    }
    else {
    $title = "2021-05 Servicing Stack Update for Windows 10 Version 1803 for x86-based Systems (KB5003364)"
        $url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/05/windows10.0-kb5003364-x86_ff3da229ef768edcb6e1eee725c3d98a82b10700.msu"
    . Install-KB        
    
    $title = "2021-05 Cumulative Update for Windows 10 Version 1803 for x86-based Systems (KB5003174)"
    $url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/05/windows10.0-kb5003174-x86_bdb2feb1b6446a46c2d425f12cb193ea64356f35.msu"
    . Install-KB
    }
}

# Install Windows 10/Server 2019 (1809) Updates
elseif (($Winbuild -like "10.0.177*") -and ($arch -like "64*")) {
#$title = "2020-10 Servicing Stack Update for Windows 10 Version 1809 for x64-based Systems (KB4577667)"
#$url = "http://download.windowsupdate.com/c/msdownload/update/software/secu/2020/10/windows10.0-kb4577667-x64_a5abb78aa80bfc785f5a3e5aeddcecfeda59bb2c.msu"
#. Install-KB   

$title = "2022-01 Cumulative Update for Windows 10 Version 1809 for x64-based Systems (KB5010791)"
$url = "http://download.windowsupdate.com/d/msdownload/update/software/updt/2022/01/windows10.0-kb5010791-x64_466ad2172d2cc77b2125420c26b7f9ac00e197f4.msu" 
. Install-KB   
}

# Install Windows 10/Server 2019 (1903) Updates
elseif (($Winbuild -like "10.0.18362*") -and ($arch -like "64*")) {
#$title = "2020-10 Servicing Stack Update for Windows 10 Version 1903 for x64-based Systems (KB4577670)"
#$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/10/windows10.0-kb4577670-x64_31cb9f2d35c2f2e9f4d9ba67caab858822d3c12c.msu"
#. Install-KB  

$title = "2022-01 Cumulative Update for Windows 10 Version 1909 for x64-based Systems (KB5010792)"
$url = "http://download.windowsupdate.com/c/msdownload/update/software/updt/2022/01/windows10.0-kb5010792-x64_0137657193d0d56be086f4bf134098c3d5dc8efa.msu"
. Install-KB  
    
}

# Install Windows 10/Server 2019 (1909) Updates
elseif (($Winbuild -like "10.0.18363*") -and ($arch -like "64*")) {
#$title = "2020-10 Servicing Stack Update for Windows 10 Version 1909 for x64-based Systems (KB4577670)"
#$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/10/windows10.0-kb4577670-x64_31cb9f2d35c2f2e9f4d9ba67caab858822d3c12c.msu"
#. Install-KB

$title = "2022-01 Cumulative Update for Windows 10 Version 1909 for x64-based Systems (KB5010792)"
$url = "http://download.windowsupdate.com/c/msdownload/update/software/updt/2022/01/windows10.0-kb5010792-x64_0137657193d0d56be086f4bf134098c3d5dc8efa.msu"
. Install-KB
}

# Install Windows 10/Server 2019 (2004) Updates
elseif (($Winbuild -like "10.0.19041*") -and ($arch -like "64*")) {
    #$title = "2020-02 Servicing Stack Update for Windows 10 Version 1909 for x64-based Systems (KB4538674)"
    #$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/02/windows10.0-kb4538674-x64_d181b6a4a516b43bc58cbc15347b83219dccee86.msu"
    #. Install-KB
    
    $title = "2022-01 Cumulative Update for Windows 10 Version 1909 for x64-based Systems (KB5010792)"
    $url = "http://download.windowsupdate.com/c/msdownload/update/software/updt/2022/01/windows10.0-kb5010792-x64_0137657193d0d56be086f4bf134098c3d5dc8efa.msu"
    . Install-KB
    }

# Install Windows 10/Server 2019 (20H2) Updates
elseif (($Winbuild -like "10.0.19042*") -and ($arch -like "64*")) {
    #$title = "2020-02 Servicing Stack Update for Windows 10 Version 1909 for x64-based Systems (KB4538674)"
    #$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/02/windows10.0-kb4538674-x64_d181b6a4a516b43bc58cbc15347b83219dccee86.msu"
    #. Install-KB
    
    $title = "2022-01 Cumulative Update for Windows 10 Version 20H2 for x64-based Systems (KB5010793)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/updt/2022/01/windows10.0-kb5010793-x64_3bae2e811e2712bd1678a1b8d448b71a8e8c6292.msu"
    . Install-KB
    }


# Install Windows 10/Server 2019 (21H1) Updates
elseif (($Winbuild -like "10.0.19043*") -and ($arch -like "64*")) {
    #$title = "2020-02 Servicing Stack Update for Windows 10 Version 1909 for x64-based Systems (KB4538674)"
    #$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/02/windows10.0-kb4538674-x64_d181b6a4a516b43bc58cbc15347b83219dccee86.msu"
    #. Install-KB
    
    $title = "2022-01 Cumulative Update for Windows 10 Version 21H1 for x64-based Systems (KB5010793)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/updt/2022/01/windows10.0-kb5010793-x64_3bae2e811e2712bd1678a1b8d448b71a8e8c6292.msu"
    . Install-KB
    }    

# Install Windows 10/Server 2019 (21H2) Updates
elseif (($Winbuild -like "10.0.19044*") -and ($arch -like "64*")) {
    #$title = "2020-02 Servicing Stack Update for Windows 10 Version 1909 for x64-based Systems (KB4538674)"
    #$url = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/02/windows10.0-kb4538674-x64_d181b6a4a516b43bc58cbc15347b83219dccee86.msu"
    #. Install-KB
    
    $title = "2022-01 Cumulative Update for Windows 10 Version 21H2 for x64-based Systems (KB5010793)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/updt/2022/01/windows10.0-kb5010793-x64_3bae2e811e2712bd1678a1b8d448b71a8e8c6292.msu"
    . Install-KB
    }

# Install Server 21H2 Updates
elseif (($Winbuild -like "10.0.19044*") -and ($osproduct -ne 1) -and ($arch -like "64*")) {
    
    $title = "2022-01 Cumulative Update for Microsoft server operating system version 21H2 for x64-based Systems (KB5010796)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/updt/2022/01/windows10.0-kb5010796-x64_ff1867cc8ed482b3115c6573e5743794c1dede8a.msu"
    . Install-KB
    }

# Install Windows 11 Updates
elseif (($Winbuild -like "10.0.220*") -and ($arch -like "64*")) {
    
    $title = "2022-01 Cumulative Update for Windows 11 for x64-based Systems (KB5010795)"
    $url = "http://download.windowsupdate.com/d/msdownload/update/software/updt/2022/01/windows10.0-kb5010795-x64_7fd6ce84756ac03585cc012568979eb08cc6d583.msu"
    . Install-KB
    }


else {
    Write-Host "There are no approved KB's that we require to manually install on this OS`n" -ForegroundColor Green
}