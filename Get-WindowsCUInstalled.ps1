<#    
    ************************************************************************************************************
    Name: Get-WindowsCUInstalled
    Version: 0.1.1 (26th February 2020)
    Purpose:    Detect device build and check whether recentCU is installed
    Pre-Reqs:    Powershell 2
    0.1.1 + Added 24/02 Cumulative Update Builds
    ************************************************************************************************************

Obtained from https://pureinfotech.com/windows-10-version-release-history/

1909
February 11, 2020	18363.657	KB4532693
January 28, 2020	18363.628	KB4532695
January 14, 2020	18363.592	KB4528760

1903
February 11, 2020	18362.657	KB4532693
January 28, 2020	18362.628	KB4532695
January 14, 2020	18362.592	KB4528760

1809
February 11, 2020	17763.1039	KB4532691
January 24, 2020	17763.1012	KB4534321
January 14, 2020	17763.973	KB4534273

1803
February 11, 2020	17134.1304	KB4537762
January 24, 2020	17134.1276	KB4534308
January 14, 2020	17134.1246	KB4534293

1709
February 11, 2020	16299.1686	KB4537789
January 24, 2020	16299.1654	KB4534318
January 14, 2020	16299.1625	KB4534276

1703
February 11, 2020	15063.2284	KB4537765
January 14, 2020	15063.2254	KB4534296

1607
February 11, 2020	14393.3504	KB4537764
January 24, 2020	14393.3474	KB4534307
January 14, 2020	14393.3443	KB4534271

#>

$CU = $Null
$requiredbuild = $null

#region functions
Function Get-OSVersion{
    $winbuild = (Get-WmiObject -class Win32_OperatingSystem).Version
    # [string]$WinBuild=[System.Environment]::OSVersion.Version
    $osproduct = (Get-WmiObject -class Win32_OperatingSystem).Producttype
    $UBR = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR).UBR
    $OSBuildVersion = $winbuild + "." + $UBR
    
    
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
                        $osvalue = "Win7"
                       } #end if
                     Else {
                        $osvalue = "Server 2008 R2"
                       } #end else
                   } #end 7600
        "6.1.7601*"   { # SP1
                     If($osproduct -eq 1) {
                        $osvalue = "Win7 SP1"
                       } #end if
                     Else {
                        $osvalue = "Server 200 8R2"
                       } #end else
                   } #end 7601
       "6.2.9200*" {
                     If($osproduct -eq 1) {
                        $osvalue = "Win 8"
                       }
                     Else {
                        $osvalue = "Server 2012"
                       }
                   }
        "6.3.9600*" {
                     If($osproduct -eq 1) {
                        $osvalue = "Win 8.1"
                       }
                     Else {
                        $osvalue = "Server 2012 R2"
                       }
                   }
       "10.0.10130*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Win10"
                       }
                    Else {
                        $osvalue = "Server 2016"
                        }
                    }
       "10.0.10240*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Win10 (1507)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1507)"
                        }  
                    }                
       "10.0.10586*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Win10 (1511)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1511)"
                        }
                    }
       "10.0.14393*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Win10 AU (1607)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1607)"
                        }
                    }
       "10.0.15063*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Win10 CU (1703)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1703)"
                        }
                    }
       "10.0.16*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Win10 FCU (1709)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1709)"
                        }
                    }
       "10.0.171*" {
                     If($osproduct -eq 1) {
                       $osvalue = "Win10 April 2018 (1803)"
                       }
                    Else {
                        $osvalue = "Server 2016 (1803)"
                        }
                    }
         "10.0.177*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Win10 October 2018 (1809)"
                    }
                Else {
                    $osvalue = "Server 2019 (1809)"
                    }
                }      
         "10.0.18362*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Win10 May 2019  (1903)"
                    }
                Else {
                    $osvalue = "Server 2019 (1903)"
                    }
                }
        "10.0.18363*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Win10 November 2019 (1909)"
                    }
                Else {
                    $osvalue = "Server 2019 (1909)"
                    }
                }
         "10.0.190*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Win10 20H1 (20xx)"
                    }
                Else {
                    $osvalue = "Server 2019 (20xx)"
                    }
                }    
        "10.0.195*" {
                    If($osproduct -eq 1) {
                    $osvalue = "Win10 20H2 (20xx)"
                    }
                Else {
                    $osvalue = "Server 2019 (20xx)"
                    }
                }           		   		   
        DEFAULT {$osvalue = "Unknown ($osversion)" }
      } #end switch
    
write-Host "Windows Version: " -nonewline; Write-Host "$osvalue" -ForegroundColor Green
write-Host "Windows Build: " -nonewline; Write-Host "$osbuildversion `n" -ForegroundColor Green
      
    } 

Function Get-CUInstalled {

#1909
if ($osbuildversion -eq "10.0.18363.657") {
    $CU = "KB4532693 (February 11, 2020)"
}

if ($osbuildversion -eq "10.0.18363.628") {
    $CU = "KB4532695 (January 28, 2020)"
}

if ($osbuildversion -eq "10.0.18363.592") {
    $CU = "KB4528760 (January 14, 2020)"
}


#1903
if ($osbuildversion -eq "10.0.18362.657") {
    $CU = "KB4532693 (February 11, 2020)"
}

if ($osbuildversion -eq "10.0.18362.628") {
    $CU = "KB4532695 (January 28, 2020)"
}

if ($osbuildversion -eq "10.0.18362.592") {
    $CU = "KB4528760 (January 14, 2020)"
}


#1809
if ($osbuildversion -eq "10.0.17763.1075") {
    $CU = "KB4537818 (February 24, 2020)"
}

if ($osbuildversion -eq "10.0.17763.1039") {
    $CU = "KB4532691 (February 11, 2020)"
}

if ($osbuildversion -eq "10.0.17763.1012") {
    $CU = "KB4534321 (January 24, 2020)"
}

if ($osbuildversion -eq "10.0.17763.973") {
    $CU = "KB4534273 (January 14, 2020)"
}


#1803
if ($osbuildversion -eq "10.0.17134.1345") {
    $CU = "KB4537795 (February 24, 2020)"
}

if ($osbuildversion -eq "10.0.17134.1304") {
    $CU = "KB4537762 (February 11, 2020)"
}

if ($osbuildversion -eq "10.0.17134.1276") {
    $CU = "KB4534308 (January 24, 2020)"
}

if ($osbuildversion -eq "10.0.17134.1246") {
    $CU = "KB4534293 (January 14, 2020)"
}


#1709
if ($osbuildversion -eq "10.0.16299.1717") {
    $CU = "KB4537816 (February 24, 2020)"
}

if ($osbuildversion -eq "10.0.16299.1686") {
    $CU = "KB4537789 (February 11, 2020)"
}

if ($osbuildversion -eq "10.0.16299.1654") {
    $CU = "KB4534318 (January 24, 2020)"
}

if ($osbuildversion -eq "10.0.16299.1625") {
    $CU = "KB4534276 (January 14, 2020)"
}
  

#1703
if ($osbuildversion -eq "10.0.15063.2284") {
    $CU = "KB4537765 (February 11, 2020)"
}

if ($osbuildversion -eq "10.0.15063.2254") {
    $CU = "KB4534296 (January 14, 2020)"
}
  

#1607
if ($osbuildversion -eq "10.0.14393.3542") {
    $CU = "KB4537806 (February 24, 2020)"
}

if ($osbuildversion -eq "10.0.14393.3504") {
    $CU = "KB4537764 (February 11, 2020)"
}

if ($osbuildversion -eq "10.0.14393.3474") {
    $CU = "KB4534307 (January 24, 2020)"
}

if ($osbuildversion -eq "10.0.14393.3443") {
    $CU = "KB4534271 (January 14, 2020)"
}

# Windows 8.1/2012 R2
if ($osbuildversion -eq "6.3.9600.19629") {
    $CU = "KB4537821 (February 11, 2020)"
}

if ($osbuildversion -eq "6.3.9600.19599") {
    $CU = "KB4534297 (January 14, 2020)"
}
  
# Windows 7/2008 R2
if ($osbuildversion -eq "6.1.7601.24544") {
$CU = "KB4534310 (January 14, 2020)"
}

# $CU = "KB4537820 (February 11, 2020)"


if (($osvalue -match "Windows 7") -or ($osvalue -match "2008 R2")) {
    $RequiredBuild = "6.1.7601.24544"
    }

if (($osvalue -match "8.1") -or ($osvalue -match "2012 R2")) {
    $RequiredBuild = "6.3.9600.19599"
    }

if ($osvalue -match "1607") {
    $RequiredBuild = "10.0.14393.3443"
    }

if ($osvalue -match "1703") {
    $RequiredBuild = "10.0.15063.2254"
    }


if ($osvalue -match "1709") {
    $RequiredBuild = "10.0.16299.1625"
    }

if ($osvalue -match "1803") {
    $RequiredBuild = "10.0.17134.1246"
    }

if ($osvalue -match "1809") {
    $RequiredBuild = "10.0.17763.973"
    }

if ($osvalue -match "1903") {
    $RequiredBuild = "10.0.18362.592"
    }

if ($osvalue -match "1909") {
    $RequiredBuild = "10.0.18363.592"
    }

if ($osvalue -match "20H") {
    $RequiredBuild = "10.0.19564.1005"
    $CU = 'N/A'
    }


if (($cu -ne $null) -and ($cu -ne 'N/A')) {
Write-Host "CU Installed: " -nonewline; Write-Host "$CU" -ForegroundColor Green
$KB = $CU.split(' ')[0]
$InstalledOn = (get-hotfix -id $KB).installedon
}

if ([version]$osbuildversion -ge [version]$requiredbuild) {
Write-Host "OS Build Version meets Required CU`n" -ForegroundColor Green
$Required = $True
}
else {
Write-Host "OS Build Version does not meet Required CU`n" -ForegroundColor Red
$Required = $False
}

}
#endregion        

. Get-OSVersion

. Get-CUInstalled

$Verdict = "$Required - $CU - $installedon"
Write-Host "$Verdict"