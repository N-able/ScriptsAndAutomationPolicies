<#
installnet45.ps1
Originally (installPS40.ps1) by: Evan Morrissey
Modified (installPS40.ps1) by: Tyler Jones
installnet45 Modification by: Stephen Testino
Last Updated: 2014-08-11

Tests if net45 is installed, updates if necessary
#>

param(
    [switch]$Reboot
)

# Is Net4.5 Installed?
function Test-Net45{

    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'){
        if (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction SilentlyContinue){
            return $True
        }
        return $False
    }
}


# Declare the function for retrieving the update file
function HTTP-Download{
  param(
  $url
  ,
  $fileName
  )
  $webClient = New-Object System.Net.WebClient
  $localPath = "$env:temp\$fileName"
  $remote = $url + $fileName
  $ErrorActionPreference = "Stop"

  $path_check = Test-Path $localPath
  if($path_check -eq $False){
      try
      {
        $webClient.downloadFile($remote,$localPath)
        return $localPath
      }
      catch [System.Management.Automation.MethodInvocationException]
      {
        "Error downloading file" | Write-Host
        "Download URL: " + $url | Write-Host
        "Full Error message: " | Write-Host
        $error[0] | fl * -f | Write-Host
        return $false
      }
  }
  else{
  Write-Host "Download already exists. Try running"
  return $localPath
  }
}

#Main Script
# Check if PC has .NET 4.5
if(!([bool]$(Test-Net45))){
    "`r`n.NET 4.5 not detected - Downloading and Installing."
    $baseurl = "http://download.microsoft.com/download/B/A/4/BA4A7E71-2906-4B2D-A0E1-80CF16844F5F/"
    $filename = "dotNetFx45_Full_setup.exe"
    $dlResult = HTTP-Download $baseURL $filename

    if ($dlResult -eq $false)
    {
      "`r`nDownload failed, please update Net 4.5 manually. Aborting script"
      exit 1
    }
    else
    {
      # Install Net 4.5
      & $dlResult /quiet /norestart
      Start-Sleep -Seconds 5
      $i = 0

      while (get-process | ? {$_.processName -eq ($filename -replace ".exe","")})
      {
        start-sleep -seconds 10
        $i++
        if ($i -gt 300)
        #if($i -gt 2)
        {
          "`r`nNET 4.5 has been trying to install for 30 minutes?!? Aborting!"
          #"`r`nNET 4.5 has been trying to install for 20 seconds?!? Aborting!"
          (get-process | ? {$_.processName -eq ($filename -replace ".exe","")}) | stop-process -force
          exit 1
        }
      }
      "`r`n.NET Framework 4.5 Install complete: " + (get-date)
      if($Reboot){
        Restart-Computer -Force
      }
    }
}
else{
    ".NET Framework 4.5 Already Installed"
}