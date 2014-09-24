<#
installPS30.ps1
Originally by: Evan Morrissey
Modified by: Tyler Jones
Last Updated: 2014-08-05

Tests the version of PowerShell and attempts to update it to PowerShell 3.0
#>

# Find OS architecture
$psVer = $host.version
$os = get-wmiobject win32_operatingsystem
$osArch = "x64"

if ($os.osarchitecture -match "86")
{
  $osArch = "x86"
}

# Write date and existing PS details for logs
$date = get-date
"PowerShell 3.0 Install/Upgrade: " + $date | Write-Host
"`r`nInstalled PowerShell version: " + $psVer.toString() | Write-Host

# Check if powershell is already version 3 or higher
if ($psVer.major -ge 3)
{
  "`r`nPowerShell is already up to date, exiting" | Write-Host
  exit 0
}

# Check for versions of Windows not compatible with PowerShell 3.0
if (($os.version -lt 6.0) -or ($os.caption -match "vista") -or ($os.version -ge 6.2))
{
  "`r`nPowerShell 3.0 requires Windows 6.0 or greater, and is not compatible with Windows Vista. It is preloaded on Windows 8. No update required/possible. Exiting." | Write-Host
  exit 0
}


# Declare the function for retrieving the update file
function HTTP-Download
{
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

# Download the installer
"`r`nAttempting to download latest PowerShell installer from Microsoft..." | Write-Host

# Create the download URL based on the windows version/architecture
$baseURL = "http://download.microsoft.com/download/E/7/6/E76850B8-DA6E-4FF5-8CCE-A24FC513FD16/"
if ($os.version.subString(0,3) -eq 6.1)
{
  $psFile = "Windows" + $os.version.subString(0,3) + "-KB2506143-" + $osArch + ".msu"
}
else
{
  if ($os.version.subString(0,3) -eq 6.0)
  {
    $psFile = "Windows" + $os.version.subString(0,3) + "-KB2506146-" + $osArch + ".msu"
  }
}

if ($psFile -eq $null)
{
  "`r`nCould not generate a URL for download. Check Windows version to see if PowerShell 3.0 is compatible or required." | Write-Host
  exit 0
}

# Download function call
$dlResult = HTTP-Download $baseURL $psFile

if ($dlResult -eq $false)
{
  "`r`nDownload failed, please update PowerShell manually. Aborting script" | Write-Host
  exit 1
}
else
{
  # Install Windows Mangement Framework 3.0 (PowerShell 3.0)
  & $dlResult /quiet /norestart
  $i = 0
  while (get-process | ? {$_.processName -eq "wusa"})
  {
    start-sleep -seconds 10
    $i++
    if ($i -gt 120)
    {
      "`r`nPowerShell has been trying to install for twenty minutes?!? Aborting!" | Write-Host
      get-process | ? {$_.path -eq "wusa"} | stop-process -force
      exit 1
    }
  }
  "`r`nInstall complete. A reboot is required..." | Write-Host
  "`r`nPowerShell 3.0 Install complete: " + (get-date) | Write-Host
}