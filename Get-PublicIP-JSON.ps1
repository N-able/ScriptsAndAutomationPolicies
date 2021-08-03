<#      
    ************************************************************************************************************
    Name: Get-PublicIP-JSON.ps1
    Version: 0.4.0 (03rd August 2021)
    Purpose: Get Public IP and GeoLocation Details
    Get Public IP Address via ifconfig.me, ipquail.com, ipinfo.io, ident.me, ipecho.net, locate.now.sh
    Get GeoLocation data via ipdata.io  
    Pre-Reqs: Powershell 2 and ipdata.io api key(s)

    0.2 + Updated to use IPStack API
    0.3 + Updated to only use API when ip change detected
    0.3.2 + Changed to use IPData.io API for added details
    0.3.3 + updated for better PS 2.0 compatibility
    0.3.4.1 + Added Multiple IP Address Lookups for better balancing
    0.3.4.2 + added provisioning to ignore ssl trust for self-signed etc
    0.3.4.3 + Improved IP Address Detection
    0.3.4.4 + Generalized Script
    0.3.4.5 + Fixed ASN Detection
    0.3.4.5.2 + Adding Local Count of API Usage to detect outliers, updating choice of public ip detection
	0.4 + New Public Release
#>
$Version = '0.4.0 (03rd August 2021)'
Write-Host "Get-PublicIP-JSON " -nonewline; Write-Host "$Version`n" -ForegroundColor Green
$Date = Get-Date

# Set Company Name to create a Registry Branch to house the IP Data
$Company='Doherty Associates'
$path="HKLM:\SOFTWARE\$Company"

<#
IPData.io API Keys. 
These are limited to 1500 lookups/day on the free tier which is the best I could find from service providers.
AFAIK; Crucialy, this does not exclude commercial usage.
Depending on the rate on public ip change you will see amongst your device base, you may need to step up to a paid tier which are also very affordable.
#>

$array = '<ipdata api key1>,<ipdata api key2>'

# Array of multiple IP lookup services to spread the load
#removed 'https://locate.now.sh/ip', 'http://ident.me' from the array
$iparray = $null
$iparray = ('https://ifconfig.me/ip','http://ipquail.com/ip','http://ipinfo.io/ip','https://ipecho.net/plain')

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

#region functions

function Test-RegistryValue {

    param (

    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]$Path,
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]$Value
    )

    try {
    Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
    return $true
    }

    catch {
    return $false
    }
}
function set-companyregistry {
        
if (!(Test-Path $path))
    {
    Write-Host "Creating Registry Branch for $path"
    New-Item $path | out-null
    }
    else {
        Write-Host "Registry Branch for $path already exists"
    }

if (!(Test-Path "$path\IP Address"))
    {
    Write-Host "Creating Registry Branch for $path\IP Address"
    New-Item "$path\IP Address" | out-null
    }
    else {
        Write-Host "Registry Branch for $path\IP Address already exists"
    }
}

function ConvertFrom-Json20([object] $item){ 
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer

    #The comma operator is the array construction operator in PowerShell
    return ,$ps_js.DeserializeObject($item)
}

Function Get-IP {
$global:OldIPAddress = (Get-ItemProperty -path "$path\IP Address").IPAddress

$iparray_rand = $iparray[(Get-Random -Maximum ([array]$iparray).count)]
$iplookup = $iparray_rand.trim() 
Write-Host "`nIP Lookup Service: " -nonewline; Write-Host "$iplookup" -ForegroundColor Green
Set-ItemProperty -path "$path\IP Address" -name IPLookupService -value "$iplookup"
$global:newipaddress=(New-Object Net.WebClient).DownloadString($iplookup)

Write-Host "Old IP Address: " -nonewline; Write-Host "$oldipaddress" -ForegroundColor Yellow
Write-Host "New IP Address: " -nonewline; Write-Host "$newipaddress" -ForegroundColor Yellow
}

Function Get-IPData {

$apikey_rand = $array[(Get-Random -Maximum ([array]$array).count)]
Write-Host "Using API Key:" -nonewline; Write-Host "$apikey_rand" -ForegroundColor Green
$StringURL="https://api.ipdata.co?api-key=$apikey_rand" 

$PSversion = $PSVersionTable.PSVersion.Major

if ($psversion -lt 2) {
    Write-Host "This device does not meent the minimum requirement of Powershell 2.0 so the script will no longer continue to run."
    exit 1
}
if ($psversion -eq 2) {
    Write-Host "Powershell Version 2 Detected" -ForegroundColor yellow
    $global:StringVal = ConvertFrom-Json20 (New-Object Net.WebClient).DownloadString($StringURL)
}

if ($psversion -ge 3) {
    Write-Host "Powershell Version $psversion Detected" -ForegroundColor yellow
    $global:StringVal = (New-Object Net.WebClient).DownloadString($StringURL) | convertfrom-json
}

$global:APICount =  $stringVal.count

    if ($stringVal.Ip) {
        $ExternalIP =  $stringVal.Ip
        Set-ItemProperty -path "$path\IP Address" -name IPAddress -value $externalip
    }
    else {
        $ExternalIP =  "N/A"
    }
    if ($stringVal.Country_Name) {
        $CountryName =  $stringVal.Country_Name
        Set-ItemProperty -path "$path\IP Address" -name CountryName -value $countryname
    }
    else {
        $CountryName =  "N/A"
    }
    if ($stringVal.Country_Code) {
        $CountryCode =  $stringVal.Country_Code
        Set-ItemProperty -path "$path\IP Address" -name CountryCode -value $countrycode
    }
    else {
        $CountryCode =  "N/A"
    }
    if ($stringVal.Region) {
        $RegionName =  $stringVal.Region
        Set-ItemProperty -path "$path\IP Address" -name RegionName -value $RegionName
    }
    else {
        $RegionName =  "N/A"
    }
    if ($stringVal.City) {
        $City =  $stringVal.City
        Set-ItemProperty -path "$path\IP Address" -name City -value $City
    }
    else {
        $City =  "N/A"
    }
    if ($stringVal.Latitude) {
        $Latitude =  [Double]$stringVal.Latitude
        Set-ItemProperty -path "$path\IP Address" -name Latitude -value $Latitude
    }
    else {
        $Latitude =  "0"
    }
    if ($stringVal.Longitude) {
        $Longitude =  [Double]$stringVal.Longitude
        Set-ItemProperty -path "$path\IP Address" -name Longitude -value $Longitude
    }
    else {
        $Longitude =  "0"
    }
    if ($stringVal.Region_Code) {
        $RegionCode =  $stringVal.Region_Code
        Set-ItemProperty -path "$path\IP Address" -name RegionCode -value $RegionCode
    }
    else {
        $RegionCode =  "N/A"
    }
    if ($stringVal.asn) {
    $ASN = $stringVal.asn.asn
    Set-ItemProperty -path "$path\IP Address" -name ASN -value $ASN
    }
    else {
        $ASN =  "N/A"
    }
    if ($stringVal.organisation) {
    $ISP = $stringVal.organisation
    write-host "ISP: $stringval.organisation"
    Set-ItemProperty -path "$path\IP Address" -name ISP -value $ISP
    }
    else {
        if ($stringVal.asn) {
        $ISP = $stringval.asn.name
        write-host "ISP: $stringval.asn.name"
        Set-ItemProperty -path "$path\IP Address" -name ISP -value $ISP
        }
        else {
            $ISP = "N/A"
        }
    }
}

#endregion

set-companyregistry


$IPLookupService = (test-registryvalue -Path "$path\IP Address" -Value 'IPLookupService')
if ($IPLookupService -eq $false) {
    Write-Host "`nSetting default value for IP Lookup Service" -ForegroundColor Cyan
    Set-ItemProperty -path "$path\IP Address" -name IPLookupService -value 'N/A'
}

Get-IP
Set-ItemProperty -path "$path\IP Address" -name LastRun -value $Date

$LastRun_Reg = (Get-ItemProperty -path "$path\IP Address").LastRun

Write-Host "Script Run: " -nonewline; Write-Host "$LastRun_Reg" -ForegroundColor Yellow

$LocalAPIUsage = (test-registryvalue -Path "$path\IP Address" -Value 'LocalAPIUsage')
if ($LocalAPIUsage -eq $false) {
    Write-Host "`nSetting default value for Local API Usage" -ForegroundColor Cyan
    Set-ItemProperty -path "$path\IP Address" -name LocalAPIUsage -value '0'
}

$CheckASN = (test-registryvalue -Path "$path\IP Address" -Value 'ASN')

if ($CheckASN -eq $False) {
    Write-Host "Running with new API to update ASN Value" -ForegroundColor Yellow
    $update = $True
}
else {
    $oldipaddress = $oldipaddress.trim()
    $newipaddress = $newipaddress.trim()
    if ($oldipaddress -eq $newipaddress){
    Write-Host "`nNo IP Change Detected. Reading GeoLocation Details from Registry..." -ForegroundColor Green
    $Update = $False
    
    }
    else {
        Write-Host "`nIP Change Detected. Updating GeoLocation Details..." -ForegroundColor Yellow
        $Update = $True
    }
}

if ($update -eq $true) {
    . Get-IPData
    $LocalAPIUsage = $LocalAPIUsage+1
    Set-ItemProperty -path "$path\IP Address" -name LocalAPIUsage -value "$LocalAPIUsage"
    Write-Host "Local API Usage Count: " -nonewline; Write-Host "$LocalAPIUsage" -ForegroundColor Green
    Write-Host "Global API Usage Count: " -nonewline; Write-Host "$APICount" -ForegroundColor Green
}

$ExternalIP = (Get-ItemProperty -path "$path\IP Address").IPAddress
$City = (Get-ItemProperty -path "$path\IP Address").City
$RegionName = (Get-ItemProperty -path "$path\IP Address").RegionName
$RegionCode = (Get-ItemProperty -path "$path\IP Address").RegionCode
$CountryName = (Get-ItemProperty -path "$path\IP Address").CountryName
$CountryCode = (Get-ItemProperty -path "$path\IP Address").CountryCode
$Latitude = (Get-ItemProperty -path "$path\IP Address").Latitude
$Longitude = (Get-ItemProperty -path "$path\IP Address").Longitude
$ASN = (Get-ItemProperty -path "$path\IP Address").ASN
$ISP = (Get-ItemProperty -path "$path\IP Address").ISP
$LatLong = "$Latitude/$Longitude"

Write-Host "`nExternal IP: " -nonewline; Write-Host "$ExternalIP" -ForegroundColor Green
Write-Host "City: " -nonewline; Write-Host "$City" -ForegroundColor Green
Write-Host "Region Name: " -nonewline; Write-Host "$RegionName" -ForegroundColor Green
Write-Host "Region Code: " -nonewline; Write-Host "$RegionCode" -ForegroundColor Green
Write-Host "Country Name: " -nonewline; Write-Host "$CountryName" -ForegroundColor Green
Write-Host "Country Code: " -nonewline; Write-Host "$CountryCode" -ForegroundColor Green
Write-Host "Lat/Long: " -nonewline; Write-Host "$LatLong" -ForegroundColor Green
Write-Host "ASN: " -nonewline; Write-Host "$ASN" -ForegroundColor Green
Write-Host "ISP: " -nonewline; Write-Host "$ISP`n" -ForegroundColor Green

$LocalAPIUsage = (Get-ItemProperty -path "$path\IP Address").LocalAPIUsage
Write-Host "Local API Usage: " -nonewline; Write-Host "$localAPIUsage" -ForegroundColor Green