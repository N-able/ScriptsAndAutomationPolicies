<#    
    ************************************************************************************************************
    Name: Get-AVDefenderErrorMessage
    Version: 0.1 (28th October 2018)
    Purpose:    Read through error message from AVD Error Manager XML
    Pre-Reqs:    Powershell 2
    0.1
    ************************************************************************************************************
#>


$OSVersion=(Get-WmiObject Win32_OperatingSystem).OSArchitecture
If ($OSVersion -eq "64-bit") {
$global:AgentConfigFolder = "$env:systemdrive\Program Files (x86)\N-able Technologies\Windows Agent\config"
}
Else {
	$global:AgentConfigFolder = "$env:systemdrive\Program Files\N-able Technologies\Windows Agent\config"
}

$LogLocation = "$agentconfigfolder\AVDefenderErrorManager.xml"
Write-Host "Log Location: " -nonewline; Write-Host "$LogLocation" -ForegroundColor Yellow

[xml]$AVDDefenderError = Get-Content "$LogLocation"
$nodeexists = $AVDDefenderError.AVDefenderErrorManager.MessageHolders.DictionarySerializableOfStringArrayOfString.ArrayOfSerializableKeyValuePairOfStringArrayOfString.SerializableKeyValuePairOfStringArrayOfString
if ($nodeexists) {
$ErrorMsg = $AVDDefenderError.AVDefenderErrorManager.MessageHolders.DictionarySerializableOfStringArrayOfString.ArrayOfSerializableKeyValuePairOfStringArrayOfString.SerializableKeyValuePairOfStringArrayOfString.Value.String[1]
}
else {
$ErrorMsg = "None"
}
if ($errormsg -eq "None") {
    Write-Host "Error: " -nonewline; Write-Host "$ErrorMsg`n" -ForegroundColor Green    
}
else {
Write-Host "Error: " -nonewline; Write-Host "$ErrorMsg`n" -ForegroundColor Red
}
