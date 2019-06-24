<#
    ************************************************************************************************************
    Name: Get-AVDefenderErrorMessage
    Version: 0.2 (07th June 2019)
    Purpose:    Read through error message from AVD Error Manager XML
    Pre-Reqs:    Powershell 2
    0.2
    ************************************************************************************************************
#>
[System.IntPtr]::Size # an integer whose size is platform-specific.
If ([System.IntPtr]::Size -eq 4) {
    $OSVersion = "32-bit"
} Else {
    $OSVersion = "64-bit"
}

If ($OSVersion -eq "64-bit") {
    $global:AgentConfigFolder = "$env:systemdrive\Program Files (x86)\N-able Technologies\Windows Agent\config"
} Else {
    $global:AgentConfigFolder = "$env:systemdrive\Program Files\N-able Technologies\Windows Agent\config"
}

$LogLocation = "$agentconfigfolder\AVDefenderErrorManager.xml"
Write-Host "Log Location: " -Nonewline
Write-Host "$LogLocation" -ForegroundColor Yellow

[xml]$AVDDefenderError = Get-Content "$LogLocation"
$nodeexists = $AVDDefenderError.AVDefenderErrorManager.MessageHolders.DictionarySerializableOfStringArrayOfString.ArrayOfSerializableKeyValuePairOfStringArrayOfString.SerializableKeyValuePairOfStringArrayOfString

If ($nodeexists) {
    $ErrorMsg = $AVDDefenderError.AVDefenderErrorManager.MessageHolders.DictionarySerializableOfStringArrayOfString.ArrayOfSerializableKeyValuePairOfStringArrayOfString.SerializableKeyValuePairOfStringArrayOfString.Value.String[1]
} Else {
    $ErrorMsg = "None"
}

If ($errormsg -eq "None") {
    Write-Host "Error: " -nonewline; Write-Host "$ErrorMsg`n" -ForegroundColor Green
} Else {
    Write-Host "Error: " -nonewline; Write-Host "$ErrorMsg`n" -ForegroundColor Red
}
