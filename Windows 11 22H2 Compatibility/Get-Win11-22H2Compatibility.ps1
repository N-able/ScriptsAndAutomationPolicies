# Using information from https://www.ghacks.net/2022/06/23/a-look-in-the-registry-reveals-if-your-pc-is-compatible-with-windows-11-version-22h2/

$CompatReg = (get-itemproperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators\NI22H2").RedReason
$CompatReg = $CompatReg | Out-String

if ($null -eq $CompatReg) {
    $Compat = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | ERROR | Unable to Read Registry Value"
    exit 2
}
else {
    if ($CompatReg -match "None") {
        $Compat = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | INFO | $CompatReg - No hardware compatibility issues have been detected by Windows Telemetry Data"
        Write-Host $Compat -ForegroundColor Green
        Exit 0
    }
    else {
        $Compat = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | ERROR | $CompatReg"
        Write-Host $Compat -ForegroundColor Red
        Exit 1
    }
}