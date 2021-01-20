Get-WmiObject -Namespace "root\CIMV2\Security\MicrosoftVolumeEncryption" -Class Win32_EncryptableVolume |
ForEach-Object {
  $ID = $_.DriveLetter ;
  Switch ($_.GetProtectionStatus().ProtectionStatus) {
    0 {$State = "PROTECTION OFF"}
    1 {$State = "PROTECTION ON"}
    2 {$State = "PROTECTION UNKNOWN"}
  }
  $RecoveryKey = (get-bitlockervolume -MountPoint $_.Driveletter).KeyProtector.RecoveryPassword
  If ($ID -eq "C:")
  {
    if($RecoveryKey -ne $null){
        $CKey = "$ID - $State - $RecoveryKey"
    }else{
        $CKey = "$ID - $State"
    }
  }
  If ($ID -eq "D:"){
    if($RecoveryKey -ne $null){
        $DKey = "D - $State - $RecoveryKey"
    }
    else{
        $DKey = "D - $State"
    }
   }
}
#output
if($Dkey.Length > 0){
    $result = $CKey + "|" + $Dkey
}else{
    $result = $CKey
}

$result