$Version = "0.2.4 (18th August 2023)"
Write-Host "Cleanup-OldNCAgentVersions $Version" -ForegroundColor Green

Function Cleanup-Agents {
    # Loop through the subkeys again and delete the ones with older versions
    foreach ($subKey in $subKeys) {
        $displayName = (Get-ItemProperty -Path "$uninstallKeyPath\$($subKey.PSChildName)" -ErrorAction SilentlyContinue).DisplayName
        $version = (Get-ItemProperty -Path "$uninstallKeyPath\$($subKey.PSChildName)" -ErrorAction SilentlyContinue).DisplayVersion

        if ($displayName -eq $targetDisplayName -and $version -ne $highestVersion) {
            Write-Host "Removing $($subKey.PSChildName): $displayName $version" -ForegroundColor Red
            Remove-Item -Path "$uninstallKeyPath\$($subKey.PSChildName)" -Force
        }
    }
}

$targetDisplayName = "Windows Agent"
$uninstallKeyPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

# Get all subkeys under [HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall]
$subKeys = Get-ChildItem -Path $uninstallKeyPath | Where-Object { $_.PSChildName -match "^\{.*\}$" }

# Create an empty array to store the versions of the target application
$versions = @()

# Loop through each subkey and check if it matches the target display name
foreach ($subKey in $subKeys) {
    $displayName = (Get-ItemProperty -Path "$uninstallKeyPath\$($subKey.PSChildName)" -ErrorAction SilentlyContinue).DisplayName
    if ($displayName -eq $targetDisplayName) {
        $version = (Get-ItemProperty -Path "$uninstallKeyPath\$($subKey.PSChildName)" -ErrorAction SilentlyContinue).DisplayVersion
        $versions += $version
    }
}

# Sort the versions in descending order
$versions = $versions | Sort-Object -Descending

Write-Host "Agents Found: $($versions.Count)" -ForegroundColor Cyan
foreach ($version in $versions) {
    Write-Host "$targetDisplayName $version" -ForegroundColor Cyan
}

if ($versions.Count -le 1) {
    $VersionsFound = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') INFO: No Multiple Agent versions found: $($versions.Count)"
    Write-Host "No remnants detected!" -ForegroundColor Green
} 
else {
    $VersionsFound = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') WARNING: Multiple Agent versions found: $($versions.Count)"
    $highestVersion = $versions[0] # Keep the highest version
    Cleanup-Agents
}