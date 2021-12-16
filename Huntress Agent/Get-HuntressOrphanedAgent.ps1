<#
   **********************************************************************************************************************************
    Name:            Get-HuntressOrphanedAgent.ps1
    Version:         0.1 (04/012/2021)
    Purpose:         Get Huntress Agent Oprhan Status based on Log analysis
    Created by:      Prejay Shah
    Thanks to:       The Team at Huntress for creating the initial CW Automate Script this is adapted from.
                     https://support.huntress.io/hc/en-us/articles/4404004935059-ConnectWise-Automate-Remote-monitor-orphaned-agent-

    Version History: 0.1.0.0 - Initial Release.
   **********************************************************************************************************************************
#>

$Version = '0.1'
$VersionDate = '(04/12/2021)'
$AgentStatus = 'ACTIVE'

Write-Host "Get-HuntressOrphanedAgent $Version ($VersionDate)" -ForegroundColor Green

$file = 'C:\Program Files (x86)\Huntress\HuntressAgent.log'

    if (-not(Test-Path -Path $file -PathType Leaf)) {
        $file = 'C:\Program Files\Huntress\HuntressAgent.log'
        Get-Content $file -Tail 10 | ForEach-Object { if ($_ -match '401') {$agentstatus = 'ORPHANED'}}
    }
    else {
    Get-Content $file -Tail 10 | ForEach-Object { if ($_ -match '401') {$agentstatus = 'ORPHANED'}}
    }

    if ($agentstatus -eq 'ACTIVE') {
        Write-Host "Agent Status: " -foregroundcolor Cyan -nonewline; Write-Host "$agentstatus" -ForegroundColor Green
    }
    else {
        Write-Host "Agent Status: " -foregroundcolor Cyan -nonewline; Write-Host "$agentstatus" -ForegroundColor Red
    }