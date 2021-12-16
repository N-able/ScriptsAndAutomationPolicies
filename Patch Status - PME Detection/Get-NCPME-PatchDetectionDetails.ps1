<#    
    ************************************************************************************************************
    Name: Get-NCPMEDetectionDetails
    Version: 0.1 (09th June 2021)
    Author: Prejay Shah (Doherty Associates)
    Purpose:    Surface PME Detection WIndow Detaisl via N-CEntrla logging
                eg Last successful patch detection windwo, last PMe Detection Error
                Allow for alerting based on days since last successful detection 
    Pre-Reqs:    PowerShell 2.0 
    Version History:    0.1.0. + Initial Public Release
    ************************************************************************************************************
#>

$Version = "0.1 (09th June 2021)"
Write-Host "Get-NCPME-PatchDetectionDetails $Version" -ForegroundColor Cyan
Write-Host ""
[datetime]$Today = $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Write-Host "Today: " -nonewline; Write-Host "$Today" -ForegroundColor Cyan

# Parse all Patch Detection Audit Logs found on device sorted in date order
$PatchDetectionAuditLogPath = "$env:programdata\MSPPlatform\PME\log\PatchDetectionAudit.log"
if ((test-path $PatchDetectionAuditLogPath) -and ((get-item $PatchDetectionAuditLogPath).length -gt '0')) {
$PatchDetectionAuditLog = get-childitem  "$PatchDetectionAuditLogPath*" | Sort-Object -Property lastwritetime | Get-Content
$PatchDetectionAuditLogScanData = $PatchDetectionAuditLog -match " Scan"
# $PatchDetectionAuditLog = Get-Content "$PatchDetectionAuditLogPath"
$PatchDetectionAuditLogEndScan = ($PatchDetectionAuditLogScanData -match "End Scan")[-1]
$EndScanArrayIndex = [array]::indexof($PatchDetectionAuditLogScanData,$PatchDetectionAuditLogEndScan)
    if ($PatchDetectionAuditLogEndScan -eq $null) {
    write-Host "No Succesful scans found!" -ForegroundColor Red
    $DayssincelastScan = '-1'
    }
    else {
        $PatchDetectionAuditLogStartScan = ($PatchDetectionAuditLogScanData -match "Start Scan")[-1]
        $StartScanArrayIndex = [array]::indexof($PatchDetectionAuditLogScanData,$PatchDetectionAuditLogStartScan)
        if ($StartScanArrayIndex = ($endScanArrayIndex -1)) {
            $indexmessage = "INFO: Start/End Scan Dates are in correct order: $StartScanArrayIndex/$EndScanArrayIndex"
            
        }
        else {
            $indexmessage =  "INFO: Start/End Scan Dates are in incorrect order: $StartScanArrayIndex/$EndScanArrayIndex"
        }
        Write-Host "$indexmessage" -ForegroundColor Yellow
        $PatchDetectionAuditLogStartScanDate = $PatchDetectionAuditLogStartScan.split(' ')[1,2]
        $PatchDetectionAuditLogEndScanDate = $PatchDetectionAuditLogEndScan.split(' ')[1,2]
        $startscandate = [datetime]::parseexact("$PatchDetectionAuditLogStartScanDate", 'yyyy-MM-dd HH:mm:ss,fff',$null)
        $endscandate = [datetime]::parseexact("$PatchDetectionAuditLogEndScanDate", 'yyyy-MM-dd HH:mm:ss,fff',$null)
        if ($startscandate -gt $endscandate) {
            Write-Host "Latest Scan has not finished. Obtaining details for previous completed scan" -ForegroundColor Yellow
            $PatchDetectionAuditLogStartScan = ($PatchDetectionAuditLogScanData)[($endScanArrayIndex -1)]
            $PatchDetectionAuditLogStartScanDate = $PatchDetectionAuditLogStartScan.split(' ')[1,2]
            $startscandate = [datetime]::parseexact("$PatchDetectionAuditLogStartScanDate", 'yyyy-MM-dd HH:mm:ss,fff',$null)
        }
        $scanduration = [math]::floor((New-TimeSpan -Start $startscandate -End $endscandate).totalminutes)
        Write-Host "$PatchDetectionAuditLogStartScan" -ForegroundColor Cyan
        Write-Host "Last Patch Detection Scan Start: " -nonewline; Write-Host "$PatchDetectionAuditLogStartScanDate" -ForegroundColor Green 
        Write-Host "$PatchDetectionAuditLogEndScan" -ForegroundColor Cyan
        Write-Host "Last Patch Detection Scan End: " -nonewline; Write-Host "$PatchDetectionAuditLogEndScanDate" -ForegroundColor Green 
        if ($scanduration -lt 0){
            $ScanMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Error: Problem computing scan duration due to multiple scans that have not finished running!"
            Write-Host $ScanMessage -ForegroundColor Red 
        }
        $DayssincelastScan = "{0:N2}" -f ($Today - $EndScanDate).TotalDays      
    }

    $PatchDetectionAuditLogError = ($PatchDetectionAuditLog -match "Error")[-1]
    if ($PatchDetectionAuditLogError -eq $null) {
    Write-Host "No Errors Found!" -ForegroundColor Green
    $DayssincelastError = '0'
    }
    else {
        Write-Host "$PatchDetectionAuditLogError" -ForegroundColor Cyan
        $PatchDetectionAuditLogErrorDate = $PatchDetectionAuditLogError.split(' ')[1,2]
        $errordate = [datetime]::parseexact("$PatchDetectionAuditLogErrorDate", 'yyyy-MM-dd HH:mm:ss,fff',$null)
        Write-Host "Last Patch Detection Error: " -nonewline; Write-Host "$PatchDetectionAuditLogErrorDate" -ForegroundColor Green 
        $DayssincelastError = "{0:N2}" -f ($Today - $ErrorDate).TotalDays
    }
    Write-Host ""
    if ([int]$Dayssincelastscan -le '7') {
        Write-Host "Days Since Last Patch Detection Scan: " -nonewline; Write-Host $Dayssincelastscan -ForegroundColor Green
    }
    else {
        Write-Host "Days Since Last Patch Detection Scan: " -nonewline; Write-Host $Dayssincelastscan -ForegroundColor Red
    }

    Write-Host "Days since Last Patch Detection Error: " -nonewline; Write-Host $dayssincelasterror -ForegroundColor Green

    if ($errordate -gt $endscandate) {
        $StatusMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Last Scan Failed!"
        Write-Host "Status: $Statusmessage" -ForegroundColor Red
    }
    else {
        $StatusMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Last Scan Successful! - $PatchDetectionAuditLogEndScanDate ($scanduration minutes)"
        Write-Host "Status: $Statusmessage" -ForegroundColor Green
    }
}
else {
    $StatusMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ERROR: No Data found in PatchDetectionAudit.Log"
    write-Host "Status: $Statusmessage" -ForegroundColor Red
    #$startscandate = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') N/A - ERROR: No Data found in PatchDetectionAudit.Log"
    #$endscandate = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') N/A - ERROR: No Data found in PatchDetectionAudit.Log"
    $PatchDetectionAuditLogStartScan = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') N/A - ERROR: No Start Scan Data found in PatchDetectionAudit.Log"
    $PatchDetectionAuditLogEndScan = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') N/A - ERROR: No End Scan Data found in PatchDetectionAudit.Log"
    $PatchDetectionAuditLogError = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') N/A - ERROR: No Error Data found in PatchDetectionAudit.Log"
    $DayssincelastScan = '-1'
    $DayssincelastError = '-1'
    $scanduration = '-1'
}

if ($scanduration -lt 0){
    $ScanMessage = "Error: Problem computing last scan duration due to multiple scans that have not finished running!"
    $Status = $StatusMessage + "<br>$ScanMessage<br>$indexmessage"
    $scanduration = '-1'
}
else {
    $Status = $StatusMessage + "<br>$indexmessage"
}
write-host ""
Write-Host "Overall Status:`n$Status"  -ForegroundColor Cyan