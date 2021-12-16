Function UnzipFile([string] $ZipPath)
{
    [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

    Write-Host "Unzipping saved zip attachment [$ZipPath] containing script output..."

    $zipFile = (Get-ChildItem $ZipPath)
    $openZipFile = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    $openZipFileEntries = $openZipFile.Entries

    $i = 0

    $extractedFiles = @()

    Write-Host "`tUnzipping [$ZipPath]..."
    foreach ($file in $openZipFileEntries)
    {
        Write-Host "`t`tExtracting [$file]..."

        $destFilePath = Join-path $zipFile.Directory.FullName "$($zipFile.BaseName)_$i.txt"
        [IO.Compression.ZipFileExtensions]::ExtractToFile($file, $destFilePath)

        $extractedFiles += $destFilePath
    }

    $openZipFile.Dispose()
    return $extractedFiles
}

Function WriteScriptOutput([string] $CustomerName, [string] $DeviceName, [string] $ScriptName, [string] $Output, [bool] $NoHeader, [string] $OutputDir, [string] $ScriptStartTime, [string] $ScriptEndTime)
{
    $outputPath += "$OutputDir\$CustomerName\$ScriptName\"

    if (-Not(Test-Path $outputPath))
    {
        New-Item -ItemType Directory $outputPath
    }
    
    $scriptStartDateTime = New-Object DateTime
    $scriptEndDateTime = New-Object DateTime
    
    $filenameTimestamp = (Get-Date).ToString("dd-MM-yyyy")
    $startTimestamp = $ScriptStartTime

    if ([DateTime]::TryParse($ScriptStartTime, [ref] $scriptStartDateTime))
    {
        $filenameTimestamp = $scriptStartDateTime.ToString("dd-MM-yyyy")
        $startTimestamp = $scriptStartDateTime.ToString("HH:mm:ss dd/MM/yyyy")
    }

    [DateTime]::TryParse($ScriptEndTime, [ref] $scriptEndDateTime)

    $outputFilename = "$outputPath$filenameTimestamp.txt"
    $consoleStartMarker = "Output:"

    Write-Host "Writing script output for [$CustomerName - $DeviceName to file [$outputFilename]"

    if (-Not($NoHeader))
    {
        $header = "---------------------------------------------------------`r`n$CustomerName - $DeviceName [$startTimestamp]`r`n---------------------------------------------------------`r`n"
        Add-Content -LiteralPath $outputFilename $header
    }

    $outputIndex = $Output.IndexOf($consoleStartMarker)

    if ($outputIndex -gt -1)
    {
        $Output = $Output.Substring($outputIndex + $consoleStartMarker.Length).Trim()
    }

    $Output |  Add-Content -LiteralPath $outputFilename
}
##############################################################################################################################################################################

$outputDir = "\\server\some\path\"

$mailServer = ""
$mailPort = 110
$mailUser = ""
$mailPass = ""

$isSSL = $false

Add-Type -LiteralPath "AE.Net.Mail.dll" | Out-Null

$path = "$((Get-Item -Path ".\" -Verbose).FullName)\"

$customerStartMarker = "Customer:"
$customerEndMarker = "Executed By:"
$deviceStartMarker = "Device:"
$typeStartMarker = "Type:"
$startTimeMarker = "Start Time:"
$endTimeMarker = "End Time:"
$returnCodeMarker = "Return Code:"

$popClient = New-Object "AE.Net.Mail.Pop3Client" -ArgumentList @($mailServer, $mailUser, $mailPass, $mailPort, $isSSL)

$messageCount = $popClient.GetMessageCount()
$processedCount = 0

for ($i = 0; $i -lt $messageCount; $i++)
{
    $msg = $popClient.GetMessage($i, $false)

    $subject = $msg.Subject
    $body = $msg.Body
    $attachments = $msg.Attachments

    Write-Host "Parsing email [$($subject)]..."
    
    $customerMatch = [regex]::match($body, "$customerStartMarker\s+(.*(?=`r))\s+$customerEndMarker")
    $deviceMatch = [regex]::match($body, "$deviceStartMarker\s+(.*(?=`r))\s+")
    $scriptNameMatch = [regex]::match($body, "(?<=$typeStartMarker)\s+.*\s+\[(.*)]")
    $scriptStartTimeMatch = [regex]::match($body, "$startTimeMarker\s+(.*)\s+$endTimeMarker")
    $scriptEndTimeMatch = $scriptStartTimeMatch = [regex]::match($body, "$endTimeMarker\s+(.*)\s+$returnCodeMarker")

    if ($attachments.Count -gt 0)
    {
        $attachments | ? { $_.IsAttachment -eq $true } | % {
             Write-Host "Saving script output attachment [$($_.FileName)]..."
             $attachmentPath = "$path\$($_.FileName)"
             $_.Save($attachmentPath)

             # TODO: Probably should validate the attachment is a zip file..
             $extractedFiles = UnzipFile -ZipPath $attachmentPath
             foreach ($extractedFile in $extractedFiles)
             {
                $logOutput = (Get-Content -Path $extractedFile) -Join "`r`n"
                WriteScriptOutput -CustomerName  $customerMatch.Groups[1].Value -DeviceName $deviceMatch.Groups[1].Value -ScriptName $scriptNameMatch.Groups[1].Value -Output $logOutput -OutputDir $outputDir -ScriptStartTime $scriptStartTimeMatch.Groups[1].Value -ScriptEndTime $scriptEndTimeMatch.Groups[1].Value
             }

             Remove-Item -ErrorAction SilentlyContinue -Force $extractedFiles
             Remove-Item -ErrorAction SilentlyContinue -Force $attachmentPath
        }
    }
    else
    {
        WriteScriptOutput -CustomerName  $customerMatch.Groups[1].Value -DeviceName $deviceMatch.Groups[1].Value -ScriptName $scriptNameMatch.Groups[1].Value -Output $body -OutputDir $outputDir -ScriptStartTime $scriptStartTimeMatch.Groups[1].Value -ScriptEndTime $scriptEndTimeMatch.Groups[1].Value
    }

    $processedCount++
    $popClient.DeleteMessage($msg)
}

$popClient.Disconnect()

Write-Host "DONE!"
Write-Host "$processedCount of $messageCount emails processed!"