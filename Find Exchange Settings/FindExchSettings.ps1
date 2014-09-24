Param($emailAddress = "logs@tiger.net.nz", $smtpserver="Mail.Global.FrontBridge.com")

$results = get-wmiobject win32_logicaldisk -filter "drivetype=3" | select-object deviceid

$csvCollection = @()

$ErrorActionPreference = "SilentlyContinue"

foreach ($device in $results) 
{
$list = Get-ChildItem  $device.deviceid -Recurse | where{$_.Extension -match "edb|stm"}
    foreach ($item in $list)
    {
$csvItem = "" | select Directory,File,Size
	$csvItem.directory = $item.directoryname
	$csvItem.file = $item.name
	$csvItem.size = $item.length
	$csvCollection += $csvItem
    }
}

$file = "c:\Ex_DB_Details.csv"
$file2 = "c:\Ex_Reg_Details.txt"
$csvCollection | Export-CSV $file

$computername = gc env:computername
$registry = "HKLM:\System\CurrentControlSet\Services\MSExchangeIS\" + $computername
Get-ChildItem $registry | ForEach-Object {Get-ItemProperty $_.pspath} | out-file $file2

$msg = new-object Net.Mail.MailMessage
$att = new-object Net.Mail.Attachment($file)
$att2 = new-object Net.Mail.Attachment($file2)
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "administrator@"+$computername
$msg.To.Add($emailAddress)
$msg.Subject = "Exchange Details"
$msg.Body = "These are the general details for the Exchange Server."
$msg.Attachments.Add($att)
$msg.Attachments.Add($att2)
$smtp.Send($msg)
$att.Dispose()

remove-item $file
remove-item $file2