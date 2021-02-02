# DeployTheNcentralAgent.ps1
#
# This script outputs the customer list, as well as the Registration Token for each customer.  The script prompts for parameters:
#    - N-Central server FQDN
#    - The JWT to be used to authenticate to N-central
#    - The CustomerID to be queried
#
# Created by: Chris Reid, Solarwinds MSP, with credit to Jon Czerwinksi and Kelvin Telegaar
# Date: Feb. 1st, 2021
# Version: 1.1
 
# Define the command-line parameters to be used by the script
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]$serverHost,
    [Parameter(Mandatory = $true)]$JWT,
    [Parameter(Mandatory = $true)]$SpecifiedCustomerID
)


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

 
# Generate a pseudo-unique namespace to use with the New-WebServiceProxy and
# associated types.
$NWSNameSpace = "NAble" + ([guid]::NewGuid()).ToString().Substring(25)
$KeyPairType = "$NWSNameSpace.EiKeyValue"
 
# Bind to the namespace, using the Webserviceproxy
$bindingURL = "https://" + $serverHost + "/dms2/services2/ServerEI2?wsdl"
$nws = New-Webserviceproxy $bindingURL -Namespace ($NWSNameSpace)
 
# Set up and execute the query
$KeyPair = New-Object -TypeName $KeyPairType
$KeyPair.Key = 'listSOs'
$KeyPair.Value = "False"
Try {
    $CustomerList = $nws.customerList("", $JWT, $KeyPair)
}
Catch {
    Write-Host "Could not connect: $($_.Exception.Message)"
    exit
}
 
$found = $False
$rowid=0
While ($rowid -lt $CustomerList.Count -and $found -eq $False)
{
    
    If($customerlist[$rowid].items[0].Value -eq [int]$SpecifiedCustomerID)
    {
        Foreach($rowitem In $CustomerList[$rowid].items)
        {
            If($rowitem.key -eq "customer.registrationtoken")
            {
                $RetrievedRegistrationToken = $rowitem.value
                If($RetrievedRegistrationToken -eq "")
                {
                    "Note that a valid Registration Token was not returned even though the customer was found. This happens when an agent install has never been downloaded for that customer. Try to download an agent from the N-Central UI and run this script again"
                }
            }
        }
    }
 
    $rowid++
}
$Customers = ForEach ($Entity in $CustomerList) {
    $CustomerAssetInfo = @{}
    ForEach ($item in $Entity.items) { $CustomerAssetInfo[$item.key] = $item.Value }
    [PSCustomObject]@{
        ID                = $CustomerAssetInfo["customer.customerid"]
        Name              = $CustomerAssetInfo["customer.customername"]
        parentID          = $CustomerAssetInfo["customer.parentid"]
        RegistrationToken = $CustomerAssetInfo["customer.registrationtoken"]
    }
}


 
# Uncomment this line if you wish to see the array of customers that has been found.
#$Customers | Sort-Object -Property ID | Format-Table -AutoSize
#$RetrievedRegistrationToken = ($Customers | Where-Object ID -eq $SpecifiedCustomerID).RegistrationToken
 
Write-Host "Here is the registration token for CustomerID" $SpecifiedCustomerID":" $RetrievedRegistrationToken -ForegroundColor Green
 
# Let's see if the Windows Agent installer has already been placed in the %TEMP% directory
If (!(Test-Path -Path "C:\Temp\windowsAgentSetup.exe")) {
    Write-Host "The Agent installer was not found in C:\Temp. Attempting download from N-central."
    $URI = "https://" + $serverHost + "/download/current/winnt/N-central/WindowsAgentSetup.exe"
    Invoke-WebRequest -Uri $URI -OutFile 'C:\Temp\WindowsAgentSetup.exe'
}
Else {
    Write-Host "Agent installer is located in C:\Temp."
}
# Now that we've got the registration token for the specified customer, let's use it to install the Windows Agent
Write-Host "Initiating the agent install."
# Start-Process -NoNewWindow -FilePath "C:\Temp\WindowsAgentSetup.exe" -ArgumentList "/s /v`" /qn CUSTOMERID=$SpecifiedCustomerID CUSTOMERSPECIFIC=1 REGISTRATION_TOKEN=$RetrievedRegistrationToken SERVERPROTOCOL=HTTPS SERVERADDRESS=$serverHost SERVERPORT=443`""
