# DeployTheNcentralAgent.ps1
#
# This script outputs the customer list, as well as the Registration Token for each customer.  The script prompts for parameters:
#    - N-Central server FQDN
#    - The JWT to be used to authenticate to N-central
#    - The CustomerID to be queried
#
# Created by: Chris Reid, Solarwinds MSP, with credit to Jon Czerwinksi and Kelvin Telegaar
# Date: Sept 18th, 2020
# Version: 1.0


# Define the command-line parameters to be used by the script
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]$serverHost,
    [Parameter(Mandatory = $true)]$JWT,
    [Parameter(Mandatory = $true)]$SpecifiedCustomerID
)

# Generate a pseudo-unique namespace to use with the New-WebServiceProxy and
# associated types.
$NWSNameSpace = "NAble" + ([guid]::NewGuid()).ToString().Substring(25)
$KeyPairType = "$NWSNameSpace.T_KeyPair"

# Bind to the namespace, using the Webserviceproxy
$bindingURL = "https://" + $serverHost + "/dms/services/ServerEI?wsdl"
$nws = New-Webserviceproxy $bindingURL -Namespace ($NWSNameSpace)

# Set up and execute the query
$KeyPair = New-Object -TypeName $KeyPairType
$KeyPair.Key = 'listSOs'
$KeyPair.Value = "false"
Try {
    $CustomerList = $nws.customerList($username, $JWT, $KeyPair)
}
Catch {
    Write-Host "Could not connect: $($_.Exception.Message)"
    exit
}

# Set up the "Customers" array, then populate
$Customers = ForEach ($Entity in $CustomerList) {
    $CustomerAssetInfo = @{}
    ForEach ($item in $Entity.Info) { $CustomerAssetInfo[$item.key] = $item.Value }
    [PSCustomObject]@{
        ID                = $CustomerAssetInfo["customer.customerid"]
        Name              = $CustomerAssetInfo["customer.customername"]
        parentID          = $CustomerAssetInfo["customer.parentid"]
        RegistrationToken = $CustomerAssetInfo["customer.registrationtoken"]
    }
}

# Uncomment this line if you wish to see the array of customers that has been found.
# $Customers | Sort-Object -Property ID | Format-Table -AutoSize
$RetrievedRegistrationToken = ($Customers | Where-Object ID -eq $SpecifiedCustomerID).RegistrationToken

Write-Host "Here is the registration token for CustomerID" $SpecifiedCustomerID":" $RetrievedRegistrationToken -ForegroundColor Green

# Let's see if the Windows Agent installer has already been placed in the %TEMP% directory
If (!(Test-Path -Path "C:\Temp\windowsAgentSetup.exe")) {
    Write-Host "The Agent installer was not found in C:\Temp. Attempting download from N-central."
    $URI = "https://" + $serverHost + "/download/2020.1.0.202/winnt/N-central/WindowsAgentSetup.exe"
    Invoke-WebRequest -Uri $URI -OutFile 'C:\Temp\WindowsAgentSetup.exe'
}
Else {
    Write-Host "Agent installer is located in C:\Temp."
}
# Now that we've got the registration token for the specified customer, let's use it to install the Windows Agent
Write-Host "Initiating the agent install."
Start-Process -NoNewWindow -FilePath "C:\Temp\WindowsAgentSetup.exe" -ArgumentList "/s /v`" /qn CUSTOMERID=$SpecifiedCustomerID CUSTOMERSPECIFIC=1 REGISTRATION_TOKEN=$RetrievedRegistrationToken SERVERPROTOCOL=HTTPS SERVERADDRESS=$serverHost SERVERPORT=443`""
