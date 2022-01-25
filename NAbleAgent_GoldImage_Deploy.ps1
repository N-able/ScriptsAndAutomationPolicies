<#
.Author
    Ron Oberjohann - https://www.advdat.com
.Version
    	1.0 29NOV21 
    	1.1 02DEC21
		BugFixes 
    	1.2 20DEC21
    		Added MSI logging
	1.3 21DEC21
		Modified scheduled task creation to correct elevation issue
		Must generate registration token in console for customer prior to deployment, otherwise lookup with fail
			18JAN22 NAble is tracking this issue as NCCF-14783
.Summary
    This script prompts for specific information and stores it as an xml file.
    Conceived as a setup script for the deployment of the NAble agent on a provisioned image.
    Example found at https://searchwindowsserver.techtarget.com/tip/Using-PowerShell-to-create-XML-documents
	https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/DeployTheNCAgent/DeployTheNCAgent.ps1
.Credits
	Mark Khalil - https://www.linkedin.com/in/mark-khalil-52b58859/
	Chris Reid - https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/DeployTheNCAgent/DeployTheNCAgent.ps1
	Adam Bertram - https://adamtheautomator.com/
    
#>

#region Obtain/Define Values

    $Path = Read-Host "Enter Path to store configuration and executable (Ex. C:\Temp)"
        #Create Path if not present
            If(-not (Test-Path $Path)){
                New-Item -Path $Path -ItemType directory
            }
    $ConfigPath = ($Path + '\NAbleAgentConfig.xml')
    $AgentFile = ($Path + '\WindowsAgentSetup.exe')
    $CustomerID = Read-Host "Enter CustomerID"
    $JWT = Read-Host "Enter JWT"
    $ServerHost = Read-Host "Enter Server"

#endregion


#region Create AgentConfig.xml
    #Generate XML document
        $xmlWriter = New-Object System.XMl.XmlTextWriter($ConfigPath,$Null)
        $xmlWriter.Formatting = 'Indented'
        $xmlWriter.Indentation = 1
        $XmlWriter.IndentChar = "`t"

    #create a well-structured XML file by writing the declaration with version 1.0
        $xmlWriter.WriteStartDocument()

    # Comment regarding the tag
        $xmlWriter.WriteComment('N-Able Agent Configuration')

    #Item in list
        $xmlWriter.WriteStartElement('Agent')
        $XmlWriter.WriteAttributeString('Customer', $CustomerID)
        $xmlWriter.WriteAttributeString('Server_Host', $ServerHost)
        $xmlWriter.WriteAttributeString('JWT', $JWT)
        $xmlWriter.WriteAttributeString('Path', $Path)
        $xmlWriter.WriteAttributeString('AgentFile', $Agentfile)
        $xmlWriter.WriteEndElement()
    $xmlWriter.WriteEndDocument()
    $xmlWriter.Flush()
    $xmlWriter.Close()
#endregion

#region Obtain Installer

# Check for Windows Agent installer in the defined directory
If (-not(Test-Path -Path $AgentFile)) {
    Write-Host "The Agent installer was not found in $Path. Attempting download from $ServerHost"
    $URI = "https://" + $serverHost + "/download/current/winnt/N-central/WindowsAgentSetup.exe"
    Invoke-WebRequest -Uri $URI -OutFile $AgentFile
}
Else {
    Write-Host "Agent installer is found in $Path."
}

#endregion

#region Create script to run as scheduled task on boot
$Content = @()
$Content += @'
#########################
#region variables
#Variables from xml
'@

$Content += "$" + "xml = [xml](Get-Content $ConfigPath)"

$Content += @'
$SpecifiedCustomerID = $xml.agent.Customer
$JWT = $xml.agent.JWT
$ServerHost = $xml.agent.Server_Host
$Path = $xml.agent.Path
$Agentfile = $xml.agent.AgentFile
$LogFile="$path\NAbleStart.log"

#endregion

#region Query for info

#Specify ProtocolType
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
        Start-Sleep 15
	exit
        }
 
$found = $False
$rowid=0
While ($rowid -lt $CustomerList.Count -and $found -eq $False){
    	If($customerlist[$rowid].items[0].Value -eq [int]$SpecifiedCustomerID){
                Foreach($rowitem In $CustomerList[$rowid].items){
                    If($rowitem.key -eq "customer.registrationtoken"){
                        $RetrievedRegistrationToken = $rowitem.value
                        If($RetrievedRegistrationToken -eq ""){
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

        #endregion
        
        #Execute
        
        If($RetrievedRegistrationToken){
        Write-Host RegistrationTokenFound
      Start-Process -NoNewWindow -FilePath $AgentFile -ArgumentList "/s /v`" /l*v $Logfile /qn CUSTOMERID=$SpecifiedCustomerID CUSTOMERSPECIFIC=1 REGISTRATION_TOKEN=$RetrievedRegistrationToken SERVERPROTOCOL=HTTPS SERVERADDRESS=$serverHost SERVERPORT=443`""
        }
        Else{
        	Write-Host RegistrationTokenNotFound
             }


'@

#Output to file
Set-Content -Path ($path + "\NableStart.ps1") -Value $Content

#endregion

#Create ScheduledTask
$Script=($path + "\NableStart.ps1")
$A = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-NoLogo -ExecutionPolicy Unrestricted -NoProfile -File $script"
$T = New-ScheduledTaskTrigger -AtStartup
$P = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$S = New-ScheduledTaskSettingsSet
$D = New-ScheduledTask -Action $A -Principal $P -Trigger $T -Settings $S
Register-ScheduledTask -TaskName "Nable Agent Startup" -InputObject $D
