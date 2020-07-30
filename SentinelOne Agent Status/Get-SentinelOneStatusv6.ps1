$script:monitorStatus=""
$script:protectionStatus=""
$script:monitorBuildId=""
$script:enforcingSecurity=""
$script:agentLoadedStatus=""
$script:javaAgentStatus=""
$script:mitigPolicyName=""
$script:isInfected=""
$script:rebootLogLine=""
$script:dynamicEngineStatus=""
$script:installationStatus=""

$script:fullOutput=""
$script:successCode=0
$script:errorMsg=""

function checkSentinelOne($retryCount) {
#Run a discovery to see if SentinelOne Successfully Installed

	$script:s1_jsonObj = $null
	$sleepInSeconds = 5

	$attempt = 0
	while ($attempt -lt $retryCount) {
		try {
			$helper = New-Object -ComObject "SentinelHelper.1" -errorvariable s1Error
			$script:s1_jsonObj = $helper.GetAgentStatusJSON()
			$script:successCode = 0
			$script:errorMsg = ""
			break;
		} catch {
			$script:successCode = $Error[0].Exception.HResult
			$script:errorMsg = $s1Error
		}
		$attempt++
		if ($attempt -lt $retryCount) {
			sleep $sleepInSeconds
		}
	}
	return $s1_jsonObj

}

function getDynamicEngineStatus {

    try {
        $sentinelOneLogLine = (Get-WinEvent -MaxEvents 1 -FilterHashtable @{ProviderName='SentinelOne'; ID=1}).Message
    } catch {
        $sentinelOneLogLine = ""
    }
	
	Switch -Wildcard ($sentinelOneLogLine) {
		'Windows Agent is starting in slim mode*' { $script:dynamicEngineStatus="Disabled" }
		'Windows Agent is starting in full mode*' { $script:dynamicEngineStatus="Enabled" }
		Default { $script:dynamicEngineStatus="Unknown" }
	}

}

function ConvertFrom-Json20($item) {

	add-type -assembly system.web.extensions
	$ps_js=new-object system.web.script.serialization.javascriptSerializer

	#The comma operator is the array construction operator in PowerShell
	return ,$ps_js.DeserializeObject($item)

}

function setDefaultValues {
	$script:installationStatus = "No"
	$script:dynamicEngineStatus = "Unknown"
	$script:monitorBuildId = "Unknown"
	$script:mitigPolicyName = "Unknown"
	$script:protectionStatus = "Unknown"
	$script:enforcingSecurity = "Unknown"
	$script:agentLoadedStatus = "No Status Returned"
	$script:isInfected="Unknown"
}

function executeSentinelCOM {
	$retryCount = 2
	checkSentinelOne($retryCount)
	
	setDefaultValues
	
	if ($script:errorMsg -eq "") {
	
		$script:installationStatus = "Yes"
		$jsonObj = ConvertFrom-Json20($s1_jsonObj)
		Switch ($jsonObj.'self-protection-enabled') { 
			'true' { $script:protectionStatus = "Enabled" }
			'false' { $script:protectionStatus = "Disabled" }
			Default { $script:protectionStatus = "Unknown" }
		}
		
		$script:monitorBuildId = $jsonObj.'agent-version'
		
		Switch ($jsonObj.'enforcing-security') { 
			'true' { $script:enforcingSecurity = "Enabled" }
			'false' { $script:enforcingSecurity = "Disabled" }
			Default { $script:enforcingSecurity = "Unknown" }
		}
		
		Switch ($jsonObj.'agent-running') { 
			'true' { $script:agentLoadedStatus = "Running Normally" }
			'false' { $script:agentLoadedStatus = "Not Running" }
			Default { $script:agentLoadedStatus = "No Status Returned" }
		}
		
		Switch ($jsonObj.'active-threats-present') {
			'true' { $script:isInfected = "Infected" }
			'false' { $script:isInfected = "Clean" }
			Default { $script:isInfected = "Unknown" }
		}

		getDynamicEngineStatus
		
	}
}


executeSentinelCOM

"Self-Protection Status is " + $script:protectionStatus
"Sentinel Agent Status is " + $script:agentLoadedStatus
"Build of the Monitor Agent is " + $script:monitorBuildId
"Dynamic Engine Status is " + $script:dynamicEngineStatus
"Is EDR Installed " + $script:installationStatus
"Enforced Security Status is " + $script:enforcingSecurity
"Infected Status is " + $script:isInfected

$selfstatus = $script:protectionStatus
$agentstatus = $script:agentLoadedStatus
$dynamicengines = $script:dynamicEngineStatus
$isedrinstalled = $script:installationStatus
$buildid = $script:monitorBuildId
$enforcesecuritystatus = $script:enforcingSecurity
$infectedstatus = $script:isInfected
