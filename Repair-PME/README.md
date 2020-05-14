![Repair-PME Logo](https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/WikiFiles/Repair-PME/Repair-PME_Logo_Small.png)
***

### What is Repair-PME?

Repair-PME is a script provided to the community that repairs Solarwinds MSP Patch Management Engine (PME) used by Solarwinds N-Central. If for any reason reason it becomes corrupt, fails an upgrade or installation process PME-Repair should be able to fix it in the majority of scenarios. 

### Why was Repair-PME created?

Repair-PME was created as I was getting frustrated with having to spend time trying resolve various issues where PME had broken and subsequently caused **Patch Status v2** to become **misconfigured** in N-Central. Obviously having this in the misconfigured state is very bad as you cannot determine what's going on with patching. An example of an error that can be fixed with this tool.

![Patch_Status_v2_Misconfigured](https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/WikiFiles/Repair-PME/Patch_Status_v2_Misconfigured.png)

### What does Repair-PME do?

Repair-PME does the following with logic, error handling and event logging to operate as user-friendly as possible:
* Writes an application event log from source 'Repair-PME' with event ID 100 reporting script has started.
* Checks to ensure script is run elevated (as an administrator) to ensure all necessary actions can be performed.
* Performs connectivity tests to destinations required for PME (PowerShell 4.0 or above required). Tests will be skipped and download of PMESetup will be obtained via HTTP instead of HTTPS if a lower version is detected.
* Invokes Solarwinds Diagnostics Tool and silently saves the log capture to **C:\ProgramData\SolarWinds MSP\Repair-PME\Diagnostic Logs**. These logs can be given to Solarwinds support for further troubleshooting hopefully resolving any bugs to make future PME releases more robust.
* Terminates any currently running instances of **PMESetup,** **CacheServiceSetup** and **RPCServerServiceSetup**.
* Cleanup cached files from **C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService** and **C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache**.
* Obtains, checks (SHA-256 Hash) and downloads (if required) the latest available version of PME from sis.n-able.com if not verified locally.
* Silently installs PME (PME Agent, Cache Service and RPC Server Service).
* Writes an application event log from source 'Repair-PME' with event ID 100 reporting script has ended.

### Important

Please ensure you rescan and run a patch detection after running this script to ensure patch management is operating fully again.

### System Requirements

**Operating System:**
* _Any OS that can install the Solarwinds MSP Patch Management Engine (PME) and is officially supported by Solarwinds (this can be found in the N-Central release notes)._

**PowerShell:**
* Required: _2.0_
* Optional: _4.0+_ (required for connectivity tests only)

### Can I use Repair-PME in an Automation Policy (AMP) within N-Central?

Yes, just add the code to a '**Run Powershell Script**' object in Automation Manager, save the AMP as 'Repair-PME' and upload to your N-Central server. It is the recommended method of using this script via N-Central.

### Can I use Repair-PME from PowerShell interactively or via N-Central?

Yes, if you wish to do so. Please be aware your execution policy is set to allow this to run though.

### Can I use Repair-PME for self-healing within N-Central?

The answer is yes but with caution. If you are using the **[**Get-PMEServices**](https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/N-Central%20PME%20Services/Get-PMEServices.ps1)** script or similar code to monitor PME status it will threshold if it is out of date while a update is pending. An update normally takes up to 24 hours before it installs from time of upload on sis.n-able.com. During this time it may trigger your custom service to fail/warn. If you set to self-heal with this script it will force an upgrade rather than wait for the agent to gracefully update it.

### Where can I get the latest version of Repair-PME?
**https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/Repair-PME.ps1**

### Demonstration
Below is an GIF of Repair-PME in action run from PowerShell.

![Repair-PME Demo](https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/WikiFiles/Repair-PME/Repair-PME-Demo.gif)

### Feedback or Issues?

Feedback is always appreciated and I look to improve this script where I can. If you have any issues with it or would like to suggest improvements please do raise an Issue in GitHub. Alternatively reach out to the maintainer (Ashley How) in the N-Able Slack Community (**n-able.slack.com**).
