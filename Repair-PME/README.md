![Repair-PME Logo](https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/WikiFiles/Repair-PME/Repair-PME_Logo_Small.png)
***

### What is Repair-PME?

Repair-PME is a script provided to the community that repairs Solarwinds MSP Patch Management Engine (PME) used by Solarwinds N-Central. If for any reason it becomes corrupt, fails an upgrade or installation process PME-Repair should be able to fix it in the most scenarios. 

### Why was Repair-PME created?

Repair-PME was created as I was getting frustrated with having to spend time trying resolve various issues where PME had broken and subsequently caused **Patch Status v2** to become **misconfigured** in N-Central. Obviously having this in the misconfigured state is very bad as you cannot determine what's going on with patching. Please note is only designed to fix issues where PME is broken and reinstall is required to resolve. There may be cases where you will need to engage support to assist further. An example of an error that can be fixed with this tool.

![Patch_Status_v2_Misconfigured](https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/WikiFiles/Repair-PME/Patch_Status_v2_Misconfigured.png)

### What does Repair-PME do?

Repair-PME does the following with logic, error handling and event logging to operate as user-friendly as possible. Any errors during execution of the script will throw to PowerShell and will also be reported to the event log (Event ID 100).

* Checks to ensure script is run elevated (as an administrator) to ensure all necessary actions can be performed.
* Writes an application event log from source 'Repair-PME' with event ID 100 reporting script has started.
* Gets operating system version, operating system architecture and PowerShell version.
* Checks if script is up to date and will throw a non-terminating error if it is outdated.
* Performs connectivity tests to destinations required for PME. Download of PMESetup will be obtained via HTTP instead of HTTPS if issues with HTTPS connectivity is detected.
* Performs certificates test to HTTPS destination required for PME (sis.n-able.com). Test will be bypassed if issues with HTTPS connectivity is detected.
* Checks if N-Central Agent is installed, reports status and compatibility with PME.
* Checks if PME is already installed and reports status.
* Checks if PME has had a recent install. If a recent install (such as an auto-update) has occured within the configured period (2 days) then script will bypass the update pending check below to allow a force install. This can be changed, see settings below for further information.
* Checks if PME has an update pending and reports status. If an update is pending within the configured period (2 days) then script will be aborted. This can be changed, see settings below for further information.
* Invokes PME Diagnostics Tool and silently saves the log capture to to **'C:\ProgramData\SolarWinds MSP\Diagnostic Logs\'** (PME version 1.x) or **'C:\ProgramData\MspPlatform\Diagnostic Logs\'** (PME version 2.x and above). These logs can be given to Solarwinds support for further troubleshooting hopefully resolving any bugs to make future PME releases more robust.
* Terminates any currently running instances of **PMESetup,** **CacheServiceSetup,** **FileCacheServiceAgentSetup,** **RPCServerServiceSetup,** **RequestHandlerAgentSetup**, **_iu14D2N,** and **unins000**.
* Stops the PME services. **SolarWinds.MSP.PME.Agent.PmeService, PME.Agent.PmeService, SolarWinds.MSP.RpcServerService and SolarWinds.MSP.CacheService**. If operation times out they will be forcefully terminated. 
* Cleanup cached files from **C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService,** **C:\ProgramData\SolarWinds MSP\SolarWinds.MSP.CacheService\cache,** **C:\ProgramData\MspPlatform\FileCacheServiceAgent** and **C:\ProgramData\MspPlatform\FileCacheServiceAgent\cache** if applicable.
* Checks existing PME config and informs of possible misconfigurations (cache size and fallback to external sources).
* Obtains, checks (SHA-256 Hash) and downloads (if required) the latest available version of PME from sis.n-able.com if not verified locally.
* Silently installs PME (PME Agent, Cache Service and RPC Server Service) and saves the install logs to **'C:\ProgramData\SolarWinds MSP\Repair-PME\'** (PME version 1.x) or **'C:\ProgramData\MspPlatform\Repair-PME\'** (PME version 2.x and above).
* Checks and reports all PME services are installed and running post-installation.
* Writes an application event log from source 'Repair-PME' with event ID 100 reporting script has ended.

### Settings

As of release 0.2.0.0 there are three user changeable settings which can be found at the beginning of the script in the settings section.

* **$RepairAfterUpdateDays** - Change this variable to number of days (must be a number!) to begin repair after new version of PME is released. Default is 2. Repair-PME will abort if an update is pending within this period.

* **$ForceRepairRecentInstallDays** - Change this variable to number of days (must be a number!) within a recent install to allow a force repair. This will bypass the update pending check. Default is 2. Ensure this is equal to $RepairAfterUpdateDays.

* **$UpdateCheck** - Change this variable to turn off/on update check of the Repair-PME script. Default is Yes. To turn this off set it to No.

### Important

Please ensure you rescan any custom PME monitoring you may have and run a patch detection after running this script to ensure Patch Status v2 is operating fully again.

### System Requirements

**Internet Connectivity:**

An internet connection is required for this script to reach out to

HTTP (Port 80)
* _sis.n-able.com_
* _download.windowsupdate.com_
* _fg.ds.b1.download.windowsupdate.com_

HTTPS (Port 443)
* _sis.n-able.com_
* _raw.githubusercontent.com_

**Operating System:**
* _Any OS that can install the Solarwinds MSP Patch Management Engine (PME) and is officially supported by Solarwinds (this can be found in the N-Central release notes)._

**PowerShell:**
* Required: _2.0+_

**N-Central:**
* Required: _12.2.0.274+_

### Can I use Repair-PME in an Automation Policy (AMP) within N-Central?

Yes, just add the code to a '**Run Powershell Script**' object in Automation Manager, save the AMP as 'Repair-PME' and upload to your N-Central server. It is the recommended method of using this script via N-Central.

### Can I use Repair-PME from PowerShell interactively or via N-Central?

Yes, if you wish to do so. Please be aware your execution policy is set to allow this to run though.

### Can I use Repair-PME for self-healing within N-Central?

Yes, but if this is executed during the wait period (2 days by default) of when an update has been released but has yet to be installed the script will abort with an error as it is recommended this is done gracefully via the built-in update mechanism. The only exception to this is if an install has occured in the last 2 days. This is to account for situations where an auto-update does not complete succesfully.

If you want to be made aware of when the script is out of date notifcations should be setup to notify on 'Failure' and 'Send task output file in to email recipients'.

It is recommended this is used as self-healing in conjunction with Prejay Shah's **[**'Patch Status - PME' AMP**](https://github.com/N-able/CustomMonitoring/tree/f007703830dab88eb7fd710a84c768d0ff119e70/N-Central%20PME%20Services)**.

### Where can I get the latest version of Repair-PME?
**https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/Repair-PME/Repair-PME.ps1**

### Demonstration
Below is a GIF of Repair-PME in action run from PowerShell.

![Repair-PME Demo](https://github.com/N-able/ScriptsAndAutomationPolicies/blob/master/WikiFiles/Repair-PME/Repair-PME-Demo.gif)

### Feedback or Issues?

Feedback is always appreciated, and I look to improve this script where I can. If you have any issues with it or would like to suggest improvements, please do raise an Issue in GitHub. Alternatively reach out to the maintainer (Ashley How) in the N-Able Slack Community (**n-able.slack.com**).
