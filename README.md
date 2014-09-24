Scripts and Automation Policies
===============================

Community contributed standalone scripts and automation policies, for use with the N-Central product.

<b>Use these at your own risk</b>

Master list

| Name             | Description         | Script Type   |
| ---------------- | ------------------- | ------------- | 
3Ware RAID Controller | This script will turn on the cache of a 3Ware RAID controller. This script requires the presence of the 3Ware command-line interface. | Batch |
Add or Delete DNS Record | AM1.5 - Add or Delete specific DNS Record | AMP | 
Advanced Device Details |  AM1.5 - Retrieve System Enclosure, Processor, Physical Memory, BaseBoard, Battery, CDRom, Computer System, System TimeZone, Windows Product Key, Default Browser, OS Service Pack, OS Architecture and Windows Features information. | AMP |
Agent Deploy | Agent Download - A modified Cleanup and Deploy. This modified Agent Cleanup and Deploy script provided by Evan Morrissey can be used to automate updating the Agent installer and script parameters when the N-Central server is updated. | AMP |
Agent_Install | | Batch |
Allow SNMP Queries from 127.0.0.1 | Very useful for those upgrading from 6.5 to 6.7. | VBS |
AV Status | Version 1.87 | VBS |
AVG - Quick or Full Scan | These two scripts will cause AVG 8.x to auto-update it's definitions, and then run either a quick scan or a full scan. | VBS |
Backup and Clear Event Logs | There are 6 scripts in this package - 3 that just backup the System/Application/Security event logs, and another 3 that backup those logs and then clear them. | VBS |
Backup Cisco IOS Config | This new version of the script uses telnet to run the backup, instead of SNMP. | VBS |
Backup Event Log and Keep Achieves | Use Case: On a recurring schedule, create a Backup of the Windows Event Log saved to the local machine.  Keep all created Windows Event Log Backups on the local machine for archiving purposes. | AMP |
Backup Exec Config Script | Configures Backup Exec for monitoring by N-central.| VBS |
CCleaner | | VBS |
Change Agent ApplianceID | This VB script will change the ApplianceID of a Windows Agent. Written by Tim Wiser.| VBS |
Change Agent Customer ID | Use this script to move agents to a new customer.| VBS |
Change Local User Password | This script will allow you to change the password on a local user account.| VBS |
Change Windows Probe Password | 
This script allows you to change the username and password being used by the Windows Probe. See the README for command-line examples. | VBS |
Check For Windows Updates | | VBS |
Clean Temporary Files from All Profiles | This script cleans temp files from all user profiles, as well as the system temp folder. It only uses system environment variables and registry reads so it should run under any user account. | VBS |
Clean Windows Update Folder | AM1.5 - Stop Windows Update Service, rename C:\Windows\SoftwareDistribution\Download to OLD, restart Windows Update Service and delete C:\Windows\SoftwareDistribution\Download_OLD | AMP |
Cleanup Windows Update | Run this EXE as a Self-Healing task on the Patch Status service - it will resolve many of the reasons why the Patch Status service sometimes transitions to a Misconfigured state. | |
Clear Print Spool | Use Case: Corrupt .shd and .spl files can stop the Print Spooler Service which stops all print jobs.  Before the Print Spooler Service can be successfully restarted, the corrupt .shd and .spl files must be deleted from the Print Spooler directory. | AMP |
Collect Network Adapter Statistics | AM1.5 - Collect Network Statistics | AMP |
Configure Windows SNMP | Configures the SNMP settings on a Windows device. This script adds 127.0.0.1 as a valid SNMP host, and configures an SNMP community string, based on what you give it as a command line parameter. It requires the SNMP service to already be installed. | VBS |
Defrag | This script will defragment the hard drive specified as a command line parameter.| Batch |
Delete Temporary Windows Files | Deletes all files present in the %temp% directory. | VBS |
Delete-ultra-vnc-from-start-menu-batch-scripts | | Batch |
Deploy LiveVault Agents | Provides instructions on how to deploy Iron Mountain's LiveVault agents en masse, and provides a script that can be used to register the LiveVault agents with the provisioning server. | VBS |
Detailed Mailbox Information | AM1.5 - For a specific email address, retrieve Email Address Owner, Mailbox Quota, Mailbox Folder Statisitcs, Mailbox Forwarding Information and Mailbox Permissions. | AMP |
Download and Run File from URL | Download file from URL and run executable on device. | AMP |
Enable and Disable UAC | A reboot is most likely going to be required before the UAC changes take effect. | VBS |
Enable and Reset Firewall | This enables the firewall if disabled, and then sets all default values back to normal. this is useful to fix network issues due to firewall rules | AMP |
EnableDisable Windows Firewall | AM1.5 - Enable or Disable Windows Firewall on device.| AMP |
Exchange 2007 - Start All Services | This script will start all of the Exchange 2007 windows services. | Batch |
Exchange 2007 - Stop All Services | This script will stop all of the Exchange 2007 windows services. | Batch |
Find Exchange Settings | Script that searches an entire machine looking for EDB and STM files, lists them all and their location, scans the registry and outputs all the results in an email. Update the email and SMTPServer addresses at the start. Thanks to Ben Walton! | PowerShell |
Find External IP | This script requires internet access on the machine it's being run against. | VBS |
Find Mailboxes Larger than | AM1.5 - Query for mailboxes larger than a specified size| AMP |
Generic App Log File Maintenance | AM1.5 - Generic Application Log File  Maintenance described in Automating your Microsoft World: Application Efficiencies webinar. | AMP |
Get AD Users | | Batch |
Get Domain Information | This Policy returns the forest information and domain information into the result of the policy | AMP |
Get File from URL | This | AMP |   file will get a file from the configured URL. Created by Wim Lamot.| AMP |
Import Registry File | AM1.5 - Import Registry File into device.  The Registry File must exist on the device.  Leveraging N-central's File Transfer and this amp within an N-central Scheduled Task Profile could be used to copy the file to the device and execute amp.| AMP |
Install PowerShell 3 | From Tyler at iSeek | PowerShell	|
install_net45_ps40 |  .NET 4.5 Install, Powershell 4.0 (WMF 4) Install Thanks to Stephen Testino @ S7 Technology Group | PowerShell	|
Intel Chipset Upgrade Script | Package to identify upgradable PCs with the Intel® Q45 or B43 Express Chipset and perform the upgrade to Level I or Level III Manageability. See http://upgrades.intel.com for more information. | Batch |
Lock and Unlock USB | These two VB scripts will lock (or unlock) the ability to mount USB drives on Windows machines. | VBS |
Logoff Users-Lock Workstation | AM1.5 - Logoff all users and lock device | AMP | 
Managed Server Maintenance | Managed Server Maintenance Automation Policy | AMP | 
Managed Workstation Maint | Workstation Maintenance Automation Policy| AMP | 
Map User Specific Network Drive | Map User Specific Network Drive for Current Logged on User | AMP |
mbsa_install | This script will install MBSA on client machines.  Be sure to edit the contact information before updating it. | VBS |
McAfee Update | | Batch |
Open FTP Ports in Windows Firewall | This script will open port 21 in Windows Firewall, and will configure the Windows FTP server to use dynamic ports.| Batch |
Powershell Set Execution Policy to allow running scripts. allows runing Powershell script on a client computer | Batch |
Reboot Windows | This script creates a Windows Scheduled Task called 'RebootPC' that immediately reboots the machine. | VBS |
reconfigWMI | | Batch |
Remove AV Entries in WMI | Removes all entries in the root\SecurityCenter WMI namespace. This will help install N-central's Endpoint Security on machines that have a pre-existing A/V solution. | VBS |
Reset Windows Update | Stops the Windows Update service, renames the Software Distribution folder, and restarts the windows service. Used to help resolve Misconfigured issues with the Patch Status service. | AMP |
Reset WMI | Runs a WMI reset, and then restarts all of the WMI-related services.| VBS |
Reset_Probe_Password | This AMP will stop probe service, reset probe AD user account password and restart probe service.  Must specify AD Common Name of current user account assigned probe.  New password must be specified within Set AD User Password object of AMP. | AMP |
Restore Point Maintenance | Manage and/or Create restore points | AMP |
Run CleanMgr | AM1.5 - Run CleanMgr with all options selected. | AMP |
Run MalwareBytes | This version adds support for x64 platforms. NOTE: This script only works with the licensed version of Malware Bytes - it will not work with the free version.| VBS |
Run Sophos Scan | There are two scripts - one for 32-bit installs and another for 64-bit installs. Both use SAV32CLI to scan for and remove threats on the following operating systems. A log file is written to C:\SAVScanLog.txt on the machine. Thanks to Matt Wilcox! | Batch |
Run Trend Micro Disk Cleaner | This script downloads and runs the Trend Micro Disk Cleaner.| VBS |
Run Windows Defender Scan | | VBS |
Runbook - CleanupAndDeploy v3.3 | This ZIP contains AgentCleanup.exe and installNableAgent.bat referenced within the Managing Agent Upgrades through Group Policy procedure of the Runbook. | Batch |
Runbook - KaseyaCustomerListMigration | |
Runbook - KaseyaMigrationandRemoval | KaseyaMigrationandRemoval | |
RunSpyBot-Split | | VBS |
Scripting Examples | This zip file contains Visual Basic Script (VBS) examples that you can use to create your own scripts - neat! | VBS |
Security Manager Patch Removal | This script will remove Security Manager Patch from your system.| AMP |
Server Maintenance | AM1.5 - Server Maintenance described in Automating your Microsoft World: Operating System Efficiencies webinar. | AMP |
Set DNS Scavenging | AM1.5 - Enable and set DNS Scavenging on the DNS Server as well as start Scavenging now.| AMP |
Set Power Plan and Screensaver | Allow to configure a screen-saver and power plan, and force a password on getting out of screen-saver if desired| AMP |
Set Power Plan | AM1.5 - Specify device Power Plan to; High Performance, Balanced or Power Saver| AMP | 
SQL Optimization | AM1.5 - SQL Diagnostics and Maintenance described in Automating your Microsoft World: Application Efficiencies webinar.| AMP |
Start Blackberry Services | This script will start all of the BES (Blackberry Enterprise Server) windows services.| Batch |
StaticDNSServer | | VBS |
Stop Blackberry Services | This script will stop all of the BES (Blackberry Enterprise Server) windows services.| Batch |
Stop-Run-Start the Windows Agent | Written by Will Smith. This script will stop/start/restart the Windows Agent, and does so in such a way that the Scheduled Task in N-central will report back the status.| VBS |	
Sync Time | AM1.5 - If ran on a member server, sync time with Domain Controller.  If ran on Domain Controller, sync time with external NTP server.| AMP |
Uninstall Connect2Help | Removes the Connect2Help agent. Yes
Uninstall Google Toolbar (Internet Explorer ONLY) |  This script will uninstall the Google Toolbar for Internet Explorer. The script does not work on Google Toolbar for Firefox.| VBS |
Uninstall iTunes and associated services | | VBS |
Uninstall the 'Ask Jeeves' Toolbar | Uninstall the Internet Explorer version of the 'Ask Jeeves' Toolbar.| VBS |
Uninstall Yahoo Toolbar (for Internet Explorer) | Uninstalls the Yahoo! Toolbar for Internet Explorer. The 'Yahoo Install Manager' application is left behind in Add/Remove Programs, but will be removed on next reboot.| VBS |
Unlock AD User | This script will unlock all locked-out AD user accounts. The script accepts the domain name as a command-line parameter: cscript UnlockUser.vbs MYDOMAIN | VBS |
Update Windows Defender Defs | | VBS |
Upload Agent Logs to N-able Support | The input parameter is suggested to be the case number (such as CAS-XXXXX-XXXXXX).  A big thanks to Mr Jon Czerwinski! | AMP |
Vipre - Quick or Full Scans | These two scripts will get Vipre A/V to auto-update, and then run either a full scan or a quick scan.| VBS |
Windows Disk Cleaner | Runs Windows Disk Cleaner. By default, all cleaning-related options are enabled, so make sure to review what's getting cleaned out before you run the script!| VBS |
Windows Update - Search Download and Install | This script will Search, Download and Install windows updates. Thanks to Steve Drees!| VBS |
Workstation Maintenance | AM1.5 - Workstation Maintenance described in Automating your Microsoft World: Operating System Efficiencies webinar.| AMP |
WSUS - Delete Existing Settings | Deletes the current WSUS settings on a device.| Batch |
WSUS - Display WSUS Settings in Notepad | This script will query the device for it's WSUS settings, and will display them in Notepad. This script should not be run through N-central - it should be run on the desktop of the device.| Batch |
WSUS - Report to WSUS Server | Causes the targeted devices to report in to the WSUS server. | Batch |
WSUS Config Script | When running this script in N-central, you need to specify a target group and the URL of your WSUS server in the "Command Line Parameters" field. | VBS |