###############################################################
#       Agent Download - A modified Cleanup and Deploy        #
#    Created by" Evan Morrissey with CommPutercations Inc     #
#              Last Updated: 2013-06-05                       #
###############################################################

Refer to N-Able's "Using the Agent Clean Up Utility" document for general setup of the Clean Up and Deploy. This modifies that setup to automate updating the Agent installer and script parameters when the N-Central server is updated.

To modify your existing Cleanup and Deploy setup:

- Edit agentparam.bat to include the correct server address, customerID and minVersion for your N-Central server and the customer site.
- Copy agentparam.bat and the provided installNableAgent.bat to the share you are using for your Cleanup and Deploy.
- Open Group Policy and edit the Group Policy Object for Cleanup and Deploy
- Navigate to Computer Configuratino > Windows Settings > Scripts (Startup/Shutdown) > Startup
- Edit the Cleanup and Deploy startup script so the ONLY parameter passed to the script is the share name
  - e.g. myservername\netlogon or mydomain.local\netlogon. I recommend you use mydomain.tld\netlogon if this customer has multiple domain controlers at branch offices as the PCs will get the files from the "nearest" DC.
- Upload "Agent Download.amp" to your N-Central and create a task to run this Automation Policy. This can be recurring, or can be run manually after an update to your N-Central server.
- AMP takes 3 parameters:
  - Version: enter the current version of your N-Central. Note that this is also used to download the Agent from your server, so if it is not accurate the policy will not work properly.
  - Server: your N-Central server address
  - Temp Local Path: A folder on the server, this is just to download the Agent to before copying it to the netlogon folder. This is just done to avoid unsightly permissions problems.


When you update your N-Central server you will just need to modify the Version parameter and re-run the policy. When devices reboot they should be installing the latest agent!