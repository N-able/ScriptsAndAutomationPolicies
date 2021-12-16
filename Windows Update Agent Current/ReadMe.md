Windows Update Agent Current
============================

Overview
--------
This is the AMP to feed an AMP-based custom service.

It returns two parameters:
* Windows Update Agent is Current - True or False - whether the windows update agent on the device is at least at the level defined in the automation policy
* Windows Update Engine Version - The version of wuaueng.dll

Maintenance
-----------

The required agent versions are hard-coded into the policy.  This is because (as of Automation Manager 2.0.3.177) there is a bug where changes to Global Variable value in the Input object are not updated in the executed policy.  As new minimum installed versions are released, update the appropriate assignments in the policy.  Note that these may be different for Windows 7 and Windows 8.1.


