EXPLENATION:
------------
This script makes partitions the max size according to their disk (ex.: after you enlarged the disk in VMWare or Hyper-V)
It only accepts 1 argument, the driveletter as a single character (only E and not E:)

REQUIREMENTS:
-------------
Windows 2012 is required, tested with Windows 2012 (so should work on 8), Windows 2012 R2 (so should work on 8.1) and Windows 10 (so should work on 2016).

UPLOADING:
----------
There are two methods to upload an Automation Policy (AMP) into N-central.  Choose one of the following methods:

How to upload AMP directly into N-central:

1. Login to the UI as Product Admin (by default, productadmin@n-able.com)
2. Choose on the OS Level from top-left drop down
3. Expand Configuration Pane -> Scheduled Tasks -> Script/Software Repository
4. Click Add -> Select Automation Policy
5. Click Browse -> Select <Automation_Policy>.AMP
6. Click OK -> OK 

How to upload AMP through Automation Designer into N-central:

1. Login to the UI as Product Admin (by default, productadmin@n-able.com)
2. Choose on the OS Level from top-left drop down
3. From Actions Pane -> Click Start Automation Manager
4. From within the Automation Manager Designer, choose File -> Open
5. Browse to <Automation_Policy>.AMP
6. Click OK
7. From within the Automation Manager Designer, choose File -> Upload

CREDITS:
--------
v1.0	-	Robby Swartenbroekx (b-inside - www.b-inside.be)
