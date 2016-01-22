GetScriptOutputAuto.ps1
=======================

A basic script to automate the collection and collation of N-central script output.

<b>Use at your own risk</b>

<b>Requirements:</b>

AE.Net.Mail.dll assembly (see https://github.com/andyedinborough/aenetmail)

<b>Usage:</b>

1. Created a dedicated mailbox for N-central scripting results
2. Create a new recipient in N-central with an email address pointing to this mailbox
3. Set $outputDir, $mailServer, $mailPort, $mailUser and $mailPass accordingly
4. Create a scheduled task to execute .\GetScriptOutputAuto.ps1 periodically
5. Any scripts you wish to collect, add the N-central scripting mailbox as a recipient for the task output.  Give the task name an appropriate description

That is it.  Wait for the results.  If you perform the above, and wrap the task in a scheduled task profile, effectively you will end up with a directory in $outputDir for each customer.  Each customer directory will contain a folder for the Task Name, and any execution attempts for that script will be output to text file named by the date of execution.  Execution across multiple servers will be collated into the one file.