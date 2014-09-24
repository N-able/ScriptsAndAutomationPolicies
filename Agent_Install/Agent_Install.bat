'###########################################
'# Written by Mark Mays - KKI Technologies #
'# Free to use for N-Able Parnters         #
'###########################################



' Batch File Name: Agent Install
' Description: This batch file will install the agent with the following parameters:
'
'		net use "Drive Letter" - this is mapping a drive letter where you have the agent copied
'               locally on the customer network
'		m="Drive Letter":\Filename.txt - Points to the location where you have the text file created (see 2nd attachment)
' 		The last line delets the mapped drive


net use o: \\server\sharename
o:\windowsagentsetup.exe /s /m=o:\devicename.txt
net use o: /delete