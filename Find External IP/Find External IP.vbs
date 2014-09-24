' ******************************************************************************
' Script: GetExternalIP.vbs
' Version: 1.0
' Author: Chris Reid
' Description: This script will find the public/external IP address of the machine
'              run against. 
' Date: November 17th, 2011
' ******************************************************************************


' Thanks to http://jackson.io/ip/service.html for the original code!

' Create objects
set ip = createobject("Microsoft.XMLHTTP")
Set output = Wscript.stdout

' Fetch the IP address
ip.open "GET", "http://queryip.net/ip/", false
ip.send(null)
output.writeline ip.responsetext
