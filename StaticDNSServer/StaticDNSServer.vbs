On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colNetCards = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objNetCard in colNetCards
    arrDNSServers = Array("192.168.1.100", "192.168.1.200")
    objNetCard.SetDNSServerSearchOrder(arrDNSServers)
Next