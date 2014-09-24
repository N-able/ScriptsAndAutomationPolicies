'*******************************************************************************
'* Script:         RmvAVWMI.vbs
'* Purpose:        Removes any AntivirusProduct from Windows Security Center
'* Created:        2011/06/26
'* Created by:     Brian Hershey
'* Supported OS:   XP, Vista
'*******************************************************************************
Option Explicit
'* ------------------------- Global Variable Declarations ------------------------------
Dim strComputer, oAVWMI, oASWMI, oFWWMI, colAS, colFW, colAV, objAntiSpywareProduct, objFirewallProduct, objAntiVirusProduct, strASGuid, strFWGuid, strAVGuid, strAV, strFW, strAS, objAVSWbemServices, objASSWbemServices, objFWSWbemServices, strAVInstance, strASInstance, strFWInstance, Err


'============================== Main Script =============================

'--- Connect to WMI \root\SecurityCenter
strComputer = "."
Set oAVWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\SecurityCenter")
Set oASWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\SecurityCenter")
Set oFWWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\SecurityCenter")
Set colAV = oAVWMI.ExecQuery("Select * from AntiVirusProduct")
Set colAS = oASWMI.ExecQuery("Select * from AntiSpywareProduct")
Set colFW = oFWWMI.ExecQuery("Select * from FirewallProduct")
  

'--- Remove ANY AV products that are registered with Windows Security Center
    If colAV.Count > 0 Then
  
        ' --- Start loop for each AV product
        For Each objAntiVirusProduct In colAV
      
            ' -- Get the Guid
            strAVGuid = objAntiVirusProduct.instanceGuid
            Set colAV = oAVWMI.ExecQuery("Select " &  strAVGuid & " from objAntiVirusProduct")
          
          
                ' -- Setup a connection to Wbem then delete Sav Instance
                strAVInstance = "AntiVirusProduct.instanceGuid='" & strAVGuid & "'"
                Set objAVSWbemServices = GetObject("winmgmts:\\" & "." & "\root\SecurityCenter")
                objAVSWbemServices.Delete strAVInstance
          

                 ' Release SwbemServices object
                 Set objAVSWbemServices = Nothing
          


 next ' --- process next product
      

'--- Remove ANY Antispyware products that are registered with Windows Security Center
    If colAS.Count > 0 Then
  
        ' --- Start loop for each AS product
        For Each objAntiSpywareProduct In colAS
      
            ' -- Get the Guid
            strASGuid = objAntiSpywareProduct.instanceGuid
            Set colAS = oASWMI.ExecQuery("Delete " &  strASGuid & " from objAntiSpywareProduct")
          
           
                ' -- Setup a connection to Wbem then delete AS Instance
                strASInstance = "AntiSpywareProduct.instanceGuid='" & strASGuid & "'"
                Set objASSWbemServices = GetObject("winmgmts:\\" & "." & "\root\SecurityCenter")
                objASSWbemServices.Delete strASInstance

                 ' Release SwbemServices object
                 Set objASSWbemServices = Nothing
          

 next ' --- process next product

'--- Remove ANY Firewall products that are registered with Windows Security Center
    If colFW.Count > 0 Then
  
        ' --- Start loop for each Firewall product
        For Each objFirewallProduct In colFW
      
            ' -- Get the Guid
            strFWGuid = objFirewallProduct.instanceGuid
            Set colFW = oFWWMI.ExecQuery("Delete " &  strFWGuid & " from objFirewallProduct")


                ' -- Setup a connection to Wbem then delete Firewall Instance
                strFWInstance = "FirewallProduct.instanceGuid='" & strFWGuid & "'"
                Set objFWSWbemServices = GetObject("winmgmts:\\" & "." & "\root\SecurityCenter")
                objFWSWbemServices.Delete strFWInstance
               
   ' Release SwbemServices object
                 Set objFWSWbemServices = Nothing
          
 next ' --- process next product
  
     End If
     End If
     End If

 

' --- Release all resources
Set objAntiVirusProduct = Nothing
Set objAntiSpywareProduct = Nothing
Set objFirewallProduct = Nothing
Set colAV = Nothing
Set colAS = Nothing
Set colFW = Nothing
Set oAVWMI = Nothing
Set oASWMI = Nothing
Set oFWWMI = Nothing

If Err <> 0 Then
    WScript.Quit(0) ' --- Quit, no errors
Else
    WScript.Quit(1) ' --- Quit with errors
End If