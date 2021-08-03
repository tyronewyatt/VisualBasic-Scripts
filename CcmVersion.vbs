'    ************************************************************************************************************ 
'    Purpose:    Get CCM Client Version
'    Modified by Tyrone Wyatt
'    Last Modified: 30/07/2021
'    ************************************************************************************************************ 

On Error Resume Next

Set objWMIService = GetObject("winmgmts:\root\ccm")

Set CcmClients = objWMIService.ExecQuery ("Select ClientVersion from SMS_Client")

For Each CcmClient in CcmClients
    If Not IsNull(CcmClient.ClientVersion) Then 
        WScript.Echo CcmClient.ClientVersion
        Echo CcmClient.ClientVersion
    End If
Next