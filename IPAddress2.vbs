'    ************************************************************************************************************ 
'    Purpose:    Get IPv4 Addresses for BgInfo
'    Modified by Tyrone Wyatt
'    Last Modified: 24/07/2021
'    ************************************************************************************************************ 

On Error Resume Next

Set objWMIService = GetObject("winmgmts:\root\cimv2")

Set IPConfigSet = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
 
For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then 
            'WScript.Echo IPConfig.IPAddress(i)
            Echo IPConfig.IPAddress(i)
    End If
Next