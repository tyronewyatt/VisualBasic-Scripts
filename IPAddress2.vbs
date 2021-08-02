'    ************************************************************************************************************ 
'    Purpose:    Get IPv4 Addresses for BgInfo
'    Modified by Tyrone Wyatt
'    Last Modified: 30/07/2021
'    ************************************************************************************************************ 

On Error Resume Next

Set objWMIService = GetObject("winmgmts:\root\cimv2")

Set IPConfigSet = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then 
            If Not Instr(IPConfig.IPAddress(i), ":") > 0 Then
            'WScript.Echo IPConfig.IPAddress(i)
            Echo IPConfig.IPAddress(i)
            End If
    End If
Next