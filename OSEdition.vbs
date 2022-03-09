'    ************************************************************************************************************ 
'    Purpose:    Get OS Edition
'    Modified by Tyrone Wyatt
'    Last Modified: 25/01/2022
'    ************************************************************************************************************

On Error Resume Next

Set objWMIService = GetObject("winmgmts:\root\cimv2")

Set OperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

For Each OperatingSystem in OperatingSystems
    If Not IsNull(OperatingSystem.Caption) Then
        If InStr(OperatingSystem.Caption,"Microsoft") = 1 Then
            OperatingSystem.Caption = Trim(Replace(OperatingSystem.Caption,"Microsoft",""))
        End If
        'WScript.Echo OperatingSystem.Caption
        Echo OperatingSystem.Caption
    End If
Next