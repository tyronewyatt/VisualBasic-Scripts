'    ************************************************************************************************************ 
'    Purpose:    Get Windows Build
'    Pre-Reqs:   Windows 10
'	 Modified by Tyrone Wyatt
'	 Last Modified: 14/06/2022
'    ************************************************************************************************************ 

On Error Resume Next

Set WshShell = CreateObject("WScript.Shell")

' Windows Edition
Set objWMIService = GetObject("winmgmts:\root\cimv2")
Set OperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each OperatingSystem in OperatingSystems
    If (OperatingSystem.Caption) Then
        OperatingSystemEdition = OperatingSystem.Caption
    End If
Next

' Windows Build
CurrentBuild = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuild")

' End Of Support Windows 10 version 1709
If (OperatingSystemEdition = "Microsoft Windows 10 Enterprise" And CurrentBuild = "16299" And Now() > #10/13/2020#) Then
    EndOfSupport = "Your Windows version has reached end of service" + vbCrLf + "Take your PC to IT Service Desk for help upgrading"

' End Of Support Windows 10 version 1803
ElseIf (OperatingSystemEdition = "Microsoft Windows 10 Enterprise" And CurrentBuild = "17134" And Now() > #05/11/2021#) Then
    EndOfSupport = "Your Windows version has reached end of service" + vbCrLf + "Take your PC to IT Service Desk for help upgrading"

' End Of Support Windows 10 version 1909
ElseIf (OperatingSystemEdition = "Microsoft Windows 10 Enterprise" And CurrentBuild = "18363" And Now() > #05/10/2022#) Then
    EndOfSupport = "Your Windows version has reached end of service" + vbCrLf + "Take your PC to IT Service Desk for help upgrading"

' End Of Support Windows 10 version 21H2
ElseIf (OperatingSystemEdition = "Microsoft Windows 10 Enterprise" And CurrentBuild = "19044" And Now() > #06/11/2024#) Then
    EndOfSupport = "Your Windows version has reached end of service" + vbCrLf + "Take your PC to IT Service Desk for help upgrading"

End If

If (EndOfSupport) Then
    WScript.Echo EndOfSupport
    Echo EndOfSupport
End If