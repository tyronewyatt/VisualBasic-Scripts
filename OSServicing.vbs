'    ************************************************************************************************************ 
'    Purpose:    Get Windows Build
'    Pre-Reqs:   Windows 10
'	 Modified by Tyrone Wyatt
'	 Last Modified: 31/03/2022
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
If (OperatingSystemEdition = "Microsoft Windows 10 Enterprise" And CurrentBuild = "16299") Then
    EndOfSupport = "Your Windows version reached end of service on 13/10/2020" + vbCrLf + "Take your PC to IT Service Desk for help upgrading"

' End Of Support Windows 10 version 1803
ElseIf (OperatingSystemEdition = "Microsoft Windows 10 Enterprise" And CurrentBuild = "17134") Then
    EndOfSupport = "Your Windows version reached end of service on 11/05/2021" + vbCrLf + "Take your PC to IT Service Desk for help upgrading"

' End Of Support Windows 10 version 1909
ElseIf (OperatingSystemEdition = "Microsoft Windows 10 Enterprise" And CurrentBuild = "18363") Then
    EndOfSupport = "Your Windows version reaches end of service on 10/05/2022" + vbCrLf + "Go to Software Center to upgrade your Operating System"

' End Of Support Windows 10 version 21H2
ElseIf (OperatingSystemEdition = "Microsoft Windows 10 Enterprise" And CurrentBuild = "19044") Then
'    EndOfSupport = "Your Windows version reaches end of service on 11/06/2024" + vbCrLf + "Go to Software Center to upgrade your Operating System"

End If

If (EndOfSupport) Then
    WScript.Echo EndOfSupport
    Echo EndOfSupport
End If