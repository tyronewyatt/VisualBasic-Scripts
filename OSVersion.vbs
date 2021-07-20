'    ************************************************************************************************************ 
'    Purpose:    Get Windows 10 Version
'    Pre-Reqs:   Windows 10
'	 Modified by Tyrone Wyatt
'	 Last Modified: 19/07/2021
'    ************************************************************************************************************ 

Set WshShell = CreateObject("WScript.Shell")

' Windows 10 versions 2004 or older
ReleaseId = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ReleaseId")

On Error Resume Next

' Windows 10 version 20H2 or newer
DisplayVersion = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DisplayVersion")

' If Error On RegRead Then Echo ReleaseId Else Echo DisplayVersion
If Err.Number <> 0 Then
    Echo ReleaseId
Else 
    Echo DisplayVersion
End If