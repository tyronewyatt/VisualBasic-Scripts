'    ************************************************************************************************************ 
'    Purpose:    Get Windows 10 Version for BgInfo
'    Pre-Reqs:   Windows 10
'    Modified by Tyrone Wyatt
'    Last Modified: 19/07/2021
'    ************************************************************************************************************ 

On Error Resume Next

Set WshShell = CreateObject("WScript.Shell")

' Windows 10 versions 2004 or older
ReleaseId = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ReleaseId")

' Windows 10 version 20H2 or newer
DisplayVersion = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DisplayVersion") 

' If DisplayVersion Exists Then Echo DisplayVersion Else Echo ReleaseId
If (DisplayVersion) Then
    'WScript.Echo DisplayVersion
    Echo DisplayVersion
Else 
    'WScript.Echo ReleaseId
    Echo ReleaseId
End If
