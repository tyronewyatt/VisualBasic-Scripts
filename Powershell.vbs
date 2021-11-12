'Set objShell = CreateObject("Wscript.shell")

'objShell.run("powershell -noexit -file c:\fso\CleanupFiles.ps1")

'objShell.run("powershell -command {Write-Out 'Hello World!'}")



' CallPowerShell.vbs
' Call PowerShell from VBScript
' Author: ITomation (http://itomation.ca)
' Version 1.0 - 2015-11-27
' --------------------------------------------'

'On Error Resume Next

Option Explicit
Dim strPSCommand
Dim strDOSCommand
Dim objShell
Dim objExec
Dim strPSResults

' Construct PowerShell Command (PS syntax)
strPSCommand = "write-host 'Hello World!'"

' Consruct DOS command to pass PowerShell command (DOS syntax)
strDOSCommand = "cmd /c start /min powershell -windowstyle hidden -command " & strPSCommand & ""

' Create shell object
Set objShell = CreateObject("Wscript.Shell")

' Execute the combined command
Set objExec = objShell.Exec(strDOSCommand)

' Read output into VBS variable
strPSResults = objExec.StdOut.ReadAll

' Echo results
Echo strPSResults