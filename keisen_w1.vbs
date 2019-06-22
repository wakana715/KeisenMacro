Option Explicit
'Keisen Macro v1.0 for Sakura Editor
Dim objShell
On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
objShell.RegWrite "HKCU\Software\Sakura\Keisen\LineWidth", "1"
On Error Goto 0
Call StatusMsg("1:Narrow", 0)
