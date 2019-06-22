Option Explicit
'Keisen Macro v1.0 for Sakura Editor
Dim objShell, sInfo, iWidth
sInfo = Array("1:Narrow", "2:Wide", "HKCU\Software\Sakura\Keisen\LineWidth")
iWidth = 1
On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
iWidth = objShell.RegRead(sInfo(2))
iWidth = CInt(Mid("21", CInt("0" & iWidth), 1))
objShell.RegWrite sInfo(2), "" & iWidth
On Error Goto 0
Call StatusMsg(sInfo(iWidth - 1), 0)
