Option Explicit
'Keisen Macro v1.0 for Sakura Editor
Call DrawLine(4)

Sub DrawLine(mode) 'mode 1:Up, 2:Down, 3:Left, 4:Right
	Dim objShell, iWidth, sInfo, sKeisen, i, iInfo1(31), iInfo2(3, 32)
	iWidth = 1
	On Error Resume Next
	Set objShell = CreateObject("WScript.Shell")
	iWidth = objShell.RegRead("HKCU\Software\Sakura\Keisen\LineWidth")
	On Error Goto 0
	sInfo =	"01001110111020022202222021110122" & _
			"01110011101022200222022120112102" & _
			"10011001111200220022220212201211" & _
			"10100111011202002220221202221011"
	sKeisen = ""
	For i = 1 To 32
		sKeisen = sKeisen + Chr(&H849E + i)
		iInfo2(1, i) = CInt(Mid(sInfo, i, 1))
		iInfo2(0, i) = CInt(Mid(sInfo, i + 32, 1))
		iInfo2(3, i) = CInt(Mid(sInfo, i + 64, 1))
		iInfo2(2, i) = CInt(Mid(sInfo, i + 96, 1))
		iInfo1(i - 1) = iInfo2(2, i) * 64 + iInfo2(3, i) * 16 + _
						iInfo2(0, i) * 4 + iInfo2(1, i)
	Next
	Call InsertText(mode, iWidth, sKeisen, iInfo1, iInfo2, 1)
	If MoveCur(mode) = 0 Then Exit Sub
	Call InsertText(mode, iWidth, sKeisen, iInfo1, iInfo2, 0)
End Sub

Sub InsertText(mode, iWidth, sKeisen, iInfo1, iInfo2, rr)
	Dim i, n, text1, iWidth1, iWidth2(1, 4)
	text1 = Array("", "", "")
	iWidth1 = Array(0, 0, 0, 0, 5, 5, 80, 80, 10, 10, 160, 160)
	n = 1
	For i = 1 To 4
		If rr <> 0 And i = mode Then
			iWidth1(2) = iWidth
		Else
			iInfo2(i - 1, 0) = 0
			iWidth1(2) = iInfo2(i - 1, _
				InStr(sKeisen, Left(GetText(i, rr) & " ", 1)))
		End If
		iWidth1(3) = iWidth1(2)
		If iWidth1(3) > 0 Then iWidth1(3) = iWidth
		iWidth1(0) = iWidth1(0) + iWidth1(2) * n
		iWidth1(1) = iWidth1(1) + iWidth1(3) * n
		n = n * 4
	Next
	n = iWidth * 4 + mode - 1
	For i = 0 To 31
		If iWidth1(0) = iInfo1(i) Then text1(0) = Chr(&H849F + i)
		If iWidth1(1) = iInfo1(i) Then text1(1) = Chr(&H849F + i)
		If iWidth1(n) = iInfo1(i) Then text1(2) = Chr(&H849F + i)
	Next
	n = Asc(GetText(0, 0) & " ")
	If n = Asc(" ") Or n >= &H849F And n <= &H84BE Then
		Call BeginSelect
		Call MoveCur(4)	 ' Right
	End If
	Call InsText(Left(text1(0) & text1(1) & text1(2), 1))
	Call Editor.Left
End Sub

Function GetEditor(mode, iStrIndex, iByteIndex, y)
	Dim i, text1(1), n, x
	text1(0) = GetLineStr(y)
	x = 1
	For i = 1 To Len(text1(0))
		If iStrIndex > 0 And i >= iStrIndex Then Exit For
		text1(1) = Mid(text1(0), i, 1)
		If iByteIndex > 0 And iByteIndex <= x Then
			GetEditor = text1(1)
			Exit Function
		End If
		n = Asc(text1(1))
		If n = 13 Or n = 10 Then Exit For
		x = x - (mode = 1 And (n > 255 Or n < 0)) + 1
	Next
	GetEditor = x
End Function

Function GetText(mode, rr) ' rr : Right Right (x + 2 position)
	GetText = " "
	Dim x, y, text
	x = CInt(ExpandParameter("$x"))
	y = CInt(ExpandParameter("$y"))
	Select Case mode
	Case 0 ' Cur
		GetText = ""
			 x = GetEditor(1, x, 0, 0)
		If x >= GetEditor(1, 0, 0, 0) Then Exit Function
		GetText = GetEditor(1, 0, x, 0)
	Case 1 ' Up
		If y = 1 Then Exit Function
		x = GetEditor(1, x, 0, 0)
		If GetEditor(1, 0, 0, y - 1) < x Then Exit Function
		GetText = GetEditor(1, 0, x, y - 1)
	Case 2 ' Down
		If y >= GetLineCount(0) Then Exit Function
		x = GetEditor(1, x, 0, 0)
		If GetEditor(1, 0, 0, y + 1) <= x Then Exit Function
		GetText = GetEditor(1, 0, x, y + 1)
	Case 3 ' Left
		If x <= 1 Then Exit Function
		x = GetEditor(1, x - 1, 0, 0)
		GetText = GetEditor(1, 0, x, 0)
	Case 4 ' Right
		   x = GetEditor(1, x + 1, 0, 0)
		If x >= GetEditor(1, 0, 0, 0) Then Exit Function
		text = GetEditor(1, 0, x, 0)
		GetText = text
		If rr = 0 Or text <> " " Then Exit Function
		x = x + 1
		If x >= GetEditor(1, 0, 0, 0) Then Exit Function
		GetText = GetEditor(1, 0, x, 0)
	End Select
End Function

Function MoveCur(mode)
	Dim x, y, iSpace, text
	y = CInt(ExpandParameter("$y"))
	x = LineIndexToColumn(y, CInt(ExpandParameter("$x")))
	iSpace = 0
	Select Case mode
	Case 1 ' Up
		MoveCur = 0
		If y = 1 Then Exit Function
		iSpace = x - LineIndexToColumn(y - 1, _
			GetEditor(0, 0, 0, y - 1))
		Call Editor.Up
		MoveCur = 1
	Case 2 ' Down
		MoveCur = 1
		If y = GetLineCount(0) Then
			Call GoLineEnd
			Call Char(13)
			iSpace = x - LineIndexToColumn(y, _
				CInt(ExpandParameter("$x")))
		Else
			iSpace = x - LineIndexToColumn(y + 1, _
				GetEditor(0, 0, 0, y + 1))
			Call Editor.Down
		End If
	Case 3 ' Left
		MoveCur = 0
		If x = 1 Then Exit Function
		Call Editor.Left
		If GetText(0, 0) = " " And _
			CInt(ExpandParameter("$x")) > 1 Then
			Call Editor.Left
			If GetText(0, 0) <> " " Then Editor.Right
		End If
		MoveCur = 1
	Case 4 ' Right
		MoveCur = 1
		text = GetText(0, 0)
		If text = "" Then Exit Function
		Call Editor.Right
		If text = " " And GetText(0, 0) = " " Then Editor.Right
	End Select
	If iSpace <= 0 Then Exit Function
	Call GoLineEnd
	Call InsText(Space(iSpace))
End Function

