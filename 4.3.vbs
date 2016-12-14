'====================4.3=====================
'
' Test Scripts by David Rhodes 12/05/16
'
'============================================

' Making changes based on choices

Dim Choice
Choice = MsgBox("Do you want to continue?", vbYesNo, "Continue?")
If Choice = vbYes Then
	MsgBox ("The Yes Buton was Clicked.")
Else MsgBox ("The No Button was Clicked.")
End If