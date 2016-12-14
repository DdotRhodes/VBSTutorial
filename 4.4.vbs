'====================4.4=====================
'
' Test Scripts by David Rhodes 12/05/16
'
'============================================

' Making choices based on inputs

Dim Hostname
Dim Confirm


Hostname = InputBox ("What is the Hostname of the Server?")

Confirm = MsgBox ("You typed in " & Hostname & vbCrLf & "Is this correct?", vbYesNo, "Continue?")

If Confirm = vbYes Then
	MsgBox ("The Yes Button was Clicked.")
Else MsgBox ("The No Bustton was Clicked.")
End If 
