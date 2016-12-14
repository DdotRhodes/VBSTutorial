'====================7.1=====================
'
' Test Scripts by David Rhodes 12/05/16
'
'============================================

' WshShell

' Set objShell = CreateObject("Wscript.shell")

Dim WshShell, BtnCode
Set WshShell = WScript.CreateObject("WScript.Shell")
'intButton = object.popup(strText,[nSecondsToWait],[strTitle,[nType])
BtnCode = WshShell.Popup("Is your name Bob?", 7, "Answer This QUestion:", 4 + 32) ' the 7 is the  number of seconds to wait - 4 & 32 refer to bottom chart
Select Case BtnCode
	Case 6 	WScript.Echo "Hello Bob."
	Case 7 	WScript.Echo "Oh, I must have you confused with someone else."
	Case -1 WScript.Echo "Hello! Are you there?"
End Select

' 0 = Ok							16 = Stop					1 = Ok
' 1 = Ok and Cancel					32 = Question Mark			2 = Cancel
' 2 = Abort, Retry and Ignore		48 = Exclimation Mark		3 = Abort
' 3 = Yes, No and Cancel			64 = Information			4 = Retry
' 4 = Yes and No												5 = Ignore
' 5 = Retry and Cancel											6 = Yes
'																7 = No
'																-1 = Nothing
