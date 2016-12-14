'====================5.2=====================
'
' Test Scripts by David Rhodes 12/07/16
'
'============================================

' Loops

' For Each...Next
' For...Next
' Do While
' Do Until

'For...Next
'Define a variable that will keep track of the loops execution (ie counts it)

Dim myLoop
For myLoop = 1 To 20 Step 2 'makes it count by 2
WScript.Echo myLoop 'WSctipt.Echo will write out the result to the command prompt rather than a message box
Next