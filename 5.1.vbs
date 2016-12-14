'====================5.1=====================
'
' Test Scripts by David Rhodes 12/07/16
'
'============================================

' Loops

' For Each...Next
' For...Next
' Do While
' Do Until


'For Each...Next

Dim arrShopping
Dim str
arrShopping = Array("Bread", "Milk", "Apples", "Carrots", "Cereal")
For Each str In arrShopping 'For each allows you to perform an action on an object then perform that same action to another object until there are no more objects
MsgBox str
Next