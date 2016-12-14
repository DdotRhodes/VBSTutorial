'====================3.2=====================
'
' Test Scripts by David Rhodes 12/05/16
'
'============================================

'Settings an array in one line of code
Dim arrShopping
arrShopping = Array("Bread", "Milk", "Apples", "Carrots", "Cereal", "Oranges", "Bananas", "Cheese", "Lemons", "Tomatoes")
ReDim Preserve arrShopping(10)
arrShopping (10) = "Coffee"
MsgBox arrShopping(9)
MsgBox arrShopping(10)