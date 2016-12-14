'====================3.2=====================
'
' Test Scripts by David Rhodes 12/05/16
'
'============================================

Dim arrShopping()'keep blank to not limit size of array
ReDim arrShopping(9)
	arrShopping(0) = "Bread"
	arrShopping(1) = "Milk"
	arrShopping(2) = "Apples"
	arrShopping(3) = "Carrots"
	arrShopping(4) = "Cereal"
	arrShopping(5) = "Oranges"
	arrShopping(6) = "Bananas"
	arrShopping(7) = "Cheese"
	arrShopping(8) = "Lemons"
	arrShopping(9) = "Tomatoes"
ReDim Preserve arrShopping(10) 'preserve tells VB to keep what was in the array and now add one
arrShopping(10) = "Coffee"
MsgBox arrShopping(9)
MsgBox arrShopping(10)