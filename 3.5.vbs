'====================3.5=====================
'
' Test Scripts by David Rhodes 12/05/16
'
'============================================

'Converting an Array to a String

Dim arrShopping(2)
arrShopping(0) = "Bread"
arrShopping(1) = "Milk"
arrShopping(2) = "Apples"
str = Join(arrShopping," , ")
MsgBox str