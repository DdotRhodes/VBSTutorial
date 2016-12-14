'====================3.1=====================
'
' Test Scripts by David Rhodes 12/05/16
'
'============================================

'Arrays

'1 dimentional array example:
'-----0-----------------------
'0. Bread
'1. Milk
'2. Apples
'3. Carrots
'4. Cereal
'5. Oranges
'6. Bananas
'7. Cheese
'8. Lemons
'9. Tomatoes

'1 dimentional array example:
'-----0----------------1------
'0. Bread			 John
'1. Milk			 Bob
'2. Apples			 Jack
'3. Carrots
'4. Cereal
'5. Oranges
'6. Bananas
'7. Cheese
'8. Lemons
'9. Tomatoes

'Declaring the Array
'Arrays always start at 0
Dim arrShopping(1,9)'1 represents how many rows (2) and the 9 represents how many columns (10)


'reference array and specify where to put the data

arrShopping(1,3) = "Alex"
arrShopping(0,3) = "Carrots"

MsgBox arrShopping(1,3) & " wants some " & arrShopping(0,3)

