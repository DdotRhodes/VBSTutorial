'====================2.1=====================
'
' Test Scripts by David Rhodes 12/05/16
'
'============================================

Option Explicit 'Tells the program that every line in the code must be defined
On Error Resume Next 'Not required, but good to use in some cases

'Start of Header

'Declare variables (DIM means dimention)
Dim objFSO 'Naming based on what it is, i.e. object
Dim ComputerName, IPAddress 'declare multiple variables on a line using a coma
Dim literalNumeric
Dim literalString
Const HOSTNAME = "Server01" 'Allows to set a value that will not change throughout the script
Const MY_NUMBER = 33 'Common practice to put these in upper case

'End Header

'Start of Worker Information
literalNumeric = 20 
literalString = "BeachBall"
MsgBox literalNumeric
MsgBox literalString