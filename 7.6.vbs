'====================7.6=====================
'
' Test Scripts by David Rhodes 12/12/16
'
'============================================

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "calc"
WScript.Sleep 100 

WshShell.AppActivate "Calculator"
WScript.Sleep 100

WshSHell.SendKeys "5"
WScript.Sleep 500
WshSHell.SendKeys "{+}"
WScript.Sleep 500
WshSHell.SendKeys "10{+}"
WScript.Sleep 500 
WshSHell.SendKeys "12"
WScript.Sleep 1000
WshSHell.SendKeys "="
WScript.Sleep 1200
