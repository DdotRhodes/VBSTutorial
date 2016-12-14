'====================7.7=====================
'
' Test Scripts by David Rhodes 12/12/16
'
'============================================

Set objShell = CreateObject("WScript.Shell")
Set objWshScriptExec = objShell.Exec("ipconfig")
Set objStdOut = objWshScriptExec.StdOut
strOutput = objStdOut.ReadAll
WScript.Echo strOutput