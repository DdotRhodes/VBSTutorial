'====================7.4=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

' Deleting from the registry

Option Explicit
Dim objShell
Dim strHostname, strNewEntry

strNewEntry = "WinstructorTest"

strHostname = "HKEY_LOCAL_MACHINE\SYSTEM\"_
& "CurrentControlSet\Services\Tcpip\Parameters\"

Set objShell = CreateObject("WScript.Shell")

strNewEntry = objShell.RegDelete(strHostname & strNewEntry)