'====================7.3=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

' Writing to the registry

Option Explicit
Dim objShell
Dim strHostname, strNewEntry

strNewEntry = "Test1"

strHostname = "HKEY_LOCAL_MACHINE\SYSTEM\"_
& "CurrentControlSet\Services\Tcpip\Parameters\"

Set objShell = CreateObject("WScript.Shell")

strNewEntry = objShell.RegWrite(strHostname & strNewEntry, "I am in the registry")

' To change a differnt variable like DWord, you would do the following
'strNewEntry = objShell.RegWrite(strHostname & strNewEntry, 1, "REG_DWORD")