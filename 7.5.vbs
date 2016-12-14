'====================7.5=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

' Writing to the registry and writting to the log

Option Explicit
Dim objShell
Dim strHostname, strNewEntry, strNewValue, strCheckEntry

strNewEntry = "WinstructorTest"

strHostname = "HKEY_LOCAL_MACHINE\SYSTEM\"_
& "CurrentControlSet\Services\Tcpip\Parameters\"

Set objShell = CreateObject("WScript.Shell")

strNewEntry = objShell.RegWrite(strHostname & strNewEntry, "I am in the registry")

strCheckEntry = objShell.RegRead(strHostname & strNewEntry)

objShell.LogEvent 0, "The location: " & strHostname & strNewEntry & _
" contains: " & strCheckEntry

' When logging event
' 0 means successful event
' 1 means error
' 2 means warning
' 4 for informational event
' 8 for audit success
' 16 for audit failure
