'====================7.2=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

' Reading from the registry

Dim objShell, strDomainName, strDefaultUser, strServerName, strWindLogon, strHostname

strDefaultUser = "DefaultUserName"
strDomainName = "DefaultDomainName"
strServerName = "Hostname"

strWinLogon = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
strHostname = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\parameters\"

Set objShell = CreateObject("WScript.Shell")

strDomainName = objShell.RegRead(strWinLogon & strDomainName)
strServerName = objShell.RegRead(strHostname & strServerName)
strDefaultUser = objShell.RegRead(strWinLogon & strDefaultUser)

WScript.Echo "Domain Name: " & vbTab & strDomainName & vbCrLf _
& "Hostname: " & vbTab & strServerName & vbCrLf _
& "User Name: " & vbTab & strDefaultUser

WScript.Quit