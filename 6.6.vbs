'====================6.6=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

Const ForReading = 1
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
objRegEx.IgnoreCase = True
objRegEx.Pattern = "new"
Set objFSO = CreateObject("scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Users\Llama\Desktop\LogFile.txt", ForReading)
strSearchString = objFile.ReadAll
objFile.close

Set colMatches = objRegEx.Execute(strSearchString)
If colMatches.Count > 0 Then
	strMessage = "The following matches were found:" & vbCrLf
	For Each strMatch In colMatches
		strMessage = strMessage & strMatch.value & vbCrLf
	Next
	Else
	WScript.Echo"There were no matches"
End If
WScript.Echo strMessage