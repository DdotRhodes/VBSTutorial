'====================6.5=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

' Searching for text in a file

Dim objFSO, objFile, LogFile, Search, FoundText
LogFile = "C:\Users\Llama\Desktop\LogFile.txt"
Const ForReading = 1 ' 1 means read only

Set objFSO = CreateObject ("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(LogFile, ForReading)

Search = objFile.ReadAll
objFile.Close
FoundText = InStr(Search, "Old")
WScript.Echo FoundText

' If script echos back 0 that means string couldnt be found
' If script echos back 1 that means the text was found
