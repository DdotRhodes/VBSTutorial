'====================6.7=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

' Find and replace text in file

Dim objFSO, objFile, LogFile, Search, FoundText
LogFile = "C:\Users\Llama\Desktop\LogFile.txt"
Const Forreading = 1
Const Forwriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(LogFile, Forreading)

Search = objFile.Readall
objFile.close
FoundText = Replace(Search, "New", "Old")

Set objFile = objFSO.OpenTextFile(LogFile, Forwriting)
objFile.WriteLine FoundText
objFile.Close