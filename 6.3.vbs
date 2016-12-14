'====================6.3=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

'Changing the name of a folder

Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.MoveFolder "C:\Users\Llama\Desktop\Renamed Folder", "C:\Users\Llama\Desktop\Test"