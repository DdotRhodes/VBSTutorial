'====================6.2=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================


' Copies files from one folder to another

' Make sure destination paths have a trailing backslash! Without it, there will be a permissions error
' Note on permission denied error

' From Stackoverflow: 
' I've only seen CopyFile fail with a permission denied error in one of these 3 scenarios:
' An actual permission problem with either source or destination.
' Destination path is a folder, but does not have a trailing backslash.
' Source file is locked by an application.


Dim objFSO, objFolder, objFile, objNewFolder
objNewFolder = "C:\Users\Llama\Desktop\Test\"

'Create FilesSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the folder we want to copy files from. C:\Users\Llama\Desktop\Learning VB
Set objFolder = objFSO.GetFolder("C:\Users\Llama\Desktop\Learning VB\")

'Check to see if C:\Users\Llama\Desktop\Test exists, if not, create it.
If objFSO.FolderExists(objNewFolder) Then

WScript.Echo "The Destination Folder " & objNewFolder & " already exists"

Else

WScript.Echo "The Destination Folder " &  objNewFolder & " will be created"
'Create new folder called C:\Users\Llama\Desktop\Test\
Set objNewFolder = objFSO.CreateFolder (objNewFolder)
End If

For Each objFile In objFolder.Files
objFile.copy "C:\Users\Llama\Desktop\Test\"
Next