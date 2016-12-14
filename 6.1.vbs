'====================6.1=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

' Objects
' DLL file that allows you to run a script over again
' CreatObject() - first step to running the object by loading it into memory
' ProgID - Finds DLL and returns reference to object 
' Set - assign new object to variable
' Variable - work with object with this 


' Properties
' like chapters in a book
' Some will be read only, write obly and some read and write
' set objHome = CreatObject("House.Kitchen") - example declaration

' Methods
' Makes an object perform an action
' If it returns a value, it goes in paretheses, if it doesn't return a value, it doesnt go in parethesis

' FileSystem
' Drives
' Folders
' Files
' Text

' Copies files from one folder to another

Dim objFSO, objFolder, objFile, objNewFolder
'Create FilesSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder we want to copy files from. C:\Users\Llama\Desktop\Learning VB
Set objFolder = objFSO.GetFolder("C:\Users\Llama\Desktop\Learning VB")
'Create new folder called C:\Users\Llama\Desktop\Test
Set objNewFolder = objFSO.CreateFolder ("C:\Users\Llama\Desktop\Test")
For Each objFile In objFolder.Files
objFile.copy "C:\Users\Llama\Desktop\Test"
Next