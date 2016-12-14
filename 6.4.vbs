'====================6.4=====================
'
' Test Scripts by David Rhodes 12/11/16
'
'============================================

'Working with Text Files

Function BrowseForFile()
    With CreateObject("WScript.Shell")
        Dim fso, tempFolder, tempName, path
         
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set tempFolder = fso.GetSpecialFolder(2)
        
        tempName = fso.GetTempName() & ".hta"
        path = "HKCU\Volatile Environment\MsgResp"
        
        With tempFolder.CreateTextFile(tempName)
            .Write "<input type=file name=f>" & _
            "<script>f.click();(new ActiveXObject('WScript.Shell'))" & _
            ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
            "close();</script>"
            .Close
            
        End With
        .Run tempFolder & "\" & tempName, 1, True
        BrowseForFile = .RegRead(path)
        .RegDelete path
        fso.DeleteFile tempFolder & "\" & tempName
    End With
End Function

Dim objFSO, objFile, LogFile
LogFile = BrowseForFile
Const Forwriting = 2 ' 2 means the file will be overwritten when we write to it
Const Forappending = 8 ' 8 appends text to existing file

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(LogFile) Then
WScript.Echo "The File " & LogFile & " already exists."
Set objFile = objFSO.OpenTextFile(LogFile, Forappending)
objFile.Write vbCrLf & "New Entry " & Now

Else

WScript.Echo "The File " & LogFile & " will be created."
Set objFile = objFSO.CreateTextFile(LogFile, Forwriting)
objFile.Write vbCrLf & "New Entry " & Now
End If

objFile.Close


