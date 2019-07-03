'This macro reads the name of all the files in a directory (and its subdirectories) and prints them in the column of the cell selected. 

'InputBox to introduce the folder path
Sub folder_path()

    path = InputBox("Select the folder path:")
    Call list_folder_files(path)
    
End Sub

'Main program
Sub list_folder_files(path)
    Dim FSO, dir, subdir, file As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path = "" Then
        Exit Sub
    ElseIf Right(path, 1) <> "" Then
        path = path & ""
    End If
    
    On Error GoTo ErrHandler
    Set dir = FSO.GetFolder(path)
    
    For Each file In dir.Files
        ActiveCell.Value = FSO.getbasename(file.Name)
        ActiveCell.Offset(1, 0).Select
    Next
    
    For Each subdir In dir.subfolders
        list_folder_files (subdir)
    Next

    ActiveCell.EntireColumn.AutoFit
    Exit Sub
    
ErrHandler:
    ActiveCell.Value = "Invalid Path"
    
End Sub
