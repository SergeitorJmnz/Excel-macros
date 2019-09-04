'This macro makes files with the name of the cells in a range.

Sub copy_file()
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim xRg As Range, xCell As Range
    Dim xSFldDlg As FileDialog, xDFldDlg As FileDialog
    Dim xSFileStr As Variant, xDPathStr As Variant
    Dim xVal As String
    
    On Error Resume Next
    
    Set xSFldDlg = Application.FileDialog(msoFileDialogFilePicker)
    xSFldDlg.Title = "Selecciona el archivo a copiar"
    If xSFldDlg.Show <> -1 Then Exit Sub
    
    xSFileStr = xSFldDlg.SelectedItems.Item(1)
    
    Set xRg = Application.InputBox("Selecciona los nombres de archivos:", , ActiveWindow.RangeSelection.Address, , , , , 8)
    If xRg Is Nothing Then Exit Sub
    
    Set xDFldDlg = Application.FileDialog(msoFileDialogFolderPicker)
    xDFldDlg.Title = "Selecciona la carpeta de destino de los archivos:"
    If xDFldDlg.Show <> -1 Then Exit Sub
    
    xDPathStr = xDFldDlg.SelectedItems.Item(1) & "\"
    
    For Each xCell In xRg
        xVal = xCell.Value
        If TypeName(xVal) = "String" And xVal <> "" Then
            objFSO.copyFile xSFileStr, xDPathStr & xVal & ".pdf"
        End If
    Next
    
End Sub
