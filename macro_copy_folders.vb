'This macro copies the subdirectories with the name of the selected cells from one folder to another.

Sub copy_dirs()
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Dim xRg As Range, xCell As Range
    Dim xSFldDlg As FileDialog, xDFldDlg As FileDialog
    Dim xSPathStr As Variant, xDPathStr As Variant
    Dim xVal As String
    
    On Error Resume Next
    
    Set xRg = Application.InputBox("Please select the folder names:", , ActiveWindow.RangeSelection.Address, , , , , 8)
    If xRg Is Nothing Then Exit Sub
    
    Set xSFldDlg = Application.FileDialog(msoFileDialogFolderPicker)
    xSFldDlg.Title = "Please select the original folder:"
    If xSFldDlg.Show <> -1 Then Exit Sub
    
    xSPathStr = xSFldDlg.SelectedItems.Item(1) & "\"
    
    Set xDFldDlg = Application.FileDialog(msoFileDialogFolderPicker)
    xDFldDlg.Title = "Please select the destination folder:"
    If xDFldDlg.Show <> -1 Then Exit Sub
    
    xDPathStr = xDFldDlg.SelectedItems.Item(1) & "\"
    
    For Each xCell In xRg
        xVal = xCell.Value
        If TypeName(xVal) = "String" And xVal <> "" Then
            objFSO.copyFolder xSPathStr & xVal, xDPathStr & xVal
        End If
    Next
    
End Sub
