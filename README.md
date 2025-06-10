' In clsDynamicButton
Public WithEvents btn As MSForms.CommandButton

Private Sub btn_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select a file"
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        If .Show = -1 Then
            MsgBox "Selected file: " & .SelectedItems(1)
        End If
    End With
End Sub
