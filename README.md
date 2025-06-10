Private Sub UserForm_Initialize()
    Dim btn As MSForms.CommandButton
    Set btn = Me.Frame1.Controls.Add("Forms.CommandButton.1", "btnDynamic", True)

    With btn
        .Caption = "Click Me"
        .Left = 10
        .Top = 50
        .Width = 100
        .Height = 30
    End With
End Sub
