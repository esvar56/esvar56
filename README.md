' In the UserForm module
Dim btnHandler As clsDynamicButton ' Must be module-level

Private Sub UserForm_Initialize()
    Dim btn As MSForms.CommandButton

    ' Create dynamic command button
    Set btn = Me.Controls.Add("Forms.CommandButton.1", "btnBrowse", True)
    With btn
        .Caption = "Browse File"
        .Left = 20
        .Top = 20
        .Width = 100
        .Height = 30
    End With

    ' Link the button to the event handler class
    Set btnHandler = New clsDynamicButton
    Set btnHandler.btn = btn
End Sub
