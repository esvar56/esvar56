Private Sub RemoveDynamicControls()
    Dim ctrl As Control

    ' Loop through controls in the frame
    For Each ctrl In Me.Frame1.Controls
        If ctrl.Name = "lblDynamic" Or ctrl.Name = "txtDynamic" Then
            Me.Frame1.Controls.Remove ctrl.Name
        End If
    Next ctrl
End Sub
