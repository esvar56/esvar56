






Private Sub UserForm_Initialize()
    Dim txtBox As Object
    Dim lbl As Object
    Dim frame As Object

    Set frame = Me.Frame1

    ' Add Label
    Set lbl = frame.Controls.Add("Forms.Label.1", "lblDynamic", True)
    With lbl
        .Caption = "Dynamic Label:"
        .Left = 10
        .Top = 10
        .Width = 100
    End With

    ' Add TextBox
    Set txtBox = frame.Controls.Add("Forms.TextBox.1", "txtDynamic", True)
    With txtBox
        .Left = 120
        .Top = 10
        .Width = 150
    End With
End Sub




## Hi there 👋

<!--
**esvar56/esvar56** is a ✨ _special_ ✨ repository because its `README.md` (this file) appears on your GitHub profile.

Here are some ideas to get you started:

- 🔭 I’m currently working on ...
- 🌱 I’m currently learning ...
- 👯 I’m looking to collaborate on ...
- 🤔 I’m looking for help with ...
- 💬 Ask me about ...
- 📫 How to reach me: ...
- 😄 Pronouns: ...
- ⚡ Fun fact: ...
-->
