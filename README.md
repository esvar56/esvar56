Dim filePath As String
Dim fileName As String

filePath = "C:\Users\YourName\Documents\example.xlsx"
fileName = Dir(filePath)  ' Extracts: example.xlsx

MsgBox fileName
