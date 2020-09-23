Attribute VB_Name = "Module8"
Public Sub CreateTempFile()
Open "temp.tmp" For Output As #1
    Print #1, "y"
Close #1
End Sub
