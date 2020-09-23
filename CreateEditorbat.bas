Attribute VB_Name = "Module6"
'Procedure to create Editor.bat
Public Sub CreateEditorbat(str As String)
    Call changeapppath    'Changes Current Working Directory to Application Path
    Open "editor.bat" For Output As #1
    Print #1, str
    Close #1
End Sub

