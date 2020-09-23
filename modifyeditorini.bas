Attribute VB_Name = "Module4"
'Procedure to modify editor.ini file
Public Sub modifyeditorini()
    ' saving all the settings in editor.ini file
    Call changeapppath    'Changes Current Working Directory to Application Path
    'opening editor.ini file
    With editor.writepad
        Open "editor.ini" For Output As #1
        Print #1, .BackColor
        Print #1, .SelColor
        Print #1, .SelBold
        Print #1, .SelItalic
        Print #1, .SelFontName
        Print #1, .SelFontSize
    End With
    
    With editor
        Print #1, .javapath
        Print #1, .commandpath
        Print #1, .browserpath
        Print #1, .apipath
        Print #1, .defaultpath
    End With

    Close #1 'closing editor.ini file

End Sub
