Attribute VB_Name = "Module3"
Public gCurPos As Integer
Public gCase As Boolean
Public gWholeWord As Boolean
Public intReplacementDone As Integer
Public intPos As Integer
'operation  0 for find
'           1 for replace
'           2 for replace all

Public Sub findandreplace(operation As Byte)
    'Dim stringfound As Boolean 'stringfound to check string found or not


    Dim strSourceString As String
    Dim strFindString As String
    Dim intstart As Integer
    Dim intOffset As Integer

    If editor.strfind <> "" And operation = 1 Then
        If StrComp(editor.writepad.SelText, editor.strfind, vbTextCompare) = 0 Then
            editor.writepad.SelText = editor.strreplace
            editor.modified = True ' need to be saved
            Exit Sub
        End If
    End If
    
    ' Set offset variable based on cursor position.
    If (gCurPos = editor.writepad.SelStart) Then
        intOffset = 1
    Else
        intOffset = 0
    End If
    
    intstart = editor.writepad.SelStart + intOffset

    If gCase = False Then
        strSourceString = UCase(editor.writepad.Text)
        strFindString = UCase(editor.strfind)
    Else
        strSourceString = editor.writepad.Text
        strFindString = editor.strfind
    End If

    If editor.updown = False Then
        'editor.foundposition = editor.writepad.Find(editor.strfind, editor.position, 0, editor.findflags)
        'intPos = InStr(intstart + 1, strSourceString, strFindString)
        intPos = editor.writepad.Find(strFindString, intstart + 1, , editor.findflags) + 1
    Else
        'intPos = editor.writepad.Find(strFindString, , 0, editor.findflags)
        For intPos = intstart - 1 To 0 Step -1
            If intPos = 0 Then Exit For
            If Mid(strSourceString, intPos, Len(strFindString)) = strFindString Then Exit For
        Next intPos
    End If
    
    If intPos <> 0 Then
        editor.writepad.SelStart = intPos - 1
        editor.writepad.SelLength = Len(strFindString)
            
        ' if operation is replace then replace the found text
        If operation <> 0 Then
            editor.writepad.SelText = editor.strreplace
            editor.modified = True ' need to be saved
            If operation = 2 Then intReplacementDone = intReplacementDone + 1
        End If
    Else
        If intPos And operation = 1 Then
            editor.findflags = MsgBox("The Specified Region Has been searched ", vbExclamation, "Java Editor 1.0")
        Else
            editor.findflags = MsgBox("Cannot find " & Chr(34) & editor.strfind & Chr(34), vbExclamation, "Not Found")
        End If
    End If
    
    If editor.updown = False Then
        gCurPos = editor.writepad.SelStart
    Else
        gCurPos = editor.writepad.SelStart - Len(strSourceString)
    End If
End Sub
