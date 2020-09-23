Attribute VB_Name = "Module7"
Public Sub compilefile()
Call CreateTempFile
Dim temp As Double
If editor.calledfrom <> 8 Then
    compile_splash.Label2 = compile_splash.Label2 & " " & getFile(editor.sfile) & "..."
    compile_splash.Show
    editor.compiled = True
End If
temp = Shell("Redirect.exe", vbHide)
Dim c As String
rep:
Open "temp.tmp" For Input As #1
Do Until EOF(1)
    DoEvents
    Line Input #1, c
    If c = "y" Then
        Close #1
        GoTo rep
    Else
        Exit Do
   End If
Loop
Close #1
Call CreateTempFile
If editor.calledfrom <> 8 Then Unload compile_splash
Output.Show vbModal
End Sub

