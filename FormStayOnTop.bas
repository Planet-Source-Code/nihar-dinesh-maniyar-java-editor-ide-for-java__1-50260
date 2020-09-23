Attribute VB_Name = "Module1"
'Public section
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndinsertafter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wflags As Long) As Long
Public Const swp_nosize = &H1
Public Const swp_nomove = &H2
Public Const swp_noactivate = &H10
Public Const hwnd_topmost = -1
Public Const hwnd_notopmost = -2
Public Const swp_noownerzoreder = &H200
Public Const swp_showwindow = &H40
'Procedure to make a form stay on top
Sub FormStayOnTop(FormToSet As Form, OnTop As Boolean)
Dim lhwnd As Long
Dim lflags As Long
Dim lposflag As Long
lhwnd = FormToSet.hwnd
lflags = swp_nomove Or swp_nosize Or swp_showwindow Or swp_noactivate
Select Case OnTop
    Case True
        lposflag = hwnd_topmost
    Case False
        lposflag = hwnd_notopmost
End Select
SetWindowPos lhwnd, lposflag, 0, 0, 0, 0, lflags
End Sub

