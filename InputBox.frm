VERSION 5.00
Begin VB.Form frmInputBox 
   Caption         =   "Go To Line"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "InputBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_input 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label lbl_Prompt 
      AutoSize        =   -1  'True
      Caption         =   "Enter Line Number (1 - "
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngNumberOfLines As Long
    
Private Sub cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub cmd_OK_Click()
    Dim intRet As Integer
        
    If (IsNumeric(txt_input.Text) = False) Or (Val(txt_input.Text) < 1 Or Val(txt_input.Text) > lngNumberOfLines) Then
        intRet = MsgBox("Please enter an integer between 1 and " & lngNumberOfLines & ".", vbExclamation, "Java Editor")
        txt_input.Text = ""
        txt_input.SetFocus
        Exit Sub
    End If
    editor.GoToLine Val(txt_input.Text)
    Unload Me
End Sub

Private Sub Form_Load()
    lngNumberOfLines = editor.LineCount
    lbl_Prompt = lbl_Prompt & lngNumberOfLines & "):"
End Sub

