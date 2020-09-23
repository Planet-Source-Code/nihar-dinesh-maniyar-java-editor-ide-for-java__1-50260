VERSION 5.00
Begin VB.Form Replace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace"
   ClientHeight    =   3315
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "Replace.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Replaceall 
      Caption         =   "Replace &All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Replace_With 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmd_Replace 
      Caption         =   "&Replace..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox chk_matchcase 
      Caption         =   "&Match Case"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CheckBox chk_wwonly 
      Caption         =   "Find Whole Word &Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Frame Direction 
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   11
      Top             =   1440
      Width           =   1935
      Begin VB.OptionButton opt_down 
         Caption         =   "&Down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opt_up 
         Caption         =   "&Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox FindWhat 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton FindNext 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl_replacewith 
      AutoSize        =   -1  'True
      Caption         =   "Replace &With:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1230
   End
   Begin VB.Label lbl_findwhat 
      AutoSize        =   -1  'True
      Caption         =   "&Find What:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "Replace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Procedure that unloads the form
Private Sub Cancel_Click()
    Unload Me
End Sub
'Procedure that call setvalues and findandreplace procedure
Private Sub cmd_Replace_Click()
    setvalues
    findandreplace 1
End Sub
'Procedure that call setvalues and findandreplace procedure
Private Sub FindNext_Click()
    setvalues
    findandreplace 0
End Sub
'Procedure that sets the find and replace option values
Private Sub setvalues()
    editor.writepad.HideSelection = False
    editor.strfind = FindWhat.Text
    editor.strreplace = Replace_With.Text
    editor.findflags = chk_wwonly.Value * 2 + chk_matchcase.Value * 4
    editor.updown = opt_up
    gCase = chk_matchcase.Value
    gWholeWord = chk_wwonly.Value
End Sub

Private Sub Form_Initialize()
    If editor.strfind <> "" Then
        FindWhat.Text = editor.strfind
        opt_up = editor.updown
        opt_down = Not opt_up
        chk_wwonly.Value = gWholeWord
        chk_matchcase.Value = gCase
        'Replace_With.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        If KeyCode = Asc("F") Then FindWhat.SetFocus
        If KeyCode = Asc("W") And Replace_With.Visible = True Then Replace_With.SetFocus
    End If
    KeyCode = 0
    Shift = 0
End Sub
Private Sub FindWhat_Change()
    If FindWhat.Text = "" Then
        FindNext.Enabled = False
    Else
        FindNext.Enabled = True
    End If
    If strReplacementDone > 0 Then
        editor.findflags = MsgBox(intReplacementDone & "Replacements done.", vbExclamation, "Java Editor 1.0")
    Else
        editor.findflags = MsgBox("Search Text is not found.", vbExclamation, "Java Editor 1.0")
    End If
End Sub
'Procedure that replaces all the values
Private Sub Replaceall_Click()
    setvalues
    Do 'loop for replacing all the text till the end
        findandreplace 2
    Loop While intPos <> 0
End Sub
