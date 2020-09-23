VERSION 5.00
Begin VB.Form setup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Folders And Files"
   ClientHeight    =   4050
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6015
   Icon            =   "Setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      Caption         =   "&Ok"
      Default         =   -1  'True
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
      Left            =   2160
      TabIndex        =   15
      ToolTipText     =   "Click to Save Changes"
      Top             =   3480
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   2400
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton CancelButton 
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
      Left            =   3480
      TabIndex        =   16
      ToolTipText     =   "Click to Cancel"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txt_path 
      Height          =   405
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Specify Command.com File (Optional)"
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txt_path 
      Height          =   405
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      ToolTipText     =   "Specify Browser Application File (Optional)"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txt_path 
      Height          =   405
      Index           =   3
      Left            =   2160
      TabIndex        =   10
      ToolTipText     =   "Specify API html File (Optional)"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txt_path 
      Height          =   405
      Index           =   4
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "Specify Default Open/Save Directory (Needed)"
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txt_path 
      Height          =   405
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Specify Java Program Directory (needed)"
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton setup_path 
      Caption         =   "..."
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
      Index           =   0
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   435
   End
   Begin VB.CommandButton setup_path 
      Caption         =   "..."
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
      Index           =   1
      Left            =   5400
      TabIndex        =   5
      Top             =   960
      Width           =   435
   End
   Begin VB.CommandButton setup_path 
      Caption         =   "..."
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
      Index           =   2
      Left            =   5400
      TabIndex        =   8
      Top             =   1560
      Width           =   435
   End
   Begin VB.CommandButton setup_path 
      Caption         =   "..."
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
      Index           =   3
      Left            =   5400
      TabIndex        =   11
      Top             =   2160
      Width           =   435
   End
   Begin VB.CommandButton setup_path 
      Caption         =   "..."
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
      Index           =   4
      Left            =   5400
      TabIndex        =   14
      Top             =   2760
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Ms-Dos App:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Browser File:"
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
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&API:"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label label5 
      AutoSize        =   -1  'True
      Caption         =   "&Default Open/Save:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Java Program Dir:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1530
   End
End
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelButton_Click()
    editor.SetFocus
    Unload Me
End Sub

'Procedure that loads the forms and also loads paths into textboxes
Private Sub Form_Load()
    'FormStayOnTop Me, True
    txt_path.Item(0).Text = editor.javapath
    txt_path.Item(1).Text = editor.commandpath
    txt_path.Item(2).Text = editor.browserpath
    txt_path.Item(3).Text = editor.apipath
    txt_path.Item(4).Text = editor.defaultpath
End Sub
'Procedure that saves all paths from the textboxes
Private Sub OKButton_Click()
    With editor
        .javapath = setup.txt_path.Item(0).Text
        .commandpath = setup.txt_path.Item(1).Text
        .browserpath = setup.txt_path.Item(2).Text
        .apipath = setup.txt_path.Item(3).Text
        .defaultpath = setup.txt_path.Item(4).Text
    End With
    
    Call modifyeditorini   'Save the settings in editor.ini file
    editor.SetFocus
    Unload Me
End Sub
'Procedure which calls PathFile form and also sets the label
Private Sub setup_path_Click(Index As Integer)
    editor.ind = Index
    editor.calledfrom = fromSetup
    'to indicated that PathFile form is called from setup form
    If Index = 0 Then
        PathFile.File1.Visible = False
        PathFile.Label1.Caption = "Select the Java Program Directory:"
        'PathFile.Drive1.Drive = getDrive(editor.javapath)
        'PathFile.Dir1.path = editor.javapath
    ElseIf Index = 1 Then
        PathFile.File1.Visible = True
        PathFile.Label1.Caption = "Select the Command.com File:"
        'PathFile.Drive1.Drive = getDrive(editor.commandpath)
        'PathFile.Dir1.path = getPath(editor.commandpath)
    ElseIf Index = 2 Then
        PathFile.File1.Visible = True
        PathFile.Label1.Caption = "Select any Browser Application File:"
        'PathFile.Drive1.Drive = getDrive(editor.browserpath)
        'PathFile.Dir1.path = getPath(editor.browserpath)
    ElseIf Index = 3 Then
        PathFile.File1.Visible = True
        PathFile.Label1.Caption = "Select any API html File:"
        'PathFile.Drive1.Drive = getDrive(editor.apipath)
        'PathFile.Dir1.Path = getPath(editor.apipath)
    ElseIf Index = 4 Then
        PathFile.File1.Visible = False
        PathFile.Label1.Caption = "Select the Default Open/SaveAs Directory:"
        'PathFile.Drive1.Drive = getDrive(editor.defaultpath)
        'PathFile.Dir1.path = editor.defaultpath
    End If
    PathFile.Show vbModal
End Sub
