VERSION 5.00
Begin VB.Form appletviewer_with 
   Caption         =   "appletviewer - The Java Applet Viewer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "Appletviewer_with.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_j 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   12
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txt_encoding 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton help 
      Caption         =   "&Help"
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
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Click to get help on Javac Compiler"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton reset 
      Caption         =   "&Reset"
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
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Click to reset all the options"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cancel 
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
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      ToolTipText     =   "Click to Cancel"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton ok 
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
      Height          =   495
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   "Click to Compile the source file"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox chk_j 
      Caption         =   "-J<javaflag(s)>"
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
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Generate no warnings"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CheckBox chk_encoding 
      Caption         =   "-encoding <encodingname>"
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
      TabIndex        =   5
      ToolTipText     =   "Optimize; may hinder debugging or large class files "
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CheckBox chk_debug 
      Caption         =   "-debug"
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
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "To generate all or some or no debugging info"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmd_appletviewer 
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
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "Click to Select .Java File"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txt_appletviewer 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Modify here if necessary"
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "appletviewer [options] url | file(s)..."
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
      TabIndex        =   0
      Top             =   120
      Width           =   2970
   End
End
Attribute VB_Name = "appletviewer_with"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This function is used to unload this form
Private Sub Cancel_Click()
    editor.SetFocus
    Unload Me
End Sub

Private Sub chk_debug_Click()
    Call changetext
End Sub
'This function is called when any of the checkboxes are clicked
'and is used to update the appletviewer textbox
Private Sub changetext()
    
    'Generating appletviewer text that should be executed
    txt_appletviewer.Text = Chr(34) & editor.javapath & "\appletviewer" & Chr(34) & " "
    If chk_debug.Value = 1 Then txt_appletviewer.Text = txt_appletviewer.Text & "-debug "
    If chk_encoding.Value = 1 Then
        txt_appletviewer.Text = txt_appletviewer.Text & "-encoding "
        If txt_encoding.Text <> "" Then txt_appletviewer.Text = txt_appletviewer.Text & txt_encoding.Text & " "
    End If
    If chk_J.Value = 1 Then
        txt_appletviewer.Text = txt_appletviewer.Text & "-j "
        If txt_J.Text <> "" Then txt_appletviewer.Text = txt_appletviewer.Text & txt_J.Text & " "
    End If
    txt_appletviewer.Text = txt_appletviewer.Text & Chr(34) & getFile(editor.sfile) & Chr(34)

End Sub
Private Sub chk_encoding_Click()
    Call changetext
    If chk_encoding.Value = 1 Then
        txt_encoding.Enabled = True
        txt_encoding.SetFocus
    Else
        txt_encoding.Enabled = False
    End If
End Sub

Private Sub chk_J_Click()
    Call changetext
    If chk_J.Value = 1 Then
        txt_J.Enabled = True
        txt_J.SetFocus
    Else
        txt_J.Enabled = False
    End If
End Sub
'This function is called when a user wants to select a class file.
Private Sub cmd_appletviewer_Click()
    'calledfrom variable is set to 4 to indicate that
    'PathFile is called from Appletviewer
    editor.calledfrom = fromAppletviewer
    'Setting PathFile forms controls dynamically
    PathFile.File1.Visible = True
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Pattern = "*.java;*.htm;*.html"
    PathFile.Label1.Caption = "Select the html/java File :"
    'calling PathFile Form
    PathFile.Show vbModal
    PathFile.File1.Pattern = "*.*"
End Sub

Private Sub Form_Load()
    'SetParent Me.hwnd, editor.hwnd
    'FormStayOnTop Me, True
    txt_appletviewer.Text = Chr(34) & editor.javapath & "\appletviewer" & Chr(34) & " " & Chr(34) & editor.sfile & Chr(34)
End Sub

Private Sub Form_Resize()
    editor.SetFocus
End Sub

Private Sub Form_Unload(cancel As Integer)
    FormStayOnTop Me, False
End Sub
'This procedure is called when a user wants to execute an applet
'this procedure is resposible for creating Editor.bat file and execute it
Private Sub ok_Click()
    Dim path As String, drv As String
    drv = getDrive(editor.sfile)
    path = getPath(editor.sfile)
    'calling CreateEditorbat procedure to generate Editor.bat file
    CreateEditorbat "@echo off" & vbCrLf & "cd " & Chr(34) & path & Chr(34) & vbCrLf & drv & vbCrLf & txt_appletviewer.Text & vbCrLf & "pause"

    Dim temp As Double
    'Executing the applet
    temp = Shell("editor.bat ", vbMaximizedFocus)
    Call changeapppath

End Sub

'This procedure is used to reset all the settings
Private Sub reset_Click()
    chk_J.Value = 0
    chk_encoding.Value = 0
    chk_debug = 0
    txt_J.Text = ""
    txt_J.Enabled = False
    txt_encoding.Enabled = False
    txt_encoding.Text = ""
    txt_appletviewer.Text = Chr(34) & editor.javapath & "\appletviewer" & Chr(34) & " " & Chr(34) & editor.sfile & Chr(34)

End Sub

Private Sub txt_encoding_Change()
    Call changetext
End Sub

Private Sub txt_J_Change()
    Call changetext
End Sub
