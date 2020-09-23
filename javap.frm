VERSION 5.00
Begin VB.Form javap 
   Caption         =   "Javap - The Class File Disassembler"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "javap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_J 
      Height          =   285
      Left            =   2760
      TabIndex        =   28
      ToolTipText     =   "Every java flag must be preceded by -J (except the first one)"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txt_paths 
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   27
      ToolTipText     =   "write or select the path(s)"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txt_paths 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   26
      ToolTipText     =   "write or select the path(s)"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txt_paths 
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   25
      ToolTipText     =   "write or select the path(s)"
      Top             =   1170
      Width           =   2775
   End
   Begin VB.CommandButton cmd_paths 
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
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   24
      ToolTipText     =   "select the extension directories"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmd_paths 
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
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   23
      ToolTipText     =   "select the bootclasspath(s)"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmd_paths 
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
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   22
      ToolTipText     =   "select the classpath(s)"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CheckBox chk_s 
      Caption         =   "-s"
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
      Left            =   1680
      TabIndex        =   21
      ToolTipText     =   "print internal type signatures"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CheckBox chk_l 
      Caption         =   "-l"
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
      TabIndex        =   20
      ToolTipText     =   "print line number and local variable tables"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CheckBox chk_c 
      Caption         =   "-c"
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
      Left            =   1680
      TabIndex        =   19
      ToolTipText     =   "Dissasseble the code"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CheckBox chk_b 
      Caption         =   "-b"
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
      TabIndex        =   18
      ToolTipText     =   "Backward compatability with javap in JDK 1.1"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CheckBox chk_help 
      Caption         =   "-help"
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
      Left            =   1680
      TabIndex        =   17
      ToolTipText     =   "print this usage message"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chk_verbose 
      Caption         =   "verbose"
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
      TabIndex        =   16
      ToolTipText     =   "print stack size, number of locals and arguments"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CheckBox chk_private 
      Caption         =   "-private"
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
      Left            =   1680
      TabIndex        =   15
      ToolTipText     =   "show all classes and members"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CheckBox chk_package 
      Caption         =   "-package"
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
      TabIndex        =   14
      ToolTipText     =   "show only package/protected/public classes and members"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox chk_protected 
      Caption         =   "-protected"
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
      Left            =   1680
      TabIndex        =   13
      ToolTipText     =   "show only protected/public classes and members"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chk_public 
      Caption         =   "-public"
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
      TabIndex        =   12
      ToolTipText     =   "show only public classes and members"
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox chk_J 
      Caption         =   "-J <flag(s)>"
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
      TabIndex        =   11
      ToolTipText     =   "pass <flag(s)> directly to runtime machine"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CheckBox chk_extdirs 
      Caption         =   "-extdirs <(dirs)>"
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
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox chk_bootclasspath 
      Caption         =   "-bootclasspath <path(s)>"
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
      TabIndex        =   9
      ToolTipText     =   "override location of clas files loaded"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CheckBox chk_classpath 
      Caption         =   "-classpath <path(s)>"
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
      TabIndex        =   8
      ToolTipText     =   "specify where to find user class files"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmd_help 
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
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   "Click to get help on javap Disassembler"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmd_reset 
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
      Left            =   3120
      TabIndex        =   6
      ToolTipText     =   "Click to Reset all options"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmd_cancel 
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
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Click to Cancel"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ok 
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
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "Click to disasseble the class file"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmd_javap 
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
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      ToolTipText     =   "Click to select class file(s) one by one"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txt_javap 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Modify here if necessary"
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "options:"
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
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "javap [options] <classes>..."
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
      Width           =   2355
   End
End
Attribute VB_Name = "javap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sel As Byte ' 1 for selecting class file
                   ' 2 for setting classpath
                   ' 3 for setting bootclasspath
                   ' 4 for setting extdirs
Public fname As String

Sub chk_b_Click()
    Call changetext
End Sub
'Procedure to update javap text box
Public Sub changetext()
    txt_javap.Text = Chr(34) & editor.javapath & "\javap" & Chr(34) & " "
            ' selection of options
    If chk_classpath.Value = 1 Then txt_javap.Text = txt_javap.Text & "-classpath " & txt_paths(0).Text & " "
    If chk_bootclasspath.Value = 1 Then txt_javap.Text = txt_javap.Text & "-bootclasspath " & txt_paths(1).Text & " "
    If chk_extdirs.Value = 1 Then txt_javap.Text = txt_javap.Text & "-extdirs " & txt_paths(2).Text & " "
    If chk_J.Value = 1 Then txt_javap.Text = txt_javap.Text & "-J " & txt_J.Text & " "
    If chk_public.Value = 1 Then txt_javap.Text = txt_javap.Text & "-public "
    If chk_protected.Value = 1 Then txt_javap.Text = txt_javap.Text & "-protected "
    If chk_package.Value = 1 Then txt_javap.Text = txt_javap.Text & "-package "
    If chk_private.Value = 1 Then txt_javap.Text = txt_javap.Text & "-private "
    If chk_verbose.Value = 1 Then txt_javap.Text = txt_javap.Text & "-verbose "
    If chk_help.Value = 1 Then txt_javap.Text = txt_javap.Text & "-help "
    If chk_b.Value = 1 Then txt_javap.Text = txt_javap.Text & "-b "
    If chk_c.Value = 1 Then txt_javap.Text = txt_javap.Text & "-c "
    If chk_l.Value = 1 Then txt_javap.Text = txt_javap.Text & "-l "
    If chk_s.Value = 1 Then txt_javap.Text = txt_javap.Text & "-s "
    txt_javap.Text = txt_javap.Text & fname
End Sub

Private Sub chk_bootclasspath_Click()
    If chk_bootclasspath.Value = 1 Then
        txt_paths(1).Enabled = True
        cmd_paths(1).Enabled = True
    Else
        txt_paths(1).Enabled = False
        cmd_paths(1).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_c_Click()
    Call changetext
End Sub

Private Sub chk_classpath_Click()
    If chk_classpath.Value = 1 Then
        txt_paths(0).Enabled = True
        cmd_paths(0).Enabled = True
    Else
        txt_paths(0).Enabled = False
        cmd_paths(0).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_extdirs_Click()
    If chk_extdirs.Value = 1 Then
        txt_paths(2).Enabled = True
        cmd_paths(2).Enabled = True
    Else
        txt_paths(2).Enabled = False
        cmd_paths(2).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_help_Click()
    Call changetext
End Sub

Private Sub chk_J_Click()
    If chk_J.Value = 1 Then
        txt_J.Enabled = True
    Else
        txt_J.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_l_Click()
    Call changetext
End Sub

Private Sub chk_package_Click()
    Call changetext
End Sub

Private Sub chk_private_Click()
    Call changetext
End Sub

Private Sub chk_protected_Click()
    Call changetext
End Sub

Private Sub chk_public_Click()
    Call changetext
End Sub

Private Sub chk_s_Click()
    Call changetext
End Sub

Private Sub chk_verbose_Click()
    Call changetext
End Sub

Private Sub cmd_javap_Click()
    editor.calledfrom = fromJavap 'set called from to 6
    'to indicated that PathFile form is called from javap form
    editor.ind = -1
    PathFile.File1.Visible = True
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Pattern = "*.class"
    PathFile.Label1.Caption = "Select the class File :"
    PathFile.Show vbModal
    Call changetext
End Sub
'Procedure to unload the form
Private Sub cmd_Cancel_Click()
    Unload Me
End Sub
'Procedure which creates Editor.bat and executes it
Private Sub cmd_OK_Click()
    CreateEditorbat txt_javap.Text & vbCrLf & "pause"
    Dim temp As Double
    temp = Shell("editor.bat ", vbMaximizedFocus)
    'editor.calledfrom = 8
    'Call compilefile
End Sub

Private Sub cmd_paths_Click(Index As Integer)
    PathFile.File1.Visible = False
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.Label1.Caption = "Select the Directory :"
    editor.calledfrom = fromJavap 'set called from to 6
    'to indicated that PathFile form is called from javap form
    editor.ind = Index
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub cmd_reset_Click()
    Call Form_Load
End Sub

Private Sub Form_Load()
    fname = "Noname"
    chk_classpath.Value = 1: chk_bootclasspath.Value = 0: chk_extdirs.Value = 0: chk_J.Value = 0
    txt_paths(1).Text = "": txt_paths(1).Enabled = False: cmd_paths(1).Enabled = False
    txt_paths(2).Text = "": txt_paths(2).Enabled = False: cmd_paths(2).Enabled = False
    txt_J.Text = "": txt_J.Enabled = False
    txt_paths(0).Text = Chr(34) & editor.defaultpath & Chr(34) & ";"
    txt_javap.Text = Chr(34) & editor.javapath & "\javap" & Chr(34) & " -classpath " & Chr(34) & editor.defaultpath & Chr(34) & " Noname"
    chk_public.Value = 0: chk_protected.Value = 0: chk_package.Value = 0: chk_private.Value = 0
    chk_verbose.Value = 0: chk_help.Value = 0: chk_b.Value = 0: chk_c.Value = 0: chk_l.Value = 0: chk_s.Value = 0
End Sub

Private Sub txt_J_Change()
    Call changetext
End Sub

Private Sub txt_paths_Change(Index As Integer)
    Call changetext
End Sub
