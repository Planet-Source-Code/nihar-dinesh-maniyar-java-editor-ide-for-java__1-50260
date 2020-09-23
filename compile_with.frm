VERSION 5.00
Begin VB.Form compile_with 
   Caption         =   "Javac - The Java Compiler..."
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   Icon            =   "compile_with.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_g 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   32
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmd_paths 
      Caption         =   "..."
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
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   31
      ToolTipText     =   "Click to select destination directory"
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton cmd_paths 
      Caption         =   "..."
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
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   30
      ToolTipText     =   "Click to select extensions"
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton cmd_paths 
      Caption         =   "..."
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
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   29
      ToolTipText     =   "Click to select bootclasspath(s)"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmd_paths 
      Caption         =   "..."
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
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   28
      ToolTipText     =   "Click to select sourcepath(s)"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton cmd_paths 
      Caption         =   "..."
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
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   27
      ToolTipText     =   "Click to select classpath(s)"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txt_target 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   26
      ToolTipText     =   "Write the VM version"
      Top             =   5250
      Width           =   2295
   End
   Begin VB.TextBox txt_encoding 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   25
      ToolTipText     =   "Write the Encoding name"
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   24
      ToolTipText     =   "Write or Select the paths"
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   23
      ToolTipText     =   "Write or Select the paths"
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   22
      ToolTipText     =   "Write or Select the paths"
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   21
      ToolTipText     =   "Write or Select the paths"
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   20
      ToolTipText     =   "Write or Select the paths"
      Top             =   3120
      Width           =   3375
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
      Left            =   5400
      TabIndex        =   19
      ToolTipText     =   "Click to get help on Javac Compiler"
      Top             =   2400
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
      Left            =   3960
      TabIndex        =   18
      ToolTipText     =   "Click to reset all the options"
      Top             =   2400
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
      Left            =   5400
      TabIndex        =   17
      ToolTipText     =   "Click to Cancel"
      Top             =   1800
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
      Left            =   3960
      TabIndex        =   16
      ToolTipText     =   "Click to Compile the source file"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox chk_target 
      Caption         =   "-target <release>"
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
      TabIndex        =   15
      ToolTipText     =   "Generate class files for Specific VM version"
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CheckBox chk_encoding 
      Caption         =   "-encoding <encoding>"
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
      TabIndex        =   14
      ToolTipText     =   "Specify character encoding used by source files"
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CheckBox chk_d 
      Caption         =   "-d <directory>"
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
      TabIndex        =   13
      ToolTipText     =   "Specify where to place generated class files"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CheckBox chk_extdirs 
      Caption         =   "-extdirs <dir(s)>"
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
      ToolTipText     =   "Override locations of installed extensions"
      Top             =   4200
      Width           =   1695
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
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Override location of bootstrap class files"
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CheckBox chk_sourcepath 
      Caption         =   "-sourcepath <path(s)>"
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
      TabIndex        =   10
      ToolTipText     =   "Specify where to find input source files"
      Top             =   3480
      Width           =   2295
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
      Height          =   195
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Specify where to find user class files"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CheckBox chk_deprecation 
      Caption         =   "-deprecation"
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
      TabIndex        =   8
      ToolTipText     =   "Output Source locations where deprecated APIs are used"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CheckBox chk_verbose 
      Caption         =   "-verbose"
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
      TabIndex        =   7
      ToolTipText     =   "Output messages about what the compiler is doing"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CheckBox chk_nowarn 
      Caption         =   "-nowarn"
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
      Width           =   1095
   End
   Begin VB.CheckBox chk_O 
      Caption         =   "-O"
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
      Width           =   615
   End
   Begin VB.CheckBox chk_g 
      Caption         =   "-g[:none | {lines,var,source}] "
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
      Width           =   2895
   End
   Begin VB.CommandButton cmd_compile 
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
      Left            =   6360
      TabIndex        =   2
      ToolTipText     =   "Click to Select .Java File"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txt_compile 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Modify here if necessary"
      Top             =   360
      Width           =   6015
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "javac [options] <sourcefile | @filelist>..."
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
      Width           =   3390
   End
End
Attribute VB_Name = "compile_with"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This procedure is called if Cancel Button is pressed
Private Sub Cancel_Click()
    editor.SetFocus
    Unload Me
End Sub

Private Sub chk_bootclasspath_Click()
    'changetext is called to update the Compile_with text
    Call changetext
    If chk_bootclasspath.Value = 1 Then
        txt_paths(2).Enabled = True
        txt_paths(2).SetFocus
        cmd_paths(2).Enabled = True
    Else
        cmd_paths(2).Enabled = False
        txt_paths(2).Enabled = False
    End If
End Sub

Private Sub chk_classpath_Click()
    Call changetext
    If chk_classpath.Value = 1 Then
        cmd_paths(0).Enabled = True
        txt_paths(0).Enabled = True
        txt_paths(0).SetFocus
    Else
        cmd_paths(0).Enabled = False
        txt_paths(0).Enabled = False
End If
End Sub

Private Sub chk_d_Click()
    Call changetext
    If chk_d.Value = 1 Then
        cmd_paths(4).Enabled = True
        txt_paths(4).Enabled = True
        txt_paths(4).SetFocus
    Else
        cmd_paths(4).Enabled = False
        txt_paths(4).Enabled = False
    End If
End Sub

Private Sub chk_deprecation_Click()
    Call changetext
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

Private Sub chk_extdirs_Click()
    Call changetext
    If chk_extdirs.Value = 1 Then
        cmd_paths(3).Enabled = True
        txt_paths(3).Enabled = True
        txt_paths(3).SetFocus
    Else
        cmd_paths(3).Enabled = False
        txt_paths(3).Enabled = False
    End If
End Sub

Private Sub chk_g_Click()
    Call changetext
    If chk_g.Value = 1 Then
        txt_g.Enabled = True
        txt_g.SetFocus
    Else
        txt_g.Enabled = False
    End If
End Sub
'This procedure is called when any of the checkboxes are clicked
'and is used to update the compile_with textbox
Private Sub changetext()
    txt_compile.Text = Chr(34) & editor.javapath & "\javac" & Chr(34) & " "
    
    If chk_g.Value = 1 Then
        txt_compile.Text = txt_compile.Text & "-g "
        If txt_g.Text <> "" Then txt_compile.Text = txt_compile.Text & " " & txt_g.Text & " "
    End If
    
    If chk_O.Value = 1 Then txt_compile.Text = txt_compile.Text & "-O "
    
    If chk_nowarn.Value = 1 Then txt_compile.Text = txt_compile.Text & "-nowarn "
    
    If chk_verbose.Value = 1 Then txt_compile.Text = txt_compile.Text & "-verbose "
    
    If chk_deprecation.Value = 1 Then txt_compile.Text = txt_compile.Text & "-deprecation "
    
    If chk_classpath.Value = 1 Then
        txt_compile.Text = txt_compile.Text & "-classpath "
        If txt_paths(0).Text <> "" Then txt_compile.Text = txt_compile.Text & " " & txt_paths(0).Text & " "
    End If
    
    If chk_sourcepath.Value = 1 Then
        txt_compile.Text = txt_compile.Text & "-sourcepath "
        If txt_paths(1).Text <> "" Then txt_compile.Text = txt_compile.Text & " " & txt_paths(1).Text & " "
    End If
    
    If chk_bootclasspath.Value = 1 Then
        txt_compile.Text = txt_compile.Text & "-bootclasspath "
        If txt_paths(2) <> "" Then txt_compile.Text = txt_compile.Text & " " & txt_paths(2).Text & " "
    End If
    
    If chk_extdirs.Value = 1 Then
        txt_compile.Text = txt_compile.Text & "-extdirs "
        If txt_paths(3) <> "" Then txt_compile.Text = txt_compile.Text & " " & txt_paths(3).Text & " "
    End If
    
    If chk_d.Value = 1 Then
        txt_compile.Text = txt_compile.Text & "-d "
        If txt_paths(4).Text <> "" Then txt_compile.Text = txt_compile.Text & " " & txt_paths(4).Text & " "
    End If
    
    If chk_encoding.Value = 1 Then
        txt_compile.Text = txt_compile.Text & "-encoding "
        If txt_encoding.Text <> "" Then txt_compile.Text = txt_compile.Text & " " & txt_encoding.Text & " "
    End If
    
    If chk_target.Value = 1 Then
        txt_compile.Text = txt_compile.Text & "-target "
        If txt_target.Text <> "" Then txt_compile.Text = txt_compile.Text & " " & txt_target.Text & " "
    End If
    
    'Generating the compile_with text
    txt_compile.Text = txt_compile.Text & Chr(34) & editor.sfile & Chr(34)
End Sub
Private Sub chk_nowarn_Click()
    Call changetext
End Sub

Private Sub chk_O_Click()
    Call changetext
End Sub

Private Sub chk_sourcepath_Click()
    Call changetext
    If chk_sourcepath.Value = 1 Then
        txt_paths(1).Enabled = True
        cmd_paths(1).Enabled = True
        txt_paths(1).SetFocus
    Else
        txt_paths(1).Enabled = False
        cmd_paths(1).Enabled = False
    End If
End Sub

Private Sub chk_target_Click()
    Call changetext
    If chk_target.Value = 1 Then
        txt_target.Enabled = True
        txt_target.SetFocus
    Else
        txt_target.Enabled = False
    End If
End Sub

Private Sub chk_verbose_Click()
    Call changetext
End Sub

Private Sub cmd_compile_Click()
    'setting calledfrom variable to 3 to indicate that PathFile
    'form is called from compile_with
    editor.calledfrom = fromCompileWithForSelectingFile
    'setting PathFile controls dynamically
    PathFile.File1.Visible = True
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Pattern = "*.java"
    PathFile.Label1.Caption = "Select the Source File :"
    'calling PathFile form
    PathFile.Show vbModal
End Sub

Private Sub cmd_paths_Click(Index As Integer)
    'setting calledfrom variable to 1 to indicate that PathFile
    'form is called from compile_with with arguments
    editor.calledfrom = fromCompileWithForArguments
    editor.ind = Index
    'setting PathFile controls dynamically
    PathFile.File1.Visible = False
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.Label1.Caption = "Select the Directory:"
    'calling PathFile form
    PathFile.Show vbModal
End Sub
Private Sub Form_Load()
    'SetParent Me.hwnd, editor.hwnd
    'FormStayOnTop Me, True
    txt_compile.Text = Chr(34) & editor.javapath & "\javac" & Chr(34) & " " & Chr(34) & editor.sfile & Chr(34)
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
    Dim icon As String
    icon = App.path + "\icons\Trffc10c.ico"
    'Calling CreateEditorbat procedure to create Editor.bat
    CreateEditorbat txt_compile.Text
    'setting the compiled variable to true
    editor.compiled = True
    editor.MouseIcon = LoadPicture(icon)
    compile_splash.Label2 = compile_splash.Label2 & " " & getFile(editor.sfile) & "..."
    'Calling Compile_splash form
    compile_splash.Show vbModal
End Sub

Private Sub reset_Click()
    chk_bootclasspath.Value = 0
    chk_classpath.Value = 0
    chk_d = 0
    chk_deprecation.Value = 0
    chk_encoding.Value = 0
    chk_extdirs.Value = 0
    chk_g.Value = 0
    chk_nowarn.Value = 0
    chk_O.Value = 0
    chk_sourcepath.Value = 0
    chk_target.Value = 0
    chk_verbose.Value = 0
    Dim i As Byte
    For i = 0 To cmd_paths.Count - 1
        txt_paths(i).Text = ""
        cmd_paths(i).Enabled = False
        txt_paths(i).Enabled = False
    Next i
    txt_g.Text = ""
    txt_encoding.Text = ""
    txt_target.Text = ""
    txt_g.Enabled = False
    txt_encoding.Enabled = False
    txt_target.Enabled = False
    txt_compile.Text = Chr(34) & editor.javapath & "\javac" & Chr(34) & " " & Chr(34) & editor.sfile & Chr(34)
End Sub

Private Sub txt_encoding_Change()
    Call changetext
End Sub
Private Sub txt_g_Change()
    Call changetext
End Sub

Private Sub txt_paths_Change(Index As Integer)
    Call changetext
End Sub

Private Sub txt_target_Change()
    Call changetext
End Sub
