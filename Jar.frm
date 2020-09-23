VERSION 5.00
Begin VB.Form Jar 
   Caption         =   "Jar - The Java Archive Tool"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   Icon            =   "Jar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opt_x 
      Caption         =   "-x [files...]"
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
      Left            =   2400
      TabIndex        =   24
      ToolTipText     =   "Extract named (or all) files from the archieve"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton opt_u 
      Caption         =   "-u"
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
      TabIndex        =   23
      ToolTipText     =   "Update existing archive"
      Top             =   1200
      Width           =   495
   End
   Begin VB.OptionButton opt_t 
      Caption         =   "-t"
      CausesValidation=   0   'False
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
      TabIndex        =   22
      ToolTipText     =   "List table of contents for archive"
      Top             =   1200
      Width           =   495
   End
   Begin VB.OptionButton opt_c 
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
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Create New Archive"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txt_c 
      Height          =   285
      Left            =   2040
      TabIndex        =   19
      ToolTipText     =   "Write or select the destination or input directory"
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txt_m 
      Height          =   285
      Left            =   2040
      TabIndex        =   18
      ToolTipText     =   "write or select manifest file"
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txt_f 
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      ToolTipText     =   "Write the jar file to create or select an existing jar file"
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CheckBox chk_Cdir 
      Caption         =   "-C <in, output dir>"
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
      TabIndex        =   16
      ToolTipText     =   "Change to specified directory either for input or output"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CheckBox chk_M 
      Caption         =   "-M"
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
      ToolTipText     =   "Do not create manifest files for the entries"
      Top             =   3120
      Width           =   735
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
      Height          =   255
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "store only; Use no ZIP compression"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox chk_manifest 
      Caption         =   "-m <manifest-file>"
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
      TabIndex        =   13
      ToolTipText     =   "Include manifest information from manifest file"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CheckBox chk_f 
      Caption         =   "-f <jar-file>"
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
      TabIndex        =   12
      ToolTipText     =   "Specify archieve filename"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CheckBox chk_v 
      Caption         =   "-v"
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
      TabIndex        =   11
      ToolTipText     =   "Generate verbose output on standard output"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txt_files 
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      ToolTipText     =   "write the filenames to be extracted seperated by space"
      Top             =   1170
      Width           =   2055
   End
   Begin VB.CommandButton cmd_f 
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
      Left            =   5400
      TabIndex        =   8
      ToolTipText     =   "Click to select the jar file"
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmd_m 
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
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "click to select manifest file"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmd_c 
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
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Click to select the destination or input directory"
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmd_jar 
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
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Click to select files from input directory one by one"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txt_jar 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Modify here if necessary"
      Top             =   360
      Width           =   5655
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
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "click to create|display|update|extract jar file"
      Top             =   4320
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
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Click to cancel"
      Top             =   4320
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
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Click to reset all options"
      Top             =   4320
      Width           =   1215
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
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "Click to get help on jar file"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "jar {ctxu}[vfm0M] [jar-file] [manifest-file} [-C dir] files..."
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
      TabIndex        =   20
      Top             =   120
      Width           =   4605
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
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   780
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   120
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "Jar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sel As Byte 'variable indicates which button was clicked
Public txtjar As String 'variable to store the total jar command
'Procedure to change the text box
Private Sub changetext()
    txt_jar.Text = Chr(34) & editor.javapath & "\jar" & Chr(34) & " "
    If opt_c.Value = True Then txt_jar.Text = txt_jar.Text & "c"
    If opt_t.Value = True Then txt_jar.Text = txt_jar.Text & "t"
    If opt_u.Value = True Then txt_jar.Text = txt_jar.Text & "u"
    If opt_x.Value = True Then txt_jar.Text = txt_jar.Text & "x"
    If chk_v.Value = 1 Then txt_jar.Text = txt_jar.Text & "v"
    If chk_f.Value = 1 Then txt_jar.Text = txt_jar.Text & "f"
    If chk_manifest.Value = 1 Then txt_jar.Text = txt_jar.Text & "m"
    If chk_O.Value = 1 Then txt_jar.Text = txt_jar.Text & "0"
    If chk_M.Value = 1 Then txt_jar.Text = txt_jar.Text & "M"
    If chk_manifest.Value = 1 And txt_m.Text <> "" Then txt_jar.Text = txt_jar.Text & " " & txt_m.Text
    If chk_f.Value = 1 And txt_f.Text <> "" Then txt_jar.Text = txt_jar.Text & " " & txt_f.Text
    If opt_x.Value = 0 And chk_Cdir.Value = 1 Then txt_jar.Text = txt_jar.Text & " -C"
    If opt_x.Value = 0 And chk_Cdir.Value = 1 And txt_c.Text <> "" Then txt_jar.Text = txt_jar.Text & " " & txt_c.Text
    If opt_x.Value = True And txt_files.Text <> "" Then txt_jar.Text = txt_jar.Text & " " & txt_files.Text
    If cmd_jar.Enabled = True And txtjar <> "" Then txt_jar.Text = txt_jar.Text & txtjar
End Sub
Private Sub chk_Cdir_Click()
    If chk_Cdir.Value = 1 Then
        txt_c.Enabled = True
        cmd_c.Enabled = True
    Else
        cmd_c.Enabled = False
        txt_c.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_f_Click()
    If chk_f.Value = 1 Then
        txt_f.Enabled = True
        If opt_x.Value = True Or opt_u.Value = True Then
            cmd_f.Enabled = True
            txt_f.Enabled = True
            txt_f.SetFocus
        End If
    Else
        cmd_f.Enabled = False
        txt_f.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_M_Click()
    Call changetext
End Sub

Private Sub chk_manifest_Click()
    If chk_manifest.Value = 1 Then
        txt_m.Enabled = True
        cmd_m.Enabled = True
        txt_m.SetFocus
    Else
        cmd_m.Enabled = False
        txt_m.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_O_Click()
    Call changetext
End Sub

Private Sub chk_v_Click()
    Call changetext
End Sub

Private Sub cmd_c_Click()
    sel = 4
    editor.calledfrom = fromJar 'set called from to 7
    'to indicated that PathFile form is called from Jar form
    PathFile.File1.Visible = False
    PathFile.Label1.Caption = "Select a Directory :"
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub cmd_f_Click()
    sel = 2
    editor.calledfrom = fromJar 'set called from to 7
    'to indicated that PathFile form is called from Jar form
    PathFile.File1.Visible = True
    PathFile.Label1.Caption = "Select a Jar File :"
    PathFile.File1.Pattern = "*.jar"
    PathFile.Show vbModal
    Call changetext
    PathFile.File1.Pattern = "*.*"
End Sub

Private Sub cmd_jar_Click()
    editor.calledfrom = fromJar 'set called from to 7
    'to indicated that PathFile form is called from Jar form
    sel = 1
    PathFile.opt_File.Visible = True
    PathFile.opt_File.Value = True
    PathFile.opt_all.Visible = True
    PathFile.Opt_Directory.Visible = True
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub cmd_m_Click()
    sel = 3
    editor.calledfrom = fromJar
    PathFile.File1.Visible = True
    PathFile.Label1.Caption = "Select a Manifest File :"
    PathFile.Show vbModal
    Call changetext
End Sub
'Procedure that creates the editor.bat file and executes it
Private Sub cmd_OK_Click()
    CreateEditorbat txt_jar.Text & vbCrLf & "pause"
    Dim temp As Double
    temp = Shell("editor.bat ", vbMaximizedFocus)
End Sub
'Procedure to reset the values
Private Sub cmd_reset_Click()
    chk_v.Value = 0
    chk_O.Value = 0
    chk_M.Value = 0
    chk_manifest.Value = 0
    chk_f.Value = 0
    txt_files.Text = ""
    Call Form_Load
End Sub

Private Sub Form_Load()
    txtjar = ""
    txt_jar.Text = Chr(34) & editor.javapath & "\jar " & Chr(34) & " c -C " & Chr(34) & editor.defaultpath & Chr(34)
    opt_c.Value = True
    chk_Cdir.Value = 1: txt_c.Text = Chr(34) & editor.defaultpath & Chr(34)
    txt_f.Text = "": txt_f.Enabled = False: cmd_f.Enabled = False
    txt_m.Text = "": txt_m.Enabled = False: cmd_m.Enabled = False
    txt_files.Enabled = False
    Call changetext
End Sub

Private Sub opt_c_Click()
    txt_files.Enabled = False
    txtjar = ""
    cmd_f.Enabled = False
    chk_Cdir.Value = 1
    chk_Cdir.Caption = "-C <in,output dir>"
    chk_manifest.Enabled = True
    chk_O.Enabled = True
    chk_M.Enabled = True
    chk_Cdir.Enabled = True
    chk_Cdir.Value = 1
    cmd_jar.Enabled = True
    Call changetext
End Sub

Private Sub opt_t_Click()
    txt_files.Enabled = False
    chk_v.Value = 0
    chk_f.Value = 0
    txt_f.Text = ""
    chk_manifest.Value = 0
    chk_manifest.Enabled = False
    chk_O.Enabled = False
    chk_O.Value = 0
    chk_M.Enabled = False
    chk_M.Value = 0
    chk_Cdir.Caption = "-C <dir>"
    chk_Cdir.Value = 0
    chk_Cdir.Enabled = False
    cmd_jar.Enabled = False
    Call changetext
End Sub

Private Sub opt_u_Click()
    If chk_f.Value = 1 Then
        txt_f.Enabled = True
        cmd_f.Enabled = True
    End If
    txt_files.Enabled = False
    txtjar = ""
    chk_Cdir.Caption = "-C <input dir>"
    cmd_jar.Enabled = True
    chk_Cdir.Value = 1
    chk_manifest.Enabled = True
    chk_O.Enabled = True
    chk_M.Enabled = True
    chk_Cdir.Enabled = True
    chk_Cdir.Value = 1
    Call changetext
End Sub

Private Sub opt_x_Click()
    txt_files.Enabled = True
    txt_files.SetFocus
    If chk_f.Value = 1 Then cmd_f.Enabled = True
    chk_Cdir.Enabled = True
    chk_Cdir.Value = 1
    chk_manifest.Enabled = False
    chk_M.Enabled = False
    chk_O.Enabled = False
    cmd_jar.Enabled = False
    chk_Cdir.Caption = "-C <output dir>"
    Call changetext
End Sub
Private Sub txt_c_Change()
    Call changetext
End Sub

Private Sub txt_f_Change()
    Call changetext
End Sub

Private Sub txt_files_Change()
    Call changetext
End Sub

Private Sub txt_m_Change()
    Call changetext
End Sub
