VERSION 5.00
Begin VB.Form run_with 
   Caption         =   "Java - The Java Application Launcher"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   Icon            =   "run_with.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_Xdebug 
      Caption         =   "-Xdebug"
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
      Left            =   6360
      TabIndex        =   44
      ToolTipText     =   "Enable Remote Debugging"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CheckBox chk_Xcheck 
      Caption         =   "-Xcheck:jni"
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
      Left            =   6360
      TabIndex        =   43
      ToolTipText     =   "Perform additional checks for JNI functions"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CheckBox chk_Xrs 
      Caption         =   "-Xrs"
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
      Left            =   6360
      TabIndex        =   42
      ToolTipText     =   "Reduce the use of OS signals"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CheckBox chk_Xnoclassgc 
      Caption         =   "-Xnoclassgc"
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
      Left            =   6360
      TabIndex        =   41
      ToolTipText     =   "Disables Class garbage collection"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txt_Xrunhprof 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   40
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox txt_Xmx 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   39
      ToolTipText     =   "Write the Maximum java heap size"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txt_Xms 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   38
      ToolTipText     =   "write the initial java heap size"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txt_Xverify 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   37
      ToolTipText     =   "Write either :none|all|remote"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmd_Xbootclasspath 
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
      TabIndex        =   36
      ToolTipText     =   "Click to select the Xbootclasspath(s)"
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox txt_Xbootclasspath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   35
      ToolTipText     =   "Write or select the path(s)"
      Top             =   4320
      Width           =   3135
   End
   Begin VB.CheckBox chk_Xrunhprof 
      Caption         =   "-Xrunhprof[:help]|[:<option>=<value>, ... ]"
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
      TabIndex        =   34
      ToolTipText     =   "Perform heap, cpu or monitor profiling"
      Top             =   5760
      Width           =   3855
   End
   Begin VB.CheckBox chk_Xmx 
      Caption         =   "-Xmx<size>"
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
      TabIndex        =   33
      ToolTipText     =   "Set Maximum Java Heap size"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CheckBox chk_Xms 
      Caption         =   "-Xms<size>"
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
      TabIndex        =   32
      ToolTipText     =   "Set initial Java heap size"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CheckBox chk_Xverify 
      Caption         =   "-Xverify[:none|all|remote]"
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
      TabIndex        =   31
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CheckBox chk_Xbootclasspath 
      Caption         =   "-Xbootclasspath:<paths>"
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
      TabIndex        =   30
      ToolTipText     =   "Search path for Bootstrap classes and resoursed"
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmd_classpath 
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
      TabIndex        =   29
      ToolTipText     =   "Click to select Classpath(s)"
      Top             =   3000
      Width           =   375
   End
   Begin VB.CheckBox chk_X 
      Caption         =   "-X"
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
      Left            =   6480
      TabIndex        =   28
      ToolTipText     =   "Print help on Non-stardard Options"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox chk_help 
      Caption         =   "-help"
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
      Left            =   6480
      TabIndex        =   27
      ToolTipText     =   "Prints help message"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox chk_version 
      Caption         =   "-version"
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
      Left            =   6480
      TabIndex        =   26
      ToolTipText     =   "Print Product Version"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txt_verbose 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   25
      ToolTipText     =   "Write either :class|gc|jni"
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txt_D 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   24
      ToolTipText     =   "Every property should be preceded by -D (except the first one)"
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox txt_classpath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   23
      ToolTipText     =   "Write or Select Path(s)"
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CheckBox chk_verbose 
      Caption         =   "-verbose[:class | gc | jni]"
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
      TabIndex        =   22
      ToolTipText     =   "Enable Verbose Output"
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CheckBox chk_D 
      Caption         =   "-D<name>=<value>"
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
      ToolTipText     =   "Set a System Property"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CheckBox chk_classpath 
      Caption         =   "-classpath <paths>"
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
      TabIndex        =   20
      ToolTipText     =   "Search path for application classes and resourses"
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox txt_execute 
      Height          =   285
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Modify here if necessary"
      Top             =   2280
      Width           =   6255
   End
   Begin VB.TextBox txt_argument 
      Height          =   285
      Left            =   4560
      TabIndex        =   17
      ToolTipText     =   "Specify Arguments if any..."
      Top             =   1920
      Width           =   1695
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
      Left            =   6600
      TabIndex        =   16
      ToolTipText     =   "Click to get the help on Java interpreter"
      Top             =   2280
      Width           =   975
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
      Left            =   6600
      TabIndex        =   15
      ToolTipText     =   "Click to Reset all the options"
      Top             =   1560
      Width           =   975
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
      Left            =   6600
      TabIndex        =   14
      ToolTipText     =   "Click to Cancel"
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
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
      Left            =   6600
      TabIndex        =   13
      ToolTipText     =   "Click to run the class or jar file"
      Top             =   240
      Width           =   975
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
      Left            =   6000
      TabIndex        =   12
      ToolTipText     =   "Click to select jar file"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmd_class 
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
      Left            =   6000
      TabIndex        =   11
      ToolTipText     =   "Click to Select the class file"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txt_class 
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      ToolTipText     =   "Write or Select the Class file"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txt_jar 
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      ToolTipText     =   "Write or select jar file"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Application Launcher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton opt_javaclass 
         Caption         =   "oldjavaw [ options ] class  [ argument... ]"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   3855
      End
      Begin VB.OptionButton opt_javaclass 
         Caption         =   "oldjava [ options ] class [ argument... ]"
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
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   3735
      End
      Begin VB.OptionButton opt_javajar 
         Caption         =   "javaw [ options ] -jar file.jar  [ argument... ]"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton opt_javaclass 
         Caption         =   "javaw  [ options ] class [ argument... ]"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   3615
      End
      Begin VB.OptionButton opt_javajar 
         Caption         =   "java [ options ] -jar file.jar  [ argument... ]"
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.OptionButton opt_javaclass 
         Caption         =   "java [ options ] class [ argument... ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Argument..."
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
      Left            =   4560
      TabIndex        =   45
      Top             =   1680
      Width           =   990
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7800
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "file.jar"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "class"
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
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "run_with"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sel As Byte ' 1 for selecting class file
                   ' 2 for selecting jar file
                   ' 3 for setting classpath
                   ' 4 for setting Xbootclasspath
'Procedure which sets txt_execute textbox
Private Sub changetext()
    txt_execute.Text = Chr(34) & editor.javapath
            
                ' selection of java command
    If opt_javaclass.Item(0).Value = True Then txt_execute.Text = txt_execute.Text & "\java" & Chr(34) & " "
    If opt_javaclass.Item(1).Value = True Then txt_execute.Text = txt_execute.Text & "\javaw" & Chr(34) & " "
    If opt_javaclass.Item(2).Value = True Then txt_execute.Text = txt_execute.Text & "\oldjava" & Chr(34) & " "
    If opt_javaclass.Item(3).Value = True Then txt_execute.Text = txt_execute.Text & "\oldjavaw" & Chr(34) & " "
    If opt_javajar.Item(0).Value = True Then txt_execute.Text = txt_execute.Text & "\java" & Chr(34) & " -jar "
    If opt_javajar.Item(1).Value = True Then txt_execute.Text = txt_execute.Text & "\javaw" & Chr(34) & " -jar "
    
                ' selection of options
                
    If chk_classpath.Value = 1 Then txt_execute.Text = txt_execute.Text & "-classpath " & txt_classpath.Text & " "
    If chk_d.Value = 1 Then txt_execute.Text = txt_execute.Text & "-D " & txt_d.Text & " "
    If chk_verbose.Value = 1 Then txt_execute.Text = txt_execute.Text & "-verbose " & txt_verbose.Text & " "
    If chk_version.Value = 1 Then txt_execute.Text = txt_execute.Text & "-version "
    If chk_help.Value = 1 Then txt_execute.Text = txt_execute.Text & "-help "
    If chk_X.Value = 1 Then txt_execute.Text = txt_execute.Text & "-X "
    
                ' selection of Non-Standard Options
                '-Xbootclasspath: -Xnoclassgc -Xms -Xmx -Xrs -Xcheck:jni -Xrunhprof -Xdebug -Xverify
    If chk_Xbootclasspath.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xbootclasspath " & txt_Xbootclasspath.Text & " "
    If chk_Xnoclassgc.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xnoclassgc" & " "
    If chk_Xms.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xms " & txt_Xms.Text & " "
    If chk_Xmx.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xmx " & txt_Xmx.Text & " "
    If chk_Xrs.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xrs "
    If chk_Xcheck.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xcheck:jni "
    If chk_Xrunhprof.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xrunhprof " & txt_Xrunhprof.Text & " "
    If chk_Xdebug.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xdebug "
    If chk_Xverify.Value = 1 Then txt_execute.Text = txt_execute.Text & "-Xverify " & txt_Xverify & " "
    If opt_javajar.Item(0).Value = True Or opt_javajar.Item(1).Value = True Then txt_execute.Text = txt_execute.Text & txt_jar.Text & " "
    If txt_class.Text <> "" Then txt_execute.Text = txt_execute.Text & txt_class.Text
    If txt_argument.Text <> "" Then txt_execute.Text = txt_execute.Text & " " & txt_argument
End Sub
Private Sub chk_classpath_Click()
    If chk_classpath.Value = 1 Then
        txt_classpath.Enabled = True
        cmd_classpath.Enabled = True
        txt_classpath.SetFocus
    Else
        txt_classpath.Enabled = False
        cmd_classpath.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_d_Click()
    If chk_d.Value = 1 Then
        txt_d.Enabled = True
        txt_d.SetFocus
    Else
        txt_d.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_help_Click()
    Call changetext
End Sub

Private Sub chk_verbose_Click()
    If chk_verbose.Value = 1 Then
        txt_verbose.Enabled = True
        txt_verbose.SetFocus
    Else
        txt_verbose.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_version_Click()
    Call changetext
End Sub

Private Sub chk_X_Click()
    Call changetext
End Sub

Private Sub chk_Xbootclasspath_Click()
    If chk_Xbootclasspath.Value = 1 Then
        txt_Xbootclasspath.Enabled = True
        cmd_Xbootclasspath.Enabled = True
        txt_Xbootclasspath.SetFocus
    Else
        txt_Xbootclasspath.Enabled = False
        cmd_Xbootclasspath.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_Xcheck_Click()
    Call changetext
End Sub

Private Sub chk_Xdebug_Click()
    Call changetext
End Sub

Private Sub chk_Xms_Click()
    If chk_Xms.Value = 1 Then
        txt_Xms.Enabled = True
        txt_Xms.SetFocus
    Else
        txt_Xms.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_Xmx_Click()
    If chk_Xmx.Value = 1 Then
        txt_Xmx.Enabled = True
        txt_Xmx.SetFocus
    Else
        txt_Xmx.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_Xnoclassgc_Click()
    Call changetext
End Sub

Private Sub chk_Xrs_Click()
    Call changetext
End Sub

Private Sub chk_Xrunhprof_Click()
    If chk_Xrunhprof.Value = 1 Then
        txt_Xrunhprof.Enabled = True
        txt_Xrunhprof.SetFocus
    Else
        txt_Xrunhprof.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_Xverify_Click()
    If chk_Xverify.Value = 1 Then
        txt_Xverify.Enabled = True
        txt_Xverify.SetFocus
    Else
        txt_Xverify.Enabled = False
    End If
    Call changetext
End Sub

Private Sub cmd_Cancel_Click()
    editor.SetFocus
    Unload Me
End Sub

Private Sub cmd_class_Click()
    editor.calledfrom = fromRun_with 'set called from to 5
    'to indicated that PathFile form is called from run_with form
    sel = 1
    PathFile.File1.Visible = True
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Pattern = "*.class"
    PathFile.Label1.Caption = "Select the class File :"
    PathFile.Show vbModal
    Call changetext
    PathFile.File1.Pattern = "*.*"
End Sub

Private Sub cmd_classpath_Click()
    sel = 3
    editor.calledfrom = fromRun_with 'set called from to 5
    'to indicated that PathFile form is called from run_with form
    PathFile.File1.Visible = False
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.Label1.Caption = "Select the Directory :"
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub cmd_jar_Click()
    editor.calledfrom = fromRun_with 'set called from to 5
    'to indicated that PathFile form is called from run_with form
    sel = 2
    PathFile.File1.Visible = True
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Pattern = "*.jar"
    PathFile.Label1.Caption = "Select the jar File :"
    PathFile.Show vbModal
    Call changetext
    PathFile.File1.Pattern = "*.*"
End Sub

Private Sub cmd_OK_Click()
    'Creates Editor.Bat
    CreateEditorbat "@echo off" & vbCrLf & txt_execute.Text & vbCrLf & "pause"
    Dim cmdLine As String
    cmdLine = CurDir() & "\editor.bat"

    'Change to Class's Directory
    Call changeDirectory(getDrive(editor.sfile), getPath(editor.sfile))
    
    'Run Java Class File
    RunShell cmdLine, vbMaximizedFocus
    
    'Change to Application Path
    Call changeapppath
End Sub
'Procedure to reset all the options
Private Sub cmd_reset_Click()
    opt_javaclass.Item(0).Value = True
    txt_class.Text = "Noname"
    txt_jar.Text = "Noname.jar"
    txt_argument.Text = ""
    chk_classpath.Value = 1: txt_classpath.Text = Chr(34) & editor.defaultpath & Chr(34) & ";": cmd_classpath.Enabled = True
    chk_d.Value = 0: txt_d.Enabled = False: txt_d.Text = ""
    chk_verbose.Value = 0: txt_verbose = ""
    chk_version.Value = 0: chk_help.Value = 0: chk_X.Value = 0
    chk_Xbootclasspath.Value = 0: txt_Xbootclasspath.Text = "": cmd_Xbootclasspath.Enabled = False
    chk_Xverify.Value = 0: txt_Xverify.Text = ""
    chk_Xms.Value = 0: txt_Xms.Text = ""
    chk_Xmx.Value = 0: txt_Xmx.Text = ""
    chk_Xrunhprof.Value = 0: txt_Xrunhprof.Text = ""
    chk_Xnoclassgc.Value = 0: chk_Xrs.Value = 0: chk_Xcheck.Value = 0: chk_Xdebug.Value = 0
End Sub

Private Sub cmd_Xbootclasspath_Click()
    PathFile.File1.Visible = False
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.Label1.Caption = "Select the Directory :"
    editor.calledfrom = fromRun_with 'set called from to 5
    'to indicated that PathFile form is called from run_with form
    sel = 4
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub Form_Unload(cancel As Integer)
    FormStayOnTop Me, False
End Sub
Private Sub Form_Load()
    txt_jar.Text = "Noname.jar": txt_jar.Enabled = False: cmd_jar.Enabled = False
    txt_Xbootclasspath.Enabled = False: cmd_Xbootclasspath.Enabled = False
    txt_class.Text = removeExt(getFile(editor.sfile))
    txt_classpath.Text = Chr(34) & getPath(editor.sfile) & Chr(34) & ";"
    txt_execute.Text = Chr(34) & editor.javapath & Chr(34) & " -classpath " & Chr(34) & getPath(editor.sfile) & Chr(34) & ";" & txt_class.Text
End Sub

Private Sub Form_Resize()
    editor.SetFocus
End Sub

Private Sub opt_javaclass_Click(Index As Integer)
    cmd_class.Enabled = True
    txt_class.Enabled = True: txt_class.SetFocus
    cmd_jar.Enabled = False
    txt_jar.Enabled = False
    Call changetext
End Sub

Private Sub opt_javajar_Click(Index As Integer)
    cmd_jar.Enabled = True
    txt_jar.Enabled = True: txt_jar.SetFocus
    cmd_class.Enabled = False
    txt_class.Enabled = Falseub
    Call changetext
End Sub

Private Sub txt_argument_Change()
    Call changetext
End Sub

Private Sub txt_class_Change()
    Call changetext
End Sub

Private Sub txt_classpath_Change()
    Call changetext
End Sub

Private Sub txt_d_Change()
    Call changetext
End Sub

Private Sub txt_jar_Change()
    Call changetext
End Sub

Private Sub txt_verbose_Change()
    Call changetext
End Sub

Private Sub txt_Xbootclasspath_Change()
    Call changetext
End Sub

Private Sub txt_Xms_Change()
    Call changetext
End Sub

Private Sub txt_Xmx_Change()
    Call changetext
End Sub

Private Sub txt_Xrunhprof_Change()
    Call changetext
End Sub

Private Sub txt_Xverify_Change()
    Call changetext
End Sub
