VERSION 5.00
Begin VB.Form PathFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Path / File"
   ClientHeight    =   4185
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6885
   Icon            =   "PathFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opt_all 
      Caption         =   "*"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.OptionButton opt_File 
      Caption         =   "&File"
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
      Left            =   1080
      TabIndex        =   7
      Top             =   1260
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.OptionButton Opt_Directory 
      Caption         =   "&Directory"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
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
      Left            =   3720
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
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
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   90
   End
End
Attribute VB_Name = "PathFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    'Restore editor settings to the text boxes
    If editor.calledfrom = fromSetup Then ' called from setup
        With setup
        .txt_path.Item(0).Text = editor.javapath
        .txt_path.Item(1).Text = editor.commandpath
        .txt_path.Item(2).Text = editor.browserpath
        .txt_path.Item(3).Text = editor.apipath
        .txt_path.Item(4).Text = editor.defaultpath
        End With
    End If
    Unload Me 'unload the form
End Sub

Private Sub Dir1_Change()
    If File1.Visible Then File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub

'Procedure which saves the settings to the editor.ini file
Private Sub OKButton_Click()

    If File1.FileName = "" And File1.Visible Then
        Dim temp As Double
        temp = MsgBox("Please Select Required File", vbExclamation, "File not selected")
        Exit Sub
    End If

    If editor.calledfrom = fromCompileWithForArguments Then ' called from compile_with and for arguments
        If editor.ind = 0 Or editor.ind = 1 Or editor.ind = 2 Or editor.ind = 3 Then
            If compile_with.txt_paths.Item(editor.ind).Text <> "" Then
                compile_with.txt_paths.Item(editor.ind).Text = compile_with.txt_paths.Item(editor.ind).Text & Chr(34) & Dir1.path & Chr(34) & ";"
            Else
                compile_with.txt_paths.Item(editor.ind).Text = Chr(34) & Dir1.path & Chr(34) & ";"
            End If
        Else
            compile_with.txt_paths.Item(editor.ind).Text = Chr(34) & Dir1.path & Chr(34)
        End If
    End If
    
    If editor.calledfrom = fromSetup Then ' called from setup
        If editor.ind = 0 Then setup.txt_path.Item(0).Text = Dir1.path
        If editor.ind = 4 Then setup.txt_path.Item(4).Text = Dir1.path
        If editor.ind = 1 Or editor.ind = 2 Or editor.ind = 3 Then
            setup.txt_path.Item(editor.ind).Text = File1.path & "\" & File1.FileName
        End If
    End If

    If editor.calledfrom = fromCompileWithForSelectingFile Then 'called from compile_with for selecting file
        compile_with.txt_compile.Text = Chr(34) & editor.javapath & "\javac" & Chr(34) & " " & Chr(34) & File1.path & "\" & File1.FileName & Chr(34)
    End If
    
    If editor.calledfrom = fromAppletviewer Then 'called from appletviewer_with for selecting file
        With appletviewer_with
            .txt_appletviewer.Text = Chr(34) & editor.javapath & "\appletviewer" & Chr(34) & " "
            If .chk_debug.Value = 1 Then .txt_appletviewer.Text = .txt_appletviewer.Text & "-debug "
            If .chk_J.Value = 1 Then
                .txt_appletviewer.Text = .txt_appletviewer.Text & "-j "
                If .txt_J.Text <> "" Then .txt_appletviewer.Text = .txt_appletviewer.Text & .txt_J.Text
            End If
            If .chk_encoding.Value = 1 Then
                .txt_appletviewer.Text = .txt_appletviewer.Text & "-encoding "
                If .txt_encoding.Text <> "" Then .txt_appletviewer.Text = .txt_appletviewer.Text & .txt_encoding.Text
            End If
            editor.sfile = File1.path & "\" & File1.FileName
            .txt_appletviewer.Text = .txt_appletviewer.Text & Chr(34) & File1.FileName & Chr(34)
        End With
    End If

    If editor.calledfrom = fromRun_with Then 'called from run_with for selection
        Dim backupclasspath As String
        backupclasspath = run_with.txt_classpath
        If run_with.sel = 1 Then ' to select class file
            run_with.txt_class.Text = removeExt(File1.FileName)
            run_with.txt_classpath.Text = Chr(34) & File1.path & Chr(34) & ";" & backupclasspath
        ElseIf run_with.sel = 2 Then ' to select jar file
            run_with.txt_jar.Text = File1.FileName
            run_with.txt_classpath.Text = Chr(34) & File1.path & Chr(34) & ";" & backupclasspath
        ElseIf run_with.sel = 3 Then ' to select class path
            run_with.txt_classpath.Text = Chr(34) & Dir1.path & Chr(34) & ";" & backupclasspath
        ElseIf run_with.sel = 4 Then
            Dim backupbootclasspath As String
            backupbootclasspath = run_with.txt_Xbootclasspath
            run_with.txt_Xbootclasspath.Text = Chr(34) & Dir1.path & Chr(34) & ";" & backupbootclasspath
        End If
    End If
    
    If editor.calledfrom = fromJavap Then ' called from javap for selection
        If editor.ind = -1 Then
            Dim backuppath As String
            backuppath = javap.txt_paths(0).Text
            javap.txt_javap.Text = javap.txt_javap.Text & removeExt(File1.FileName)
            If javap.fname <> "Noname" Then
                javap.fname = javap.fname & " " & removeExt(File1.FileName)
            Else
                javap.fname = removeExt(File1.FileName)
            End If
            javap.txt_paths(0).Text = Chr(34) & File1.path & Chr(34) & ";" & backuppath
            If javap.chk_classpath.Value = 0 Then javap.chk_classpath.Value = 1
        Else
            If javap.txt_paths.Item(editor.ind).Text <> "" Then
                javap.txt_paths.Item(editor.ind).Text = javap.txt_paths.Item(editor.ind).Text & Chr(34) & Dir1.path & Chr(34) & ";"
            Else
                javap.txt_paths.Item(editor.ind).Text = Chr(34) & Dir1.path & Chr(34) & ";"
            End If
        End If
    End If

    If editor.calledfrom = fromJar Then 'called from jar for selection
        If Jar.sel = 1 Then
            If opt_File.Value = True Then   ' if Files are being selected then
                Jar.txtjar = Jar.txtjar & " " & File1.FileName
            End If
            If opt_all.Value = True Then
                Jar.txtjar = Jar.txtjar & " *"
            End If
            If Opt_Directory.Value = True Then
                Jar.txtjar = Jar.txtjar & " " & getFile(Dir1.path)
            End If
        ElseIf Jar.sel = 2 Or Jar.sel = 3 Then
            Jar.txt_m.Text = Chr(34) & File1.path & File1.FileName & Chr(34)
        ElseIf Jar.sel = 4 Then
            Jar.txt_c.Text = Chr(34) & Dir1.path & Chr(34)
        End If
    End If

    If editor.calledfrom = fromJavadoc Then 'called from javadoc for selection
        If editor.ind = 2 Or editor.ind = 3 Or editor.ind = 4 Or editor.ind = 5 Then
            If javadoc.txt_paths.Item(editor.ind).Text <> "" Then
                javadoc.txt_paths.Item(editor.ind).Text = javadoc.txt_paths.Item(editor.ind).Text & Chr(34) & Dir1.path & Chr(34) & ";"
            Else
                javadoc.txt_paths.Item(editor.ind).Text = Chr(34) & Dir1.path & Chr(34) & ";"
            End If
        ElseIf editor.ind = 0 Then
            javadoc.txt_paths.Item(editor.ind).Text = Chr(34) & File1.path & "\" & File1.FileName & Chr(34)
        ElseIf editor.ind = 1 Then
            javadoc.txt_paths.Item(editor.ind).Text = Chr(34) & Dir1.path & Chr(34)
        End If
        editor.ind = -1
            
        If javadoc.sel = 1 Then 'javadoc
            If opt_File.Value = True Then
            '
            ElseIf javadoc.sel = 2 Then 'overview
                javadoc.txt_overview.Text = Chr(34) & File1.path & "\" & File1.FileName & Chr(34)
            ElseIf javadoc.sel = 3 Then 'd
                javadoc.txt_d.Text = Chr(34) & Dir1.path & Chr(34)
            ElseIf javadoc.sel = 4 Then 'helpfile
                javadoc.txt_helpfile.Text = Chr(34) & File1.path & "\" & File1.FileName & Chr(34)
            ElseIf javadoc.sel = 5 Then 'stylesheetfile
                javadoc.txt_stylesheetfile.Text = Chr(34) & File1.path & "\" & File1.FileName & Chr(34)
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub opt_all_Click()
    If editor.calledfrom = 8 Then
        Drive1.Visible = True
        Dir1.Visible = True
        File1.Visible = True
    Else
        File1.Visible = False
        Dir1.Visible = False
        Drive1.Visible = False
    End If
End Sub

Private Sub Opt_Directory_Click()
    Drive1.Visible = True
    Dir1.Visible = True
    File1.Visible = False
End Sub

Private Sub opt_File_Click()
    Drive1.Visible = True
    Dir1.Visible = True
    File1.Visible = True
End Sub
