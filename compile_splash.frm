VERSION 5.00
Begin VB.Form compile_splash 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1740
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "compile_splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "compile_splash.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   1740
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compiling "
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
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait While Compilation Completes..."
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5235
   End
End
Attribute VB_Name = "compile_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is loaded when user compiles a java program
'this form is just a splash screen
Private Sub Form_Activate()

    Dim FileDrive As String, FilePath As String, cmdLine As String

    FileDrive = getDrive(editor.sfile)
    FilePath = getPath(editor.sfile)
    cmdLine = CurDir() & "\Redirect.exe"

    'Change to Class's Directory
    Call changeDirectory(FileDrive, FilePath)

    'Run the Redirect.exe
    RunShell cmdLine, vbHide

    'Change to Applications Path Directory
    Call changeapppath

    If editor.calledfrom <> fromJavadoc Then Unload Me
    
    'Show the Javac Compilation Output
    Output.Show vbModal

End Sub

