VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Output 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output"
   ClientHeight    =   6225
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   8205
   ClipControls    =   0   'False
   Icon            =   "Output.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4296.605
   ScaleMode       =   0  'User
   ScaleWidth      =   7704.919
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtf_output 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9551
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1e7
      TextRTF         =   $"Output.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton OutputOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
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
      Height          =   345
      Left            =   3600
      TabIndex        =   0
      ToolTipText     =   "Click to Hide the Output"
      Top             =   5640
      Width           =   1260
   End
End
Attribute VB_Name = "Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'It opens Error file i.e., is Output and shows in the RTF box
Private Sub Form_Load()
    If editor.compiled = True Then
        Call changeapppath    'Changes Current Working Directory to Application Path
        Open "output" For Input As #1
        Do Until EOF(1)
            Line Input #1, newline
            Output.rtf_output.Text = Output.rtf_output.Text + newline + vbCrLf
        Loop
        Close #1
    End If
End Sub
'Procedure to unload the form
Private Sub OutputOK_Click()
    Unload Me
End Sub
