VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Java Editor"
   ClientHeight    =   4230
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7890
   Icon            =   "abouteditor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2919.621
   ScaleMode       =   0  'User
   ScaleWidth      =   7409.118
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5400
      TabIndex        =   0
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Image Image2 
      Height          =   2565
      Left            =   5400
      Picture         =   "abouteditor.frx":0442
      Top             =   120
      Width           =   2160
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7437.29
      Y1              =   1987.827
      Y2              =   1987.827
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "abouteditor.frx":19EE
      Top             =   240
      Width           =   480
   End
   Begin VB.Label mail 
      AutoSize        =   -1  'True
      Caption         =   "nihar_d_maniyar@yahoo.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      MouseIcon       =   "abouteditor.frx":1E30
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3600
      Width           =   2100
   End
   Begin VB.Label Label2 
      Caption         =   "Email : "
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "All Right Reserved to Nihar Dinesh Maniyar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   5
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Java Editor for Windows Operating Systems"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   3915
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Java Editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   1380
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   4
      Top             =   780
      Width           =   990
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      Caption         =   "NIHAR DINESH MANIYAR., Msc (Comp.sci)"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   255
      TabIndex        =   2
      Top             =   3105
      Width           =   3165
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form is used to show about information

' This is the button when pressed the form is unloaded
Private Sub cmdOK_Click()
    Unload Me
End Sub

'This label procedure is invoked when a user wants to mail to the creater
'of this editor
Private Sub mail_Click()
    Dim ret 'ret value is used to store the value returned by Shell() function
    ret = Shell("start mailto:" & "nihar.maniyar@polaris.co.in" & "?subject=" & "Java+Editor+1.0" & "", 0)
End Sub
