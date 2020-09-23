VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Env_settings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Environment Settings"
   ClientHeight    =   3285
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5640
   Icon            =   "Environment_Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_dbcolor 
      Caption         =   "Default"
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
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Click to Set the Default Background Color"
      Top             =   240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   600
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Preview 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   13
      Text            =   "Java"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Preview 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Preview 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmd_dfont 
      Caption         =   "Default"
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
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "Click to set Default Font"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmd_font 
      Caption         =   "Font..."
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
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Click to set the Font"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmd_dfcolor 
      Caption         =   "Default"
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
      Left            =   4200
      TabIndex        =   5
      ToolTipText     =   "Click to Set the Default Foreground Color"
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmd_fcolor 
      Caption         =   "Fore Color..."
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
      Left            =   2760
      TabIndex        =   4
      ToolTipText     =   "Click to set Foreground Color"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmd_bcolor 
      Caption         =   "Back Color..."
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
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "Click To Set Background Color"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Test 
      Caption         =   "&Test"
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
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Click to Test Changes"
      Top             =   2400
      Width           =   1095
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
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Click to Cancel"
      Top             =   2400
      Width           =   1095
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
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Click to Save Changes"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Font"
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
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Foreground Color:"
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
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Background Color:"
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
      TabIndex        =   8
      Top             =   360
      Width           =   1590
   End
End
Attribute VB_Name = "Env_settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is used to set background and foreground colors and as well as font name,sytle

Option Explicit
Dim bcolor As Long, bprevcolor As Long, fcolor As Long, fprevcolor As Long
Dim isbold As Boolean
Dim isitalic As Boolean
Dim isStrikethru As Boolean
Dim isUnderline As Boolean
Dim previsbold As Boolean
Dim previsitalic As Boolean
Dim previsStrikethru As Boolean
Dim previsUnderline As Boolean
Dim LastCharacter As Integer, prevsize As Integer, size As Integer
Dim prevfname As String, fname As String
'Restore the previous settings
Private Sub CancelButton_Click()
    editor.writepad.SelStart = 0
    editor.writepad.SelLength = Len(editor.writepad.TextRTF)
    editor.writepad.SelColor = fprevcolor
    editor.writepad.BackColor = bprevcolor
    editor.writepad.SelFontSize = prevsize
    editor.writepad.SelBold = previsbold
    editor.writepad.SelItalic = previsitalic
    editor.writepad.SelStrikeThru = previsStrikethru
    editor.writepad.SelUnderline = previsUnderline
    editor.writepad.SelFontName = prevfname
    editor.picLines.Font = prevfname
    editor.picLines.FontSize = prevsize
    editor.writepad.SelStart = LastCharacter
    Unload Me
End Sub
'This procedure is invoked to select a background color
Private Sub cmd_bcolor_Click()
    cdg1.ShowColor
    bcolor = cdg1.Color
    Preview(0).BackColor = bcolor
End Sub
'This procedure is invoked to select a default background color
    Private Sub cmd_dbcolor_Click()
    bcolor = 16777215
    Preview(0).BackColor = bcolor
End Sub
'This procedure is invoked to select a default foreground color
Private Sub cmd_dfcolor_Click()
    fcolor = 0
    Preview(1).BackColor = fcolor
End Sub
'This procedure is invoked to select a default font
Private Sub cmd_dfont_Click()
    Preview(2).FontBold = False
    Preview(2).FontItalic = False
    Preview(2).FontName = "Courier New"
    Preview(2).FontSize = 12
End Sub
'This procedure is invoked to select a foreground color
Private Sub cmd_fcolor_Click()
    cdg1.ShowColor
    fcolor = cdg1.Color
    Preview(1).BackColor = fcolor
End Sub
'This procedure shows font dialog box
Private Sub cmd_font_Click()
    On Error GoTo cancel
        cdg1.Flags = cdlCFBoth Or cdlCFEffects Or cdlCFLimitSize
        cdg1.FontName = "FixedSys"
        cdg1.Min = 12
        cdg1.Max = 24
        cdg1.ShowFont
        Preview(2).FontBold = cdg1.FontBold
        Preview(2).FontItalic = cdg1.FontItalic
        Preview(2).FontStrikethru = cdg1.FontStrikethru
        Preview(2).FontUnderline = cdg1.FontUnderline
        Preview(2).FontName = cdg1.FontName
        Preview(2).FontSize = cdg1.FontSize
cancel:
End Sub

Private Sub Form_Load()
    LastCharacter = Len(editor.writepad.Text)
    bprevcolor = editor.writepad.BackColor
    fprevcolor = editor.writepad.SelColor
    previsbold = editor.writepad.SelBold
    previsitalic = editor.writepad.SelItalic
    previsStrikethru = editor.writepad.SelStrikeThru
    previsUnderline = editor.writepad.SelUnderline
    prevfname = editor.writepad.SelFontName
    prevsize = editor.writepad.SelFontSize
    Call Prvw
End Sub
'This procedure saves new settings into editor.ini and change the editor
Private Sub OKButton_Click()
    Call Test_Click
    Call modifyeditorini   'Saving settings in Editor.ini file
    Unload Me
End Sub
'Change all the new settings
Private Sub Test_Click()
    editor.writepad.SelStart = 0
    editor.writepad.SelLength = Len(editor.writepad.TextRTF)
    editor.writepad.SelColor = Preview(1).BackColor
    editor.writepad.BackColor = Preview(0).BackColor
    editor.writepad.SelFontName = Preview(2).FontName
    editor.writepad.SelFontSize = Preview(2).FontSize
    editor.picLines.Font = Preview(2).FontName
    editor.picLines.FontSize = Preview(2).FontSize
    editor.writepad.SelBold = Preview(2).FontBold
    editor.writepad.SelItalic = Preview(2).FontItalic
    editor.writepad.SelStrikeThru = Preview(2).FontStrikethru
    editor.writepad.SelUnderline = Preview(2).FontUnderline
    editor.writepad.SelStart = LastCharacter
End Sub
'Set preview text box
Private Sub Prvw()
    Preview(0).BackColor = bprevcolor
    Preview(1).BackColor = fprevcolor
    Preview(2).FontBold = previsbold
    Preview(2).FontItalic = previsitalic
    Preview(2).FontStrikethru = previsStrikethru
    Preview(2).FontUnderline = previsUnderline
    Preview(2).FontName = prevfname
    Preview(2).FontSize = prevsize
End Sub
