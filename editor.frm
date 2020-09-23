VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form editor 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6525
   Icon            =   "editor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   5730
   ScaleWidth      =   6525
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   19
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "New (Ctrl+N)"
            Object.Tag             =   "1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Open (Ctrl+O)"
            Object.Tag             =   "2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Save(Ctrl+S)"
            Object.Tag             =   "3"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Print (Ctrl+P)"
            Object.Tag             =   "4"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   "5"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cut (Ctrl+X)"
            Object.Tag             =   "6"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Copy (Ctrl+C)"
            Object.Tag             =   "7"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Paste (Ctrl+V)"
            Object.Tag             =   "8"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Undo (Ctrl+Z)"
            Object.Tag             =   "9"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Find (Ctrl+F)"
            Object.Tag             =   "10"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Compile (F9,Ctrl+F9)"
            Object.Tag             =   "11"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Run (F8,Ctrl+F8)"
            Object.Tag             =   "12"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "AppletViewer Run (F8,Ctrl+F8)"
            Object.Tag             =   "13"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Ms-Dos Prompt"
            Object.Tag             =   "14"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Browser (Ctrl+B)"
            Object.Tag             =   "15"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Help (F1)"
            Object.Tag             =   "16"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   6135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   6135
      Begin VB.PictureBox picLines 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   0
         Picture         =   "editor.frx":0442
         ScaleHeight     =   3255
         ScaleWidth      =   375
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin RichTextLib.RichTextBox writepad 
         Height          =   3255
         Left            =   360
         TabIndex        =   2
         Top             =   0
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5741
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         Appearance      =   0
         RightMargin     =   1e7
         TextRTF         =   $"editor.frx":0884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog Cdg1 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   5400
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   582
      SimpleText      =   "+"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "SCRL"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1667
            MinWidth        =   882
            TextSave        =   "9/27/03"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "5:32 PM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":094C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":0E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":13D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":1912
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":1E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":2396
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":28D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":2E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":335C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":346E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":3788
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":3AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":3DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":40D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "editor.frx":43F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_save 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_saveas 
         Caption         =   "S&ave as..."
      End
      Begin VB.Menu mnu_separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_print_setup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnu_print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu_undo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnu_cut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnu_copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnu_paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnu_delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnu_separator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_clear 
         Caption         =   "C&lear"
      End
      Begin VB.Menu mnu_select_all 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnu_separator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_word_wrap 
         Caption         =   "&Word Wrap"
      End
   End
   Begin VB.Menu mnu_search 
      Caption         =   "&Search"
      Begin VB.Menu mnu_find 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu_find_next 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_replace 
         Caption         =   "&Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnu_separator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_go_to 
         Caption         =   "&Go To"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnu_view 
      Caption         =   "&View"
      Begin VB.Menu mnu_tool_bar 
         Caption         =   "&Tool Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_status_bar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_line_number 
         Caption         =   "&Line Number"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_separator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_normal 
         Caption         =   "No&rmal"
      End
      Begin VB.Menu mnu_minimize 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu mnu_maximize 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu mnu_separator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_output 
         Caption         =   "&Output"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnu_project 
      Caption         =   "&Project"
      Begin VB.Menu mnu_compile 
         Caption         =   "&Compile"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnu_compile_with 
         Caption         =   "Co&mpile With..."
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnu_separator8 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_run 
         Caption         =   "Ru&n.."
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnu_run_with 
         Caption         =   "R&un With..."
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnu_separator9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_appletviewer 
         Caption         =   "&Appletviewer"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu_appletviewer_with 
         Caption         =   "Applet&viewer With..."
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnu_separator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_javadoc 
         Caption         =   "&Javadoc..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu_separator11 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_jar 
         Caption         =   "Ja&r..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_separator12 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_javap 
         Caption         =   "Java&p..."
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnu_application 
      Caption         =   "&Application"
      Begin VB.Menu mnu_msdosprompt 
         Caption         =   "Ms-&Dos Prompt"
      End
      Begin VB.Menu mnu_browser_view 
         Caption         =   "&Browser View"
      End
      Begin VB.Menu mnu_api 
         Caption         =   "&API"
      End
   End
   Begin VB.Menu mnu_option 
      Caption         =   "&Option"
      Begin VB.Menu mnu_setup 
         Caption         =   "&Setup..."
      End
      Begin VB.Menu mnu_environment 
         Caption         =   "&Environment..."
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_contents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnu_separator13 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_about_javaedit 
         Caption         =   "&About JavaEdit..."
      End
   End
End
Attribute VB_Name = "editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' variable used to know which form invoked PathFile Form
Public Enum calledFromWhichForm
    fromCompileWithForArguments = 1
    fromSetup = 2
    fromCompileWithForSelectingFile = 3
    fromAppletviewer = 4
    fromRun_with = 5
    fromJavap = 6
    fromJar = 7
    fromJavadoc = 8
End Enum
Public calledfrom As calledFromWhichForm
Public ind As Integer 'ind used as an index between forms
Public strfind As String   'Storing string to be found
Public strreplace As String 'Storing string to be replaced
Public findflags As Integer 'setting RTF find flags
Public foundposition As Integer 'to store foundposition
Public position As Integer 'to store position
Public updown As Boolean 'to search in updirection
Public padbcolor As Long
Public padfcolor As Long
Public padfontsize As Byte
Public padbold As Boolean
Public paditalic As Boolean
Public pfontname
Public javapath As String
Public commandpath As String
Public browserpath As String
Public apipath As String
Public defaultpath As String
Public lineno As Integer
Public col As Integer
Public sfile As String
Public modified As Boolean
Dim isfilesaved As Boolean
Dim movepos As Integer
Public tempchar As String
Public compiled As Boolean

Private Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function SendMessageByRef Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "USER32" (ByVal hwndLock As Long) As Long

Private Const EM_GETSEL As Double = &HB0
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Public Sub GoToLine(ByVal line_number As Integer)
Dim txt As String
Dim i As Integer
Dim pos As Integer

    ' Find the line's position.
    txt = writepad.Text
    pos = 1
    For i = 2 To line_number
        pos = InStr(pos, txt, vbCrLf)
        If pos = 0 Then
            pos = 1
            Exit For
        End If
        pos = pos + 2
    Next i

    ' Go to this position.
    LockWindowUpdate writepad.hWnd
    writepad.SelStart = Len(txt)
    writepad.SelStart = pos - 1
    writepad.SelLength = 0
    LockWindowUpdate 0

    On Error Resume Next
    writepad.SetFocus
End Sub

Private Sub Form_Initialize()
    lineno = 1: col = 1: movepos = 0
    compiled = False
    modified = False
    isfilesaved = False
    Dim lines(12) As String
    On Error GoTo filenotfound
    Call changeapppath    'Changes Current Working Directory to Application Path
    'Open Editor.ini and load all the settings
    Open "editor.ini" For Input As #1
    Do Until EOF(1)
        Line Input #1, lines(ind)
        ind = ind + 1
    Loop
    Close #1 'closes the editor.ini file
    ind = 0
    padbcolor = Val(Trim(lines(0))) 'sets editor back color
    padfcolor = Val(Trim(lines(1))) 'sets editor fore color
    padbold = Trim(lines(2))        'sets editor font style
    paditalic = Trim(lines(3))
    pfontname = Trim(lines(4))      'sets editor font name
    padfontsize = Val(Trim(lines(5))) 'sets editor font size
    javapath = Trim(lines(6))       'sets javapath
    commandpath = Trim(lines(7))    'sets command.com path
    browserpath = Trim(lines(8))    'sets browser path
    apipath = Trim(lines(9))        'sets java api help path
    defaultpath = Trim(lines(10))   'sets java default path
    sfile = defaultpath & "\" & "Noname.java" 'sets filename to Noname.java
      
    Exit Sub  'if all the things go well exit
filenotfound:
    MsgBox "Editor.ini File Not Found"
    Close #1
    End
End Sub

Private Sub Form_Load()
    ' Subclass the rtb so we can scroll the line numbers
    lPrevWndProc = SetWindowLong(writepad.hWnd, GWL_WNDPROC, AddressOf WindowProc)
    writepad.BackColor = padbcolor
    writepad.SelColor = padfcolor
    writepad.SelBold = padbold
    writepad.SelItalic = paditalic
    writepad.Font = pfontname
    writepad.SelFontSize = padfontsize
    picLines.Font = pfontname
    picLines.FontSize = padfontsize
    editor.Caption = "Java Editor 1.0  -  " & sfile
    
    With Toolbar1.Buttons
        .Item(7).Enabled = False
        .Item(8).Enabled = False
        .Item(9).Enabled = False
        .Item(11).Enabled = False
        .Item(12).Enabled = False
    End With
    
    picLines.PaintPicture picLines.Picture, 1, 1
    'invoking editor form
    editor.Show
    'invoking splash screen
    Splash.Show vbModal
End Sub

Public Function LineCount() As Long
    LineCount = SendMessageByRef(writepad.hWnd, EM_GETLINECOUNT, 0&, 0&)
End Function

Public Function LineForCharacterIndex(lIndex As Long) As Long
    LineForCharacterIndex = SendMessageByLong(writepad.hWnd, EM_LINEFROMCHAR, lIndex, 0)
End Function

Public Function FirstVisibleLine() As Long
    FirstVisibleLine = SendMessageByLong(writepad.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)
End Function

' This actually draws the line numbers created by the guys at vbaccelerator. Visit them at
' http://www.vbaccelerator.com. If you use this make sure to give them credit
Public Sub DrawLines(picTo As PictureBox)
Dim lTotChar As Long
Dim lLine As Long
Dim lCol As Long
Dim lCount As Long
Dim lCurrent As Long
Dim hBr As Long
Dim lEnd As Long
Dim lhDC As Long
Dim bComplete As Boolean
Dim tR As RECT, tTR As RECT
Dim oCol As OLE_COLOR
Dim lStart As Long
Dim lEndLine As Long
Dim tPO As POINTAPI
Dim lLineHeight As Long
Dim hPen As Long
Dim hPenOld As Long

   'Debug.Print "DrawLines"
   lhDC = picTo.hdc
   DrawText lhDC, "Hy", 2, tTR, DT_CALCRECT
   lLineHeight = tTR.Bottom - tTR.Top
   Dim Val As Double
   
   'Total number Lines
   lCount = LineCount
   'Current Line Number
   lCurrent = SendMessageLong(writepad.hWnd, EM_LINEFROMCHAR, writepad.SelStart, 0&)
   
   'Current Column Number
   Val = (SendMessageLong(writepad.hWnd, EM_LINEFROMCHAR, -1, 0&)) + 1
   lCol = ((SendMessageLong(writepad.hWnd, EM_GETSEL, 0, 0&) \ &H10000) - (SendMessageLong(writepad.hWnd, EM_LINEINDEX, Val - 1, 0&)))
   
   'Update Status Bar
   editor.StatusBar1.Panels(2).Text = "Ln " & lCurrent + 1 & ", Col " & lCol + 1
   editor.StatusBar1.Panels(3).Text = lCount
   'Total Characters
   
   lTotChar = writepad.SelStart
   
   
   lStart = writepad.SelStart
   lEnd = writepad.SelStart + writepad.SelLength - 1
   If (lEnd > lStart) Then
      lEndLine = LineForCharacterIndex(lEnd)
   Else
      lEndLine = lCurrent
   End If
   lLine = FirstVisibleLine
   GetClientRect picTo.hWnd, tR
   lEnd = tR.Bottom - tR.Top
      
   hBr = CreateSolidBrush(TranslateColor(picTo.BackColor))
   FillRect lhDC, tR, hBr
   DeleteObject hBr
   tR.Left = 2
   tR.Right = tR.Right - 2
   tR.Top = 0
   tR.Bottom = tR.Top + lLineHeight
   
   SetTextColor lhDC, TranslateColor(vbButtonShadow)
   
   Do
      ' Ensure correct colour:
      If (lLine = lCurrent) Then
         SetTextColor lhDC, TranslateColor(vbWindowText)
      ElseIf (lLine = lEndLine + 1) Then
         SetTextColor lhDC, TranslateColor(vbButtonShadow)
      End If
      ' Draw the line number:
      'DrawText lhDC, CStr(lCol + 1), -1, tR, DT_RIGHT
      DrawText lhDC, CStr(lLine + 1), -1, tR, DT_RIGHT
      
      ' Increment the line:
      lLine = lLine + 1
      ' Increment the position:
      OffsetRect tR, 0, lLineHeight
      If (tR.Bottom > lEnd) Or (lLine + 1 > lCount) Then
         bComplete = True
      End If
   Loop While Not bComplete
   
   ' Draw a line...
   MoveToEx lhDC, tR.Right + 1, 0, tPO
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbButtonShadow))
   hPenOld = SelectObject(lhDC, hPen)
   LineTo lhDC, tR.Right + 1, lEnd
   SelectObject lhDC, hPenOld
   DeleteObject hPen
   If picTo.AutoRedraw Then
      picTo.Refresh
   End If
   
End Sub

'This Procedure is invoked when user tries to close the application
'and is used to check whether the file is saved or not if not save it

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    Call closeedit 'call closeedit function
    cancel = 1 ' Dont terminate the execution
End Sub
Private Sub closeedit()
    Dim temp As Double
    'if file is not saved and file modified then request to save the file
    If isfilesaved = False And modified = True Then
        temp = MsgBox("The text in " & defaultpath & "\" & sfile & " has been changed." & vbCrLf & "Do u want to save changes?", vbYesNoCancel, "Java Editor 1.0")
        'if user clicked yes call save procedure
        If temp = vbYes Then Call mnu_save_Click
        If temp = vbCancel Then Exit Sub
    End If
    ' This kills the subclass so we don't screw up someone's machine
    Call SetWindowLong(writepad.hWnd, GWL_WNDPROC, lPrevWndProc)
    End
End Sub
'This procedure is called when the editor application is resized
Private Sub Form_Resize()
    
    'If Window is minimized then exit
    If editor.WindowState = 1 Then Exit Sub
    
    '-----------------------------------------------------
    'Picture1 Settings Begin
    
    Picture1.Left = 0
    Picture1.Width = editor.ScaleWidth
    
    'if toolbar, statusbar and line number is set
    If mnu_tool_bar.Checked Then
        Picture1.Top = Toolbar1.Height
    Else
        Picture1.Top = 0
    End If
    
    'if toolbar, statusbar is set
    If mnu_tool_bar.Checked And mnu_status_bar.Checked Then Picture1.Height = editor.ScaleHeight - Toolbar1.Height - StatusBar1.Height
    
    'if toolbar not set and statusbar is set
    If mnu_tool_bar.Checked = False And mnu_status_bar.Checked Then Picture1.Height = editor.ScaleHeight - StatusBar1.Height
    
    'if toolbar is set and statusbar not set
    If mnu_tool_bar.Checked And mnu_status_bar.Checked = False Then Picture1.Height = editor.ScaleHeight - Toolbar1.Height
    
    'if toolbar and statusbar are not set
    If mnu_tool_bar.Checked = False And mnu_status_bar.Checked = False Then Picture1.Height = editor.ScaleHeight
       
    'Picture1 Settings End
    '-----------------------------------------------------
    
    
    '-----------------------------------------------------
    'picLines Settings Begin
    
    picLines.Height = Picture1.Height - 85
    picLines.Width = 490
    
    'picLines Settings End
    '-----------------------------------------------------
    
    
    '-----------------------------------------------------
    'writepad settings Begin
    
    'if line number box is shown
    If mnu_line_number.Checked Then
        writepad.Left = picLines.Width
        writepad.Width = Picture1.Width - picLines.Width - 85
    Else
        writepad.Left = 0
        writepad.Width = Picture1.Width - 85
    End If
    writepad.Height = picLines.Height
    writepad.SetFocus
    
    'writepad Settings End
    '-----------------------------------------------------
    
    DrawLines picLines
End Sub
'This procedure shows About form
Private Sub mnu_about_javaedit_Click()
    About.Show vbModal
End Sub
'This procedure opens java api documentation
Private Sub mnu_api_Click()
    Dim temp As Double
    temp = Shell(browserpath & " " & apipath, vbMaximizedFocus)
End Sub
'This procedure runs Applet
Private Sub mnu_appletviewer_Click()
    Dim applet As Byte
    'Check whether file saved or not
    applet = checkfilesaved("Applet could be veiwed.", "Appletview")
    If applet = 0 Then Exit Sub
    Dim file As String, path As String, drv As String
    file = getFile(sfile)
    path = getPath(sfile)
    drv = getDrive(sfile)
    'Create Editor.bat by invoking CreateEditorbat procedure
    CreateEditorbat "@echo off" & vbCrLf & "cd " & Chr(34) & path & Chr(34) & vbCrLf & drv & vbCrLf & Chr(34) & javapath & "\appletviewer" & Chr(34) & " " & file & vbCrLf & "pause"
    Dim temp As Double
    'Execute editor.bat which in turn executes Applet
    temp = Shell("editor.bat ", vbMaximizedFocus)
    Call changeapppath
End Sub
'This procedure runs Applet with options
Private Sub mnu_appletviewer_with_Click()
    Dim applet As Byte
    applet = checkfilesaved("Applet could be veiwed.", "Appletview")
    If applet = 0 Then Exit Sub
    'Invoke appletviewer_with form
    appletviewer_with.Show , Me
End Sub

Private Sub mnu_application_Click()
    StatusBar1.Panels(1).Text = "Contain Commands for invoking Applications."
End Sub
' This procedure opens internet explorer
Private Sub mnu_browser_view_Click()
    Dim temp As Double
    temp = Shell(browserpath, vbMaximizedFocus)
End Sub
' This procedure clears the editor
Private Sub mnu_clear_Click()
    modified = True
    writepad.Text = ""
    DrawLines picLines
End Sub
'Function is used to warn user to save the file
Function checkfilesaved(mess As String, tit As String) As Byte
    If modified = True And isfilesaved = False Then
        Dim temp As Double
        temp = MsgBox("The file is to be saved before " & mess & vbCrLf & "Do u want to save the file", vbYesNo + vbQuestion, tit)
        If temp = vbYes Then ' if vbyes then save the file and compile or run
            Call mnu_save_Click
            checkfilesaved = 1
        Else
            checkfilesaved = 0 'if vbno then don't save the file
            Exit Function
        End If
    End If
    checkfilesaved = 1     ' allow compilation or running
End Function
'This procedure is used to compile a Java program
Private Sub mnu_compile_Click()
    ' check whether file is saved or not
    Dim compile As Byte
    compile = checkfilesaved("it could be compiled.", "Compilation")
    If compile = 0 Then Exit Sub
    'Create Editor.bat file
    CreateEditorbat Chr(34) & javapath & "\javac" & Chr(34) & " " & Chr(34) & sfile & Chr(34)
    compiled = True
    compile_splash.Label2 = compile_splash.Label2 & " " & getFile(editor.sfile) & "..."
    'Compile the file
    compile_splash.Show vbModal
End Sub
'This procedure is used to compile a Java program with options
Private Sub mnu_compile_with_Click()
    Dim compile As Byte
    ' check whether file is saved or not
    compile = checkfilesaved("Compilation could continue.", "compilation")
    If compile = 0 Then Exit Sub
    'invoke compile_with form
    compile_with.Show , Me
End Sub
    
Private Sub mnu_copy_Click()
    SendKeys ("^{c}")
End Sub

Private Sub mnu_cut_Click()
    SendKeys ("^{x}")
End Sub
'This procedures dynamically disables and enables Edit commands
Private Sub mnu_edit_Click()
    StatusBar1.Panels(1).Text = "Contain Edit Commands."
    If writepad.Text = "" Then
        mnu_clear.Enabled = False
        mnu_select_all.Enabled = False
        mnu_word_wrap.Enabled = False
        mnu_cut.Enabled = False
        mnu_copy.Enabled = False
        mnu_delete.Enabled = False
    End If
    If writepad.Text <> "" Then
        mnu_clear.Enabled = True
        mnu_select_all.Enabled = True
        mnu_word_wrap.Enabled = True
        If writepad.SelText = "" Then
            mnu_cut.Enabled = False
            mnu_copy.Enabled = False
            mnu_delete.Enabled = False
        ElseIf writepad.SelText <> "" Then
            mnu_cut.Enabled = True
            mnu_copy.Enabled = True
            mnu_delete.Enabled = True
        End If
    End If
End Sub
'This procedure invookes Env_settings form
Private Sub mnu_environment_Click()
    Env_settings.Show vbModal
End Sub

Private Sub mnu_exit_Click()
    Call closeedit
End Sub
'This procedure invokes Find form
Private Sub mnu_find_Click()
    position = 0
    Find.Show vbModal
End Sub

Private Sub mnu_go_to_Click()
    frmInputBox.Show vbModal
End Sub

Private Sub mnu_help_Click()
    StatusBar1.Panels(1).Text = "Contain Commands for Displaying Help."
End Sub
'This procedure invokes Java Jar form
Private Sub mnu_jar_Click()
    Jar.Show , Me
End Sub
'This procedure invokes Javadoc form
Private Sub mnu_javadoc_Click()
    Dim jdoc As Byte
    jdoc = checkfilesaved("Java Documentation could continue.", "JavaDoc")
    If jdoc = 0 Then Exit Sub
    calledfrom = fromJavadoc
    javadoc.Show , Me
End Sub
'This procedure invokes javap form
Private Sub mnu_javap_Click()
    javap.Show , Me
End Sub

Private Sub mnu_line_number_Click()
    mnu_line_number.Checked = Not mnu_line_number.Checked
    picLines.Visible = mnu_line_number.Checked
    Form_Resize
End Sub

Private Sub mnu_maximize_Click()
    editor.WindowState = 2
End Sub

Private Sub mnu_minimize_Click()
    editor.WindowState = 1
End Sub
'This procedure invokes MS-Dos Command Prompt
Private Sub mnu_msdosprompt_Click()
    Dim temp As Double
    temp = Shell(commandpath, vbMaximizedFocus)
End Sub
'This procedure creates new file
Private Sub mnu_new_Click()
    If modified Then  'If File Opened and Modified then request to save the file
        Dim temp As Double
        temp = MsgBox("The text in " & defaultpath & "\" & sfile & " has been changed." & vbCrLf & "Do u want to save changes?", vbYesNoCancel, "Java Editor 1.0")
        If temp = vbYes Then Call mnu_save_Click
        If temp = vbCancel Then Exit Sub
    End If
    modified = False
    isfilesaved = False
    sfile = defaultpath & "\" & "Noname.java"
    editor.Caption = "Java Editor 1.0  -  " & sfile
    writepad.Text = ""
    DrawLines picLines
End Sub

Private Sub mnu_normal_Click()
    editor.WindowState = 0
End Sub
'This procedure opens a file
Private Sub mnu_open_Click()
    If modified Then  'If File Opened and Modified then request to save the file
        Dim temp As Double
        temp = MsgBox("The text in " & defaultpath & "\" & sfile & " has been changed." & vbCrLf & "Do u want to save changes?", vbYesNoCancel, "Java Editor 1.0")
        If temp = vbYes Then Call mnu_save_Click
        If temp = vbCancel Then Exit Sub
    End If
    On Error GoTo nofile
    'Setting Open Dialog box Properties
    With cdg1
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .InitDir = defaultpath
        .Filter = "Java file (*.java)|*java|Class file (*.class)|*.class|Jar file (*.jar)|*.jar|Html Document (*.html, *.htm)|*.html;*.htm|Rich Text File (*.rtf)|*.rtf|All files (*.*)|*.*"
        .FileName = "*.java"
        .ShowOpen
        If Len(.FileName) = 0 Or .FileName = "*.java" Then Exit Sub
    End With
    'setting sfile to Opened file
    sfile = cdg1.FileName
    Dim FilePath As String
    FilePath = getPath(sfile)
    If FilePath <> defaultpath Then defaultpath = FilePath
    modified = False
    isfilesaved = False
    'Opening the file and loading in writepad RTF
    writepad.LoadFile cdg1.FileName, 1
    DrawLines picLines
    'changing the editor.caption
    editor.Caption = "Java Editor 1.0  -  " & sfile
nofile:
End Sub
'This procedures is invoked to show (if any) errors in java program
Private Sub mnu_output_Click()
    Output.Show vbModal
End Sub

Private Sub mnu_paste_Click()
    SendKeys ("^{v}")
End Sub
'This procedure prints the java file
Private Sub mnu_print_Click()
On Error GoTo cancel
Dim temp As Integer
'if there is no text then give a message
If writepad.Text = "" Then
    temp = MsgBox("Nothing To Print.", vbExclamation, "Print Job Cancel")
Else 'setting Printer options
    cdg1.PrinterDefault = True
    cdg1.Flags = cdlPDCollate Or cdlPDReturnIC & cdlPDDisablePrintToFile
    cdg1.ShowPrinter
End If
cancel:
End Sub
'This procedure pops print setup dialog box
Private Sub mnu_print_setup_Click()
    cdg1.Flags = cdlPDPrintSetup
    cdg1.ShowPrinter
End Sub

Private Sub mnu_project_Click()
    StatusBar1.Panels(1).Text = "Contain Commands for Compiling Java files."
    If writepad.Text = "" Then
        mnu_compile.Enabled = False
        mnu_appletviewer.Enabled = False
    Else
        mnu_compile.Enabled = True
        mnu_appletviewer.Enabled = True
    End If
End Sub
'This procedure is invoked for find and replace
Private Sub mnu_replace_Click()
    'Replace.Show vbModeless, Me
    Replace.Show vbModal
End Sub
'This procedure is invoked to run a java application
Private Sub mnu_run_Click()
    Dim runclass As String
    Dim FilePath As String
    Dim FileDrive As String
    Dim FileName As String

    FileDrive = getDrive(sfile)
    FilePath = getPath(sfile)
    FileName = removeExt(getFile(sfile))
    runclass = InputBox("Enter the (main) class name without extension in" & vbLf & vbCr & FilePath & vbLf & vbCr & "(with options and arguments if any...)", "Enter the Class Name", FileName)
    If runclass = "" Then Exit Sub 'if Cancel Button Clicked then exit run sub

    'Creates Editor.Bat
    CreateEditorbat "@echo off" & vbCrLf & Chr(34) & javapath & "\java" & Chr(34) & " " & "-classpath " & Chr(34) & FilePath & Chr(34) & " " & runclass & vbCrLf & "pause"
    Dim cmdLine As String
    cmdLine = CurDir() & "\editor.bat"
    'Change to Class's Directory
    Call changeDirectory(FileDrive, FilePath)
    'Execute Editor.bat
    RunShell cmdLine, vbMaximizedFocus

    'Change to Class's Directory
    Call changeapppath
End Sub
'This procedure runs java program with options
Private Sub mnu_run_with_Click()
    'invoke run_with form
    run_with.Show , Me
End Sub
'This procedure saves the file
Private Sub mnu_save_Click()
    Dim pos As Byte
    pos = InStrRev(sfile, "Noname.java")
    If pos <> 0 Then
        pos = savethefile
    Else
        writepad.SaveFile sfile, 1
        isfilesaved = True
        modified = False
    End If
    If pos = 1 Then
        writepad.SaveFile sfile, 1
        isfilesaved = True
        modified = False
    End If
End Sub
'This Function is invoked to save the file
Private Function savethefile() As Byte
    'sets Save dialog box settings
    On Error GoTo err 'cdlOFNExtensionDifferent Or
    cdg1.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    cdg1.Filter = "Java file (*.java)|*java|Class file (*.class)|*.class|Jar file (*.jar)|*.jar|Html Document (*.html, *.htm)|*.html;*.htm|Rich Text File (*.rtf)|*.rtf|All files (*.*)|*.*"
    cdg1.FileName = "*.java"
    cdg1.InitDir = defaultpath
    cdg1.ShowSave
    If Len(cdg1.FileName) < 7 Then GoTo err
    savethefile = 1
    Dim pos As Integer
    pos = InStrRev(cdg1.FileName, ".java")
    If pos = 0 And cdg1.FilterIndex = 1 Then cdg1.FileName = cdg1.FileName & ".java"
    sfile = cdg1.FileName
    Dim FilePath As String
    FilePath = getPath(sfile)
    If FilePath <> defaultpath Then defaultpath = FilePath
    editor.Caption = "Java Editor 1.0  -  " & sfile
    Exit Function
err:
    savethefile = 0
End Function
'Save as
Private Sub mnu_saveas_Click()
    Dim saveit As Byte
    saveit = savethefile
    If saveit = 0 Then Exit Sub
    writepad.SaveFile sfile, 1
    isfilesaved = True
    modified = False
End Sub

Private Sub mnu_search_Click()
    StatusBar1.Panels(1).Text = "Contain Commands for Searching Text."
    If writepad.Text = "" Then
        mnu_find.Enabled = False
        mnu_find_next.Enabled = False
        mnu_replace.Enabled = False
        mnu_go_to.Enabled = False
    ElseIf writepad.Text <> "" Then
        mnu_find.Enabled = True
        mnu_find_next.Enabled = True
        mnu_replace.Enabled = True
        mnu_go_to.Enabled = True
    End If
End Sub

Private Sub mnu_select_all_Click()
    SendKeys ("^{a}")
End Sub
'This invokes Setup form
Private Sub mnu_setup_Click()
    setup.Show , Me
End Sub

Private Sub mnu_status_bar_Click()
    mnu_status_bar.Checked = Not mnu_status_bar.Checked
    StatusBar1.Visible = mnu_status_bar.Checked
    Form_Resize
End Sub

Private Sub mnu_tool_bar_Click()
    mnu_tool_bar.Checked = Not mnu_tool_bar.Checked
    Toolbar1.Visible = mnu_tool_bar.Checked
    Form_Resize
End Sub

Private Sub mnu_undo_Click()
    SendKeys ("^{z}")
End Sub

Private Sub mnu_view_Click()
    StatusBar1.Panels(1).Text = "Contain Commands for manipulating the View."
    If editor.WindowState = 2 Then
        mnu_maximize.Enabled = False
        mnu_normal.Enabled = True
    ElseIf editor.WindowState = 0 Then
        mnu_maximize.Enabled = True
        mnu_normal.Enabled = False
    End If
End Sub
'This procedure is invoked to turn word wrap feature on or off
Private Sub mnu_word_wrap_Click()
    mnu_word_wrap.Checked = Not mnu_word_wrap.Checked
    If mnu_word_wrap.Checked Then
        writepad.RightMargin = 0
    Else
        writepad.RightMargin = 10000000
    End If
End Sub
'This procedure is used to handle toolbar clicks
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Tag
            Case 1
                Call mnu_new_Click
            Case 2
                Call mnu_open_Click
            Case 3
                Call mnu_save_Click
            Case 4
                Call mnu_print_Click
            Case 6
                SendKeys ("^{x}")
            Case 7
                SendKeys ("^{c}")
            Case 8
                SendKeys ("^(v)")
            Case 9
                SendKeys ("^{z}")
            Case 10
                Call mnu_find_Click
            Case 11
                Call mnu_compile_Click
            Case 12
                Call mnu_run_Click
            Case 13
                Call mnu_appletviewer_Click
            Case 14
                Call mnu_msdosprompt_Click
            Case 15
                Call mnu_browser_view_Click
        End Select
End Sub

Private Sub writepad_Change()
    If writepad.Text = "" Then
        mnu_compile.Enabled = False
        mnu_appletviewer.Enabled = False
        With Toolbar1.Buttons
            .Item(7).Enabled = False  'cut
            .Item(8).Enabled = False  'copy
            .Item(9).Enabled = False  'paste
            .Item(11).Enabled = False 'undo
            .Item(12).Enabled = False 'find
            .Item(14).Enabled = False 'Compile
            .Item(16).Enabled = False 'Run
        End With
    Else
        mnu_compile.Enabled = True
        mnu_appletviewer.Enabled = True
        With Toolbar1.Buttons
            .Item(7).Enabled = True
            .Item(8).Enabled = True
            .Item(9).Enabled = True
            .Item(11).Enabled = True
            .Item(12).Enabled = True
            .Item(14).Enabled = True
            .Item(16).Enabled = True
        End With
    End If
    DrawLines picLines
End Sub

Private Sub writepad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then isfilesaved = False: modified = True
    If KeyCode = vbKeyReturn Then lineno = lineno + 1
    movepos = writepad.SelStart
End Sub

Private Sub writepad_KeyPress(KeyAscii As Integer)
    'MsgBox writepad.SelStart
    isfilesaved = False
    modified = True
    col = col + 1
    'If KeyAscii = 13 Then lineno = lineno + 1
End Sub

Private Sub writepad_KeyUp(KeyCode As Integer, Shift As Integer)
    DrawLines picLines
End Sub

Private Sub writepad_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawLines picLines
End Sub

Private Sub writepad_SelChange()
    DrawLines picLines
End Sub
