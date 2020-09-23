VERSION 5.00
Begin VB.Form javadoc 
   Caption         =   "Javadoc - The Java API Documentation Generator"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   Icon            =   "javadoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_stylesheetfile 
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
      Left            =   9360
      TabIndex        =   78
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmd_helpfile 
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
      Left            =   9360
      TabIndex        =   77
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt_helpfile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   76
      Text            =   " "
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox txt_stylesheetfile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      TabIndex        =   75
      Text            =   " "
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox txt_docencoding 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7680
      TabIndex        =   74
      Text            =   " "
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox txt_group 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   73
      Text            =   " "
      Top             =   5880
      Width           =   3255
   End
   Begin VB.CheckBox chk_group 
      Caption         =   "-group <heading> <p1>:<p2>: ..."
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
      Left            =   3240
      TabIndex        =   72
      Top             =   5880
      Width           =   3135
   End
   Begin VB.CheckBox chk_docencoding 
      Caption         =   "-docencoding <name>"
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
      Left            =   5040
      TabIndex        =   71
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CheckBox chk_stylesheetfile 
      Caption         =   "-stylesheetfile <file>"
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
      Left            =   5040
      TabIndex        =   70
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CheckBox chk_helpfile 
      Caption         =   "-helpfile <file>"
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
      Left            =   5040
      TabIndex        =   69
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CheckBox chk_nonavbar 
      Caption         =   "-nonavbar"
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
      Left            =   8400
      TabIndex        =   68
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CheckBox chk_nohelp 
      Caption         =   "-nohelp"
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
      Left            =   7320
      TabIndex        =   67
      Top             =   4440
      Width           =   975
   End
   Begin VB.CheckBox chk_noindex 
      Caption         =   "-noindex"
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
      Left            =   6120
      TabIndex        =   66
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CheckBox chk_notree 
      Caption         =   "-notree"
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
      Left            =   5040
      TabIndex        =   65
      Top             =   4440
      Width           =   975
   End
   Begin VB.CheckBox chk_nodeprecatedlist 
      Caption         =   "-nodeprecatedlist"
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
      Left            =   7320
      TabIndex        =   64
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CheckBox chk_nodeprecated 
      Caption         =   "-nodeprecated"
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
      Left            =   5040
      TabIndex        =   63
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CheckBox chk_linkoffline 
      Caption         =   "-linkoffline <ur1l> <url2>"
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
      Left            =   5040
      TabIndex        =   62
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox txt_linkoffline 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   61
      Text            =   " "
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CheckBox chk_link 
      Caption         =   "-link  <url>"
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
      Left            =   5040
      TabIndex        =   60
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chk_bottom 
      Caption         =   "-bottom <HTML -code>"
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
      Left            =   5040
      TabIndex        =   59
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CheckBox chk_footer 
      Caption         =   "-footer <HTML -code>"
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
      Left            =   5040
      TabIndex        =   58
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CheckBox chk_header 
      Caption         =   "-header <HTML -code>"
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
      Left            =   5040
      TabIndex        =   57
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txt_link 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   56
      Text            =   " "
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txt_bottom 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   55
      Text            =   " "
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox txt_footer 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   54
      Text            =   " "
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txt_header 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   53
      Text            =   " "
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txt_doctitle 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   52
      Text            =   " "
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txt_windowtitle 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   51
      Text            =   " "
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmd_d 
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
      Left            =   8280
      TabIndex        =   50
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txt_d 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   49
      Text            =   " "
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox chk_locale 
      Caption         =   "-locale <name>"
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
      TabIndex        =   48
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CheckBox chk_encoding 
      Caption         =   "-encoding <name>"
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
      TabIndex        =   47
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CheckBox chk_J 
      Caption         =   "-J <javaflag(s)>"
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
      TabIndex        =   46
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox txt_locale 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   45
      Text            =   " "
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox txt_encoding 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   44
      Text            =   " "
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txt_J 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   43
      Text            =   " "
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmd_cancel 
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
      Left            =   3960
      TabIndex        =   42
      Top             =   1320
      Width           =   975
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
      Left            =   3960
      TabIndex        =   41
      Top             =   1920
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
      Left            =   2880
      TabIndex        =   40
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
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
      Left            =   2880
      TabIndex        =   39
      Top             =   1320
      Width           =   975
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
      Index           =   5
      Left            =   4560
      TabIndex        =   38
      Top             =   4440
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
      Index           =   4
      Left            =   4560
      TabIndex        =   37
      Top             =   4080
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
      Left            =   4560
      TabIndex        =   36
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   35
      Text            =   " "
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   34
      Text            =   " "
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   33
      Text            =   " "
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CheckBox chk_extdirs 
      Caption         =   "-extdirs <path(s)>"
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
      Top             =   4440
      Width           =   1815
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
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CheckBox chk_classpath 
      Caption         =   "-classpath <paths(s)>"
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
      Top             =   3720
      Width           =   2175
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
      Left            =   4560
      TabIndex        =   29
      Top             =   3360
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
      Left            =   4560
      TabIndex        =   28
      Top             =   3000
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
      Left            =   4560
      TabIndex        =   27
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   26
      Text            =   " "
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   25
      Text            =   " "
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txt_paths 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   24
      Text            =   " "
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmd_overview 
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
      Left            =   4560
      TabIndex        =   23
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txt_overview 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   22
      Text            =   " "
      Top             =   840
      Width           =   2535
   End
   Begin VB.CheckBox chk_doctitle 
      Caption         =   "-doctitle <HTML -code>"
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
      Left            =   5040
      TabIndex        =   21
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CheckBox chk_windowtitle 
      Caption         =   "-windowtitle <title>"
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
      Left            =   5040
      TabIndex        =   20
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox chk_author 
      Caption         =   "-author"
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
      Left            =   7080
      TabIndex        =   19
      Top             =   1200
      Width           =   975
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
      Left            =   5880
      TabIndex        =   18
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox chk_use 
      Caption         =   "-use"
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
      Left            =   5040
      TabIndex        =   17
      Top             =   1200
      Width           =   735
   End
   Begin VB.CheckBox chk_d 
      Caption         =   "-d <path>"
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
      Left            =   5040
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox chk_1 
      Caption         =   "-1.1"
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
      Left            =   1560
      TabIndex        =   15
      Top             =   1920
      Width           =   735
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
      Left            =   1560
      TabIndex        =   14
      Top             =   1560
      Width           =   975
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
      Left            =   1560
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
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
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CheckBox chk_docletpath 
      Caption         =   "-dockletpath <path>"
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
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox chk_doclet 
      Caption         =   "-doclet <class>"
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
      TabIndex        =   10
      Top             =   2640
      Width           =   1695
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
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
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
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   855
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
      Left            =   120
      TabIndex        =   7
      Top             =   1560
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
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox chk_overview 
      Caption         =   "-overview <file>"
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
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmd_javadoc 
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
      Left            =   9360
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txt_javadoc 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   " "
      Top             =   240
      Width           =   9135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Standard Doclet Options :"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   600
      Width           =   2310
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Options :"
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
      TabIndex        =   3
      Top             =   600
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Javadoc [options] [ packagename(s) ]  [ sourcefil(s) ] [ @file(s) ]"
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
      TabIndex        =   0
      Top             =   0
      Width           =   5460
   End
End
Attribute VB_Name = "javadoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sel As Byte  '1 for javadoc
                    '2 for overview
                    '3 for d
                    '4 for helpfile
                    '5 for stylesheetfile
Public txtjavadoc As String
'Procedure to update the txtjavadoc textbox
Public Sub changetext()
    txt_javadoc.Text = Chr(34) & editor.javapath & "\javadoc" & Chr(34) & " "
    If chk_overview.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-overview "
        If txt_overview.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & txt_overview.Text & " "
    End If
    If chk_public.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-public "
    If chk_protected.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-protected "
    If chk_package.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-package "
    If chk_private.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-private "
    If chk_help.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-help "
    If chk_1.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-1.1 "
    If chk_verbose.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-verbose "
    If chk_doclet.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-doclet "
        If txt_paths(0).Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & " " & txt_paths(0) & " "
    End If
    If chk_docletpath.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-docletpath "
        If txt_paths(1).Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & " " & txt_paths(1) & " "
    End If
    If chk_sourcepath.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-sourcepath "
        If txt_paths(2).Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & " " & txt_paths(2) & " "
    End If
    If chk_classpath.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-classpath "
        If txt_paths(3).Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & " " & txt_paths(3) & " "
    End If
    If chk_bootclasspath.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-bootclasspath "
        If txt_paths(4).Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & " " & txt_paths(4) & " "
    End If
    If chk_extdirs.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-extdirs "
        If txt_paths(5).Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & " " & txt_paths(5) & " "
    End If
    If chk_locale.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-locale "
        If txt_locale.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_locale & Chr(34) & " "
    End If
    If chk_encoding.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-encoding "
        If txt_encoding.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_encoding & Chr(34) & " "
    End If
    If chk_J.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-J "
        If txt_J.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_J & Chr(34) & " "
    End If
    If chk_d.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-d "
        If txt_d.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & txt_d & " "
    End If
    If chk_use.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-use "
    If chk_version.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-version "
    If chk_author.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-author "
    If chk_windowtitle.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-windowtitle "
        txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_windowtitle & Chr(34) & " "
    End If
    If chk_doctitle.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-doctitle "
        txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_doctitle & Chr(34) & " "
    End If
    If chk_header.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-header "
        txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_header & Chr(34) & " "
    End If
    If chk_footer.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-footer "
        txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_footer & Chr(34) & " "
    End If
    If chk_bottom.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-bottom "
        txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_bottom & Chr(34) & " "
    End If
    If chk_link.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-link "
        If txt_link.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_link & Chr(34) & " "
    End If
    If chk_linkoffline.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-linkoffline "
        If txt_linkoffline.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_linkoffline & Chr(34) & " "
    End If
    If chk_nodeprecated.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-nodeprecated "
    If chk_nodeprecatedlist.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-nodeprecatedlist "
    If chk_notree.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-notree "
    If chk_noindex.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-noindex "
    If chk_nohelp.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-nohelp "
    If chk_nonavbar.Value = 1 Then txt_javadoc.Text = txt_javadoc.Text & "-nonavbar "
    If chk_helpfile.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-helpfile "
        If txt_helpfile.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & txt_helpfile & " "
    End If
    If chk_stylesheetfile.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-stylesheetfile "
        If txt_stylesheetfile.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & txt_stylesheetfile & " "
    End If
    If chk_docencoding.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-docencoding "
        If txt_docencoding.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_docencoding & Chr(34) & " "
    End If
    If chk_group.Value = 1 Then
        txt_javadoc.Text = txt_javadoc.Text & "-group "
        If txt_group.Text <> "" Then txt_javadoc.Text = txt_javadoc.Text & Chr(34) & txt_group & Chr(34) & " "
    End If
    txt_javadoc.Text = txt_javadoc.Text & getFile(editor.sfile)
End Sub
Private Sub chk_1_Click()
    Call changetext
End Sub

Private Sub chk_author_Click()
    Call changetext
End Sub

Private Sub chk_bootclasspath_Click()
    If chk_bootclasspath.Value = 1 Then
        cmd_paths(4).Enabled = True
        txt_paths(4).Enabled = True
        txt_paths(4).SetFocus
    Else
        cmd_paths(4).Enabled = False
        txt_paths(4).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_bottom_Click()
    If chk_bottom.Value = 1 Then
        txt_bottom.Enabled = True
        txt_bottom.SetFocus
    Else
        txt_bottom.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_classpath_Click()
    If chk_classpath.Value = 1 Then
        cmd_paths(3).Enabled = True
        txt_paths(3).Enabled = True
        txt_paths(3).SetFocus
    Else
        cmd_paths(3).Enabled = False
        txt_paths(3).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_d_Click()
    If chk_d.Value = 1 Then
        cmd_d.Enabled = True
        txt_d.Enabled = True
        txt_d.SetFocus
    Else
        cmd_d.Enabled = False
        txt_d.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_docencoding_Click()
    If chk_docencoding.Value = 1 Then
        txt_docencoding.Enabled = True
        txt_docencoding.SetFocus
    Else
        txt_docencoding.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_doclet_Click()
    If chk_doclet.Value = 1 Then
        cmd_paths(0).Enabled = True
        txt_paths(0).Enabled = True
        txt_paths(0).SetFocus
    Else
        cmd_paths(0).Enabled = False
        txt_paths(0).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_docletpath_Click()
    If chk_docletpath.Value = 1 Then
        cmd_paths(1).Enabled = True
        txt_paths(1).Enabled = True
        txt_paths(1).SetFocus
    Else
        cmd_paths(1).Enabled = False
        txt_paths(1).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_doctitle_Click()
    If chk_doctitle.Value = 1 Then
        txt_doctitle.Enabled = True
        txt_doctitle.SetFocus
    Else
        txt_doctitle.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_encoding_Click()
    If chk_encoding.Value = 1 Then
        txt_encoding.Enabled = True
        txt_encoding.SetFocus
    Else
        txt_encoding.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_extdirs_Click()
    If chk_extdirs.Value = 1 Then
        cmd_paths(5).Enabled = True
        txt_paths(5).Enabled = True
        txt_paths(5).SetFocus
    Else
        cmd_paths(5).Enabled = False
        txt_paths(5).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_footer_Click()
    If chk_footer.Value = 1 Then
        txt_footer.Enabled = True
        txt_footer.SetFocus
    Else
        txt_footer.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_group_Click()
    If chk_group.Value = 1 Then
        txt_group.Enabled = True
        txt_group.SetFocus
    Else
        txt_group.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_header_Click()
    If chk_header.Value = 1 Then
        txt_header.Enabled = True
        txt_header.SetFocus
    Else
        txt_header.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_help_Click()
    Call changetext
End Sub

Private Sub chk_helpfile_Click()
    If chk_helpfile.Value = 1 Then
        cmd_helpfile.Enabled = True
        txt_helpfile.Enabled = True
        txt_helpfile.SetFocus
    Else
        cmd_helpfile.Enabled = False
        txt_helpfile.Enabled = False
    End If
    Call changetext
End Sub
    
Private Sub chk_J_Click()
    If chk_J.Value = 1 Then
        txt_J.Enabled = True
        txt_J.SetFocus
    Else
        txt_J.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_link_Click()
    If chk_link.Value = 1 Then
        txt_link.Enabled = True
        txt_link.SetFocus
    Else
        txt_link.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_linkoffline_Click()
    If chk_linkoffline.Value = 1 Then
        txt_linkoffline.Enabled = True
        txt_linkoffline.SetFocus
    Else
        txt_linkoffline.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_locale_Click()
    If chk_locale.Value = 1 Then
        txt_locale.Enabled = True
        txt_locale.SetFocus
    Else
        txt_locale.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_nodeprecated_Click()
    Call changetext
End Sub

Private Sub chk_nodeprecatedlist_Click()
    Call changetext
End Sub

Private Sub chk_nohelp_Click()
    Call changetext
End Sub

Private Sub chk_noindex_Click()
    Call changetext
End Sub

Private Sub chk_nonavbar_Click()
    Call changetext
End Sub

Private Sub chk_notree_Click()
    Call changetext
End Sub

Private Sub chk_overview_Click()
    If chk_overview.Value = 1 Then
        txt_overview.Enabled = True
        txt_overview.SetFocus
        cmd_overview.Enabled = True
    Else
        txt_overview.Enabled = False
        cmd_overview.Enabled = False
    End If
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

Private Sub chk_sourcepath_Click()
    If chk_sourcepath.Value = 1 Then
        cmd_paths(2).Enabled = True
        txt_paths(2).Enabled = True
        txt_paths(2).SetFocus
    Else
        cmd_paths(2).Enabled = False
        txt_paths(2).Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_stylesheetfile_Click()
    If chk_stylesheetfile.Value = 1 Then
        cmd_stylesheetfile.Enabled = True
        txt_stylesheetfile.Enabled = True
        txt_stylesheetfile.SetFocus
    Else
        cmd_stylesheetfile.Enabled = False
        txt_stylesheetfile.Enabled = False
    End If
    Call changetext
End Sub

Private Sub chk_use_Click()
    Call changetext
End Sub

Private Sub chk_verbose_Click()
    Call changetext
End Sub

Private Sub chk_version_Click()
    Call changetext
End Sub

Private Sub chk_windowtitle_Click()
    If chk_windowtitle.Value = 1 Then
        txt_windowtitle.Enabled = True
        txt_windowtitle.SetFocus
    Else
        txt_windowtitle.Enabled = False
    End If
    Call changetext
End Sub

Private Sub cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub cmd_d_Click()
    sel = 3
    editor.calledfrom = fromJavadoc 'set called from to 8
    'to indicated that PathFile form is called from javadoc form
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Visible = False
    PathFile.Label1.Caption = "Select the Directory:"
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub cmd_helpfile_Click()
    sel = 4
    editor.calledfrom = fromJavadoc 'set called from to 8
    'to indicated that PathFile form is called from javadoc form
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Visible = True
    PathFile.Label1.Caption = "Select a file:"
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub cmd_javadoc_Click()
    sel = 1
    editor.calledfrom = fromJavadoc 'set called from to 8
    'to indicated that PathFile form is called from javadoc form
    PathFile.Opt_Directory.Caption = "&Package name / Directory"
    PathFile.opt_File.Caption = "Source &File"
    PathFile.opt_all.Caption = "&@Filename"
    PathFile.opt_File.Visible = True
    PathFile.opt_File.Value = True
    PathFile.opt_all.Visible = True
    PathFile.Opt_Directory.Visible = True
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Visible = True
    PathFile.Label1.Caption = "Select a file:"
    PathFile.Show vbModal
    Call changetext
End Sub
'Procedure to create Editor.bat and execute it
Private Sub cmd_OK_Click()
    CreateEditorbat txt_javadoc.Text & vbCrLf & "pause"
    Dim temp As Double
    temp = Shell("editor.bat ", vbMaximizedFocus)
End Sub

Private Sub cmd_overview_Click()
    sel = 2
    editor.calledfrom = fromJavadoc 'set called from to 8
    'to indicated that PathFile form is called from javadoc form
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Visible = True
    PathFile.Label1.Caption = "Select a file:"
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub cmd_paths_Click(Index As Integer)
    editor.calledfrom = fromJavadoc 'set called from to 8
    'to indicated that PathFile form is called from javadoc form
    editor.ind = Index
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    If Index = 0 Then
        PathFile.File1.Visible = True
        PathFile.Label1.Caption = "Select the class file:"
    Else
        PathFile.File1.Visible = False
        PathFile.Label1.Caption = "Select the Directory:"
    End If
    PathFile.Show vbModal
    Call changetext
End Sub
'Procedure that resets the settings of Javadoc options
Private Sub cmd_reset_Click()
    Call Form_Load
End Sub

Private Sub cmd_stylesheetfile_Click()
    sel = 5
    editor.calledfrom = fromJavadoc 'set called from to 8
    'to indicated that PathFile form is called from javadoc form
    PathFile.Drive1.Drive = getDrive(editor.defaultpath)
    PathFile.Dir1.path = editor.defaultpath
    PathFile.File1.Visible = True
    PathFile.Label1.Caption = "Select a file:"
    PathFile.Show vbModal
    Call changetext
End Sub

Private Sub Form_Load()
    editor.ind = -1
    txtjavadoc = ""
    txt_javadoc.Text = Chr(34) & editor.javapath & "\javadoc" & Chr(34) & " " & getFile(editor.sfile)
        
                'DISABLING ALL CHECK BOXES

    chk_1.Value = 0: chk_overview.Value = 0: chk_public.Value = 0: chk_protected.Value = 0
    chk_private.Value = 0: chk_package.Value = 0: chk_help.Value = 0: chk_verbose.Value = 0
    chk_doclet.Value = 0: chk_docletpath.Value = 0: chk_sourcepath.Value = 0: chk_classpath.Value = 0
    chk_bootclasspath.Value = 0: chk_extdirs.Value = 0: chk_locale.Value = 0: chk_encoding.Value = 0
    chk_J.Value = 0: chk_d.Value = 0: chk_use.Value = 0: chk_version.Value = 0: chk_author.Value = 0
    chk_windowtitle.Value = 0: chk_doctitle.Value = 0: chk_footer.Value = 0: chk_header.Value = 0
    chk_bottom.Value = 0: chk_link.Value = 0: chk_linkoffline.Value = 0: chk_nodeprecated.Value = 0
    chk_nodeprecatedlist.Value = 0: chk_notree.Value = 0: chk_noindex.Value = 0: chk_nohelp.Value = 0
    chk_nonavbar.Value = 0: chk_helpfile.Value = 0: chk_stylesheetfile.Value = 0
    chk_docencoding.Value = 0: chk_group.Value = 0

                'DISABLING ALL TEXT BOXES and COMMAND BUTTONS
                
    txt_overview.Enabled = False: cmd_overview.Enabled = False: txt_overview.Text = ""
    Dim i As Byte
    For i = 0 To 5
        txt_paths(i).Enabled = False: cmd_paths(i).Enabled = False: txt_paths(i).Text = ""
    Next i
    txt_locale.Enabled = False: txt_locale.Text = ""
    txt_encoding.Enabled = False: txt_encoding.Text = ""
    txt_J.Enabled = False: txt_J.Text = ""
    txt_d.Enabled = False: cmd_d.Enabled = False: txt_d.Text = ""
    txt_windowtitle.Enabled = False: txt_windowtitle.Text = ""
    txt_doctitle.Enabled = False: txt_doctitle.Text = ""
    txt_header.Enabled = False: txt_header.Text = ""
    txt_footer.Enabled = False: txt_footer.Text = ""
    txt_bottom.Enabled = False: txt_bottom.Text = ""
    txt_link.Enabled = False: txt_link.Text = ""
    txt_linkoffline.Enabled = False: txt_linkoffline.Text = ""
    txt_helpfile.Enabled = False: txt_helpfile.Text = "": cmd_helpfile.Enabled = False
    txt_stylesheetfile.Enabled = False: txt_stylesheetfile.Text = "": cmd_stylesheetfile.Enabled = False
    txt_docencoding.Enabled = False: txt_docencoding.Text = ""
    txt_group.Enabled = False: txt_group.Text = ""
    
End Sub

Private Sub txt_bottom_Change()
    Call changetext
End Sub

Private Sub txt_d_Change()
    Call changetext
End Sub

Private Sub txt_docencoding_Change()
    Call changetext
End Sub

Private Sub txt_doctitle_Change()
    Call changetext
End Sub

Private Sub txt_encoding_Change()
    Call changetext
End Sub

Private Sub txt_footer_Change()
    Call changetext
End Sub

Private Sub txt_group_Change()
    Call changetext
End Sub

Private Sub txt_header_Change()
    Call changetext
End Sub

Private Sub txt_helpfile_Change()
    Call changetext
End Sub

Private Sub txt_J_Change()
    Call changetext
End Sub

Private Sub txt_link_Change()
    Call changetext
End Sub

Private Sub txt_linkoffline_Change()
    Call changetext
End Sub

Private Sub txt_locale_Change()
    Call changetext
End Sub

Private Sub txt_overview_Change()
    Call changetext
End Sub

Private Sub txt_paths_Change(Index As Integer)
    Call changetext
End Sub

Private Sub txt_stylesheetfile_Change()
    Call changetext
End Sub

Private Sub txt_windowtitle_Change()
    Call changetext
End Sub
