VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmAuditService 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Audit Service"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   0
      Picture         =   "frmAuditService.frx":0000
      ScaleHeight     =   9615
      ScaleWidth      =   495
      TabIndex        =   55
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   12360
      TabIndex        =   50
      Top             =   2760
      Width           =   2655
      Begin VB.CheckBox chIncNol 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Nol (0)"
         Height          =   195
         Left            =   0
         TabIndex        =   53
         Top             =   2040
         Width           =   1695
      End
      Begin VB.OptionButton opField 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jasa + Sublet"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   2055
      End
      Begin VB.OptionButton opField 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sublet"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton opField 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Part"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton opField 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jasa"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "... Field sisi Service ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   51
         Top             =   60
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   5  'Dash-Dot-Dot
         Height          =   1935
         Left            =   0
         Top             =   0
         Width           =   2295
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.OptionButton opTipe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Service  >>  Ledger"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   12360
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
   End
   Begin VB.OptionButton opTipe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ledger  >>  Service"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   12360
      TabIndex        =   12
      Top             =   2040
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   480
      TabIndex        =   18
      Top             =   0
      Width           =   14175
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   2355
         TabIndex        =   19
         Top             =   840
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   5235
         TabIndex        =   21
         Top             =   840
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   8115
         TabIndex        =   24
         Top             =   840
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   6720
         TabIndex        =   4
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   10995
         TabIndex        =   27
         Top             =   840
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   9600
         TabIndex        =   5
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin TDBDate6Ctl.TDBDate DtDOAwal 
         Height          =   330
         Left            =   1680
         TabIndex        =   0
         Top             =   172
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmAuditService.frx":1E74
         Caption         =   "frmAuditService.frx":1FA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAuditService.frx":200C
         Keys            =   "frmAuditService.frx":202A
         Spin            =   "frmAuditService.frx":2088
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate DtDOAkhir 
         Height          =   330
         Left            =   3720
         TabIndex        =   1
         Top             =   172
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmAuditService.frx":20B0
         Caption         =   "frmAuditService.frx":21DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAuditService.frx":2248
         Keys            =   "frmAuditService.frx":2266
         Spin            =   "frmAuditService.frx":22C4
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   13995
         TabIndex        =   32
         Top             =   840
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   12600
         TabIndex        =   6
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   2355
         TabIndex        =   36
         Top             =   1200
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   270
         Left            =   5235
         TabIndex        =   37
         Top             =   1200
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   3840
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   8115
         TabIndex        =   38
         Top             =   1200
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   7
         Left            =   6720
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   10995
         TabIndex        =   39
         Top             =   1200
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   9600
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   270
         Left            =   13995
         TabIndex        =   40
         Top             =   1200
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   9
         Left            =   12600
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin ucXPButton.XPButton cmdExport 
         Height          =   495
         Left            =   11640
         TabIndex        =   57
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&o"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAuditService.frx":22EC
         PICN            =   "frmAuditService.frx":2308
         PICH            =   "frmAuditService.frx":289D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdAudit 
         Height          =   495
         Left            =   12840
         TabIndex        =   58
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&o"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAuditService.frx":2ED8
         PICN            =   "frmAuditService.frx":2EF4
         PICH            =   "frmAuditService.frx":348D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( pada fiskal yang sama )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5520
         TabIndex        =   54
         Top             =   210
         Width           =   2490
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1230
         Width           =   765
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   48
         Top             =   1245
         Width           =   765
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".Or."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   47
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5880
         TabIndex        =   46
         Top             =   1245
         Width           =   765
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".Or."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   45
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8760
         TabIndex        =   44
         Top             =   1245
         Width           =   765
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".Or."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8280
         TabIndex        =   43
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".Or."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11280
         TabIndex        =   42
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11760
         TabIndex        =   41
         Top             =   1245
         Width           =   765
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11760
         TabIndex        =   34
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".Or."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11280
         TabIndex        =   33
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".sd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3225
         TabIndex        =   31
         Top             =   217
         Width           =   360
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Transaksi :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".Or."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8280
         TabIndex        =   29
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8760
         TabIndex        =   28
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".Or."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   26
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5880
         TabIndex        =   25
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".Or."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   23
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   22
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Akun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   870
         Width           =   765
      End
   End
   Begin TrueOleDBGrid70.TDBGrid GridAudit 
      Height          =   7455
      Left            =   480
      TabIndex        =   35
      Top             =   2040
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13150
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1).NumberFormat=   "Standard"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "Standard"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12577002
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.alignment=3"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=82,.parent=13,.alignment=1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1,.locked=0"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin ucXPButton.XPButton TblKeluar 
      Height          =   495
      Left            =   12960
      TabIndex        =   56
      Top             =   9000
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&x"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAuditService.frx":3AB5
      PICN            =   "frmAuditService.frx":3AD1
      PICH            =   "frmAuditService.frx":3FA5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmAuditService.frx":44F0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblAudit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".Or."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   480
      TabIndex        =   52
      Top             =   1680
      Width           =   405
   End
End
Attribute VB_Name = "frmAuditService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsGrid As New ADODB.Recordset
Private cols As TrueOleDBGrid70.Columns

Private Sub clearForm()
    DtDOAwal = TglDDMMYYYY(isnow)
    DtDOAkhir = TglDDMMYYYY(isnow)
    
    For i = 0 To 9
        Text(i) = ""
    Next
    
    opTipe(0).value = True
    opTipe(1).value = False
End Sub

Private Sub cmdAudit_Click()
Dim boleh As Boolean
Dim pesan As String
Dim idxHead As Integer
Dim strFilter As String
Dim sSQLa As String


    boleh = True
    pesan = ""
    sSQLa = ""
    sSQL1 = ""
    strFilter = ""
    If Trim(Text(0)) = "" And Trim(Text(1)) = "" And Trim(Text(2)) = "" And Trim(Text(3)) = "" And Trim(Text(4)) = "" And Trim(Text(5)) = "" And Trim(Text(6)) = "" And Trim(Text(7)) = "" And Trim(Text(8)) = "" And Trim(Text(9)) = "" Then
        boleh = False
        pesan = pesan + "Kode akun harus diisi minimal 1. "
    End If
    If IsNull(DtDOAwal) Or IsNull(DtDOAkhir) Then
        boleh = False
        pesan = pesan + vbCr + "Range Tanggal harus diisi. "
    End If
    
    Set cols = GridAudit.Columns
    For Each col In cols
      col.FilterText = ""
    Next col
    
    If boleh Then
        If opField(0).value Then
            ' Nilai jasa
            If opTipe(0).value Then
                ' ledger >> Service
                idxHead = 0
                lblAudit = "Display Jasa (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
                sSQLa = "SELECT NOBUKTI AS NOINV, isnull(sum(JLHKREDIT-JLHDEBIT),0) AS nominalGL, CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN ROUND(isnull(SUM(LABOR*100/110),0),2) ELSE ROUND(isnull(SUM(LABOR*100/110),0),2) * -1 END AS JASA, cast(round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN ROUND(isnull(SUM(LABOR*100/110),0),2) ELSE ROUND(isnull(SUM(LABOR*100/110),0),2) * -1 END)),0),2) as float) AS SELISIH FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN A WITH(NOLOCK) " & _
                       "LEFT JOIN T_INVOICE B WITH(NOLOCK) ON ((B.NOINV = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2)) OR (B.FS_NO_RETUR = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2))) WHERE TGLTRN BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' AND " & _
                       "(KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "') GROUP BY NOBUKTI, FS_NO_RETUR "
                strFilter = " HAVING (cast(round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN ROUND(isnull(SUM(LABOR*100/110),0),2) ELSE ROUND(isnull(SUM(LABOR*100/110),0),2) * -1 END)),0),2) as float)) <> 0"
                sSQL1 = " ORDER BY NOBUKTI"
            ElseIf opTipe(1).value Then
                ' Service >> Ledger
                idxHead = 3
                lblAudit = "Display Jasa (Service  >>  Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
                sSQLa = "SELECT '" & Trim(gKodeFaktur) & "'+NOINV AS NOINV, (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "')) AS nominalGL, isnull(ROUND(SUM(LABOR*100/110),2),0) AS JASA, round(isnull((ISNULL(ROUND(SUM(LABOR*100/110),2),0) - (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "'))),0),2) AS SELISIH FROM T_INVOICE A WITH(NOLOCK) " & _
                       "WHERE INVDATE BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' GROUP BY NOINV "
                strFilter = " HAVING (round(isnull((ISNULL(SUM(LABOR*100/110),0) - (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                            "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "'))),0),2)) <> 0 "
                sSQL1 = " ORDER BY NOINV"
            End If
        ElseIf opField(1).value Then
            ' Nilai Part
            If opTipe(0).value Then
                ' ledger >> Service
                idxHead = 1
                lblAudit = "Display Part (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
                sSQLa = "SELECT NOBUKTI AS NOINV, isnull(sum(JLHKREDIT-JLHDEBIT),0) AS nominalGL, CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN ROUND(isnull(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0),2) ELSE ROUND(isnull(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0),2) * -1 END AS PART, round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN ROUND(isnull(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0),2) ELSE ROUND(isnull(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0),2) * -1 END)),0),2) AS SELISIH FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN A WITH(NOLOCK) " & _
                       "LEFT JOIN T_INVOICE B WITH(NOLOCK) ON ((B.NOINV = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2)) OR (B.FS_NO_RETUR = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2))) WHERE TGLTRN BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' AND " & _
                       "(KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "') GROUP BY NOBUKTI, FS_NO_RETUR "
                strFilter = " HAVING (round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN ROUND(isnull(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0),2) ELSE ROUND(isnull(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0),2) * -1 END)),0),2)) <> 0 "
                sSQL1 = " ORDER BY NOBUKTI"
            ElseIf opTipe(1).value Then
                ' Service >> Ledger
                idxHead = 4
                lblAudit = "Display Part (Service  >>  Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
                sSQLa = "SELECT '" & Trim(gKodeFaktur) & "'+NOINV AS NOINV, (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "')) AS nominalGL, isnull(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0) AS PART, round(isnull((ISNULL(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0) - (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "'))),0),2) AS SELISIH FROM T_INVOICE A WITH(NOLOCK) " & _
                       "WHERE INVDATE BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' GROUP BY NOINV "
                strFilter = " HAVING (round(isnull((ISNULL(SUM((PART_G+PART_NG+OIL_G+OIL_NG)*100/110),0) - (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                            "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "'))),0),2)) <> 0 "
                sSQL1 = " ORDER BY NOINV"
            End If
        ElseIf opField(2).value Then
            ' Nilai Sublet
            If opTipe(0).value Then
                ' ledger >> Service
                idxHead = 2
                lblAudit = "Display Sublet (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
'                sSQLa = "SELECT NOBUKTI AS NOINV, isnull(sum(JLHKREDIT-JLHDEBIT),0) AS nominalGL, CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN ROUND(isnull(SUM(SBLT*100/110),0),2) ELSE ROUND(isnull(SUM(SBLT*100/110),0),2) * -1 END AS SUBLET, round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN ROUND(isnull(SUM(SBLT*100/110),0),2) ELSE ROUND(isnull(SUM(SBLT*100/110),0),2) * -1 END)),0),2) AS SELISIH FROM [" & Left(NisDb, InStr(NisDb, "-")) & Year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN A " & _
'                       "LEFT JOIN T_INVOICE B ON ((B.NOINV = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2)) OR (B.FS_NO_RETUR = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2))) WHERE TGLTRN BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' AND " & _
'                       "(KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
'                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "') GROUP BY NOBUKTI, FS_NO_RETUR ORDER BY NOBUKTI"
                       
                sSQLa = "SELECT NOBUKTI AS NOINV, isnull(sum(JLHKREDIT-JLHDEBIT),0) AS nominalGL, CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN (CASE WHEN LEFT(NOBUKTI,1) = 'S' THEN (SELECT ROUND(ISNULL(SUM(HARGA)*100/110,0),0) * -1 FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = LEFT(SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2),8)) ELSE ROUND(isnull(SUM(SBLT*100/110),0),2) END) ELSE (CASE WHEN LEFT(NOBUKTI,1) = 'S' THEN (SELECT ROUND(ISNULL(SUM(HARGA)*100/110,0),0) FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = LEFT(SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2),8)) ELSE ROUND(isnull(SUM(SBLT*100/110),0),2) * -1 END) END AS SUBLET,  " & _
                       "round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN (CASE WHEN LEFT(NOBUKTI,1) = 'S' THEN (SELECT ROUND(ISNULL(SUM(HARGA)*100/110,0),0) * -1 FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = LEFT(SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2),8)) ELSE ROUND(isnull(SUM(SBLT*100/110),0),2) END) ELSE (CASE WHEN LEFT(NOBUKTI,1) = 'S' THEN (SELECT ROUND(ISNULL(SUM(HARGA)*100/110,0),0) FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = LEFT(SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2),8)) ELSE ROUND(isnull(SUM(SBLT*100/110),0),2) * -1 END) END)),0),2) AS SELISIH  " & _
                       "FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN A WITH(NOLOCK) " & _
                       "LEFT JOIN T_INVOICE B WITH(NOLOCK) ON ((B.NOINV = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2)) OR (B.FS_NO_RETUR = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2))) WHERE TGLTRN BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' AND " & _
                       "(KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "') GROUP BY NOBUKTI, FS_NO_RETUR "
                strFilter = " HAVING (round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN (CASE WHEN LEFT(NOBUKTI,1) = 'S' THEN (SELECT ROUND(ISNULL(SUM(HARGA)*100/110,0),0) * -1 FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = LEFT(SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2),8)) ELSE ROUND(isnull(SUM(SBLT*100/110),0),2) END) ELSE (CASE WHEN LEFT(NOBUKTI,1) = 'S' THEN (SELECT ROUND(ISNULL(SUM(HARGA)*100/110,0),0) FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = LEFT(SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2),8)) ELSE ROUND(isnull(SUM(SBLT*100/110),0),2) * -1 END) END)),0),2)) <> 0 "
                sSQL1 = " ORDER BY NOBUKTI"
                
            ElseIf opTipe(1).value Then
                ' Service >> Ledger
                idxHead = 5
                lblAudit = "Display Sublet (Service  >>  Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
                sSQLa = "SELECT '" & Trim(gKodeFaktur) & "'+NOINV AS NOINV, (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "')) AS nominalGL, isnull(SUM(SBLT*100/110),0) AS SUBLET, round(isnull((ISNULL(SUM(SBLT*100/110),0) - (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "'))),0),2) AS SELISIH FROM T_INVOICE A WITH(NOLOCK) " & _
                       "WHERE INVDATE BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' GROUP BY NOINV "
                
                
                strFilter = " HAVING (round(isnull((ISNULL(SUM(SBLT*100/110),0) - (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                            "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "'))),0),2)) <> 0 "
                sSQL1 = " ORDER BY NOINV"
                
            End If
        ElseIf opField(3).value Then
            ' Nilai Jasa + Sublet
            If opTipe(0).value Then
                ' ledger >> Service
                idxHead = 6
                lblAudit = "Display Jasa+Sublet (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
                sSQLa = "SELECT NOBUKTI AS NOINV, isnull(sum(JLHKREDIT-JLHDEBIT),0) AS nominalGL, CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN isnull(SUM((LABOR+SBLT)*100/110),0) ELSE isnull(SUM((LABOR+SBLT)*100/110),0) * -1 END AS JASA, round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN isnull(SUM((LABOR+SBLT)*100/110),0) ELSE isnull(SUM((LABOR+SBLT)*100/110),0) * -1 END)),0),2) AS SELISIH FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN A WITH(NOLOCK) " & _
                       "LEFT JOIN T_INVOICE B WITH(NOLOCK) ON ((B.NOINV = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2) OR (B.FS_NO_RETUR = SUBSTRING(A.NOBUKTI,3,LEN(A.NOBUKTI)-2))) WHERE TGLTRN BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' AND " & _
                       "(KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "') GROUP BY NOBUKTI, FS_NO_RETUR "
                strFilter = " HAVING (round(isnull((ISNULL(sum(JLHKREDIT-JLHDEBIT),0) - (CASE WHEN RIGHT(NOBUKTI,1) <> 'R' THEN isnull(SUM((LABOR+SBLT)*100/110),0) ELSE isnull(SUM((LABOR+SBLT)*100/110),0) * -1 END)),0),2)) <> 0 "
                sSQL1 = " ORDER BY NOBUKTI"
            ElseIf opTipe(1).value Then
                ' Service >> Ledger
                idxHead = 7
                lblAudit = "Display Jasa+Sublet (Service  >>  Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
                sSQLa = "SELECT '" & Trim(gKodeFaktur) & "'+NOINV AS NOINV, (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "')) AS nominalGL, isnull(SUM((LABOR+SBLT)*100/110),0) AS JASA, round(isnull((ISNULL(SUM((LABOR+SBLT)*100/110),0) - (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                       "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "'))),0),2) AS SELISIH FROM T_INVOICE A WITH(NOLOCK) " & _
                       "WHERE INVDATE BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' GROUP BY NOINV "
                strFilter = " HAVING (round(isnull((ISNULL(SUM((LABOR+SBLT)*100/110),0) - (SELECT isnull(sum(JLHKREDIT-JLHDEBIT),0) FROM [" & Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00") & "]..GLTRN WITH(NOLOCK) WHERE SUBSTRING(NOBUKTI,3,LEN(NOBUKTI)-2) = A.NOINV AND (KODEGL = '" & Trim(Text(0)) & "' OR KODEGL = '" & Trim(Text(1)) & "' OR KODEGL = '" & Trim(Text(2)) & "' OR KODEGL = '" & Trim(Text(3)) & "' OR KODEGL = '" & Trim(Text(4)) & "' OR KODEGL = '" & Trim(Text(5)) & "' " & _
                            "OR KODEGL = '" & Trim(Text(6)) & "' OR KODEGL = '" & Trim(Text(7)) & "' OR KODEGL = '" & Trim(Text(8)) & "' OR KODEGL = '" & Trim(Text(9)) & "'))),0),2)) <> 0 "
                sSQL1 = " ORDER BY NOINV"
            End If
        End If
        
        Set rsGrid = objO.FetchData(sSQLa & IIf(chIncNol.value, "", strFilter) & sSQL1)
        If Not rsGrid.EOF Then
            GridAudit.DataSource = rsGrid
        End If
        Call headAudit(idxHead)
        'MsgBox "Lihat data"
    Else
        MsgBox pesan, vbCritical, "Kelengkapan"
    End If
End Sub

Private Sub headAudit(idx As Integer)
    
    With GridAudit
        Select Case idx
            Case 0
                ' JASA [GL >> SVC]
                .Columns(0).Caption = "No. Invoice"
                .Columns(0).Width = 1200
                .Columns(0).Locked = True
                .Columns(0).DataField = "noinv"
                .Columns(1).Caption = "Nilai Ledger"
                .Columns(1).Width = 2000
                .Columns(1).Locked = True
                .Columns(1).DataField = "nominalGL"
                .Columns(2).Caption = "Nilai Jasa Service"
                .Columns(2).Width = 2000
                .Columns(2).Locked = True
                .Columns(2).DataField = "jasa"
                .Columns(3).Caption = "Selisih GL - SVC"
                .Columns(3).Width = 2000
                .Columns(3).Locked = True
                .Columns(3).DataField = "selisih"
            Case 1
                ' PART [GL >> SVC]
                .Columns(0).Caption = "No. Invoice"
                .Columns(0).Width = 1200
                .Columns(0).Locked = True
                .Columns(0).DataField = "noinv"
                .Columns(1).Caption = "Nilai Ledger"
                .Columns(1).Width = 2000
                .Columns(1).Locked = True
                .Columns(1).DataField = "nominalGL"
                .Columns(2).Caption = "Nilai Part Service"
                .Columns(2).Width = 2000
                .Columns(2).Locked = True
                .Columns(2).DataField = "part"
                .Columns(3).Caption = "Selisih GL - SVC"
                .Columns(3).Width = 2000
                .Columns(3).Locked = True
                .Columns(3).DataField = "selisih"
            Case 2
                ' SUBLET [GL >> SVC]
                .Columns(0).Caption = "No. Invoice"
                .Columns(0).Width = 1200
                .Columns(0).Locked = True
                .Columns(0).DataField = "noinv"
                .Columns(1).Caption = "Nilai Ledger"
                .Columns(1).Width = 2000
                .Columns(1).Locked = True
                .Columns(1).DataField = "nominalGL"
                .Columns(2).Caption = "Nilai Sublet Service"
                .Columns(2).Width = 2000
                .Columns(2).Locked = True
                .Columns(2).DataField = "sublet"
                .Columns(3).Caption = "Selisih GL - SVC"
                .Columns(3).Width = 2000
                .Columns(3).Locked = True
                .Columns(3).DataField = "selisih"
            Case 3
                ' JASA [SVC >> GL]
                .Columns(0).Caption = "No. Invoice"
                .Columns(0).Width = 1200
                .Columns(0).Locked = True
                .Columns(0).DataField = "noinv"
                .Columns(1).Caption = "Nilai Jasa Service"
                .Columns(1).Width = 2000
                .Columns(1).Locked = True
                .Columns(1).DataField = "jasa"
                .Columns(2).Caption = "Nilai Ledger"
                .Columns(2).Width = 2000
                .Columns(2).Locked = True
                .Columns(2).DataField = "nominalGL"
                .Columns(3).Caption = "Selisih SVC - GL"
                .Columns(3).Width = 2000
                .Columns(3).Locked = True
                .Columns(3).DataField = "selisih"
            Case 4
                ' PART [SVC >> GL]
                .Columns(0).Caption = "No. Invoice"
                .Columns(0).Width = 1200
                .Columns(0).Locked = True
                .Columns(0).DataField = "noinv"
                .Columns(1).Caption = "Nilai Part Service"
                .Columns(1).Width = 2000
                .Columns(1).Locked = True
                .Columns(1).DataField = "part"
                .Columns(2).Caption = "Nilai Ledger"
                .Columns(2).Width = 2000
                .Columns(2).Locked = True
                .Columns(2).DataField = "nominalGL"
                .Columns(3).Caption = "Selisih SVC - GL"
                .Columns(3).Width = 2000
                .Columns(3).Locked = True
                .Columns(3).DataField = "selisih"
            Case 5
                ' SUBLET [SVC >> GL]
                .Columns(0).Caption = "No. Invoice"
                .Columns(0).Width = 1200
                .Columns(0).Locked = True
                .Columns(0).DataField = "noinv"
                .Columns(1).Caption = "Nilai Sublet Service"
                .Columns(1).Width = 2000
                .Columns(1).Locked = True
                .Columns(1).DataField = "sublet"
                .Columns(2).Caption = "Nilai Ledger"
                .Columns(2).Width = 2000
                .Columns(2).Locked = True
                .Columns(2).DataField = "nominalGL"
                .Columns(3).Caption = "Selisih SVC - GL"
                .Columns(3).Width = 2000
                .Columns(3).Locked = True
                .Columns(3).DataField = "selisih"
            Case 6
                ' SUBLET [GL >> SVC]
                .Columns(0).Caption = "No. Invoice"
                .Columns(0).Width = 1200
                .Columns(0).Locked = True
                .Columns(0).DataField = "noinv"
                .Columns(1).Caption = "Nilai Jasa Service"
                .Columns(1).Width = 2000
                .Columns(1).Locked = True
                .Columns(1).DataField = "jasa"
                .Columns(2).Caption = "Nilai Ledger"
                .Columns(2).Width = 2000
                .Columns(2).Locked = True
                .Columns(2).DataField = "nominalGL"
                .Columns(3).Caption = "Selisih GL - SVC"
                .Columns(3).Width = 2000
                .Columns(3).Locked = True
                .Columns(3).DataField = "selisih"
            Case 7
                ' JASA + SUBLET [SVC >> GL]
                .Columns(0).Caption = "No. Invoice"
                .Columns(0).Width = 1200
                .Columns(0).Locked = True
                .Columns(0).DataField = "noinv"
                .Columns(1).Caption = "Nilai Jasa Service"
                .Columns(1).Width = 2000
                .Columns(1).Locked = True
                .Columns(1).DataField = "jasa"
                .Columns(2).Caption = "Nilai Ledger"
                .Columns(2).Width = 2000
                .Columns(2).Locked = True
                .Columns(2).DataField = "nominalGL"
                .Columns(3).Caption = "Selisih SVC - GL"
                .Columns(3).Width = 2000
                .Columns(3).Locked = True
                .Columns(3).DataField = "selisih"
        End Select
    End With
End Sub

Private Sub cmdExport_Click()
On Error GoTo Err_Export
    Dim Er As New ErrInfo
    If rsGrid.RecordCount = 0 Then
        Er.HeadCap = "Export To Excel !"
        Er.InfoCap = "List Audit Kosong !"
        Beep
        Er.Show (vbModal)
        Exit Sub
    End If
    Dim MyPath As String
    Dim onVal As String
    MyPath = ""
    Dim f As New FormExplorer
    f.Show (vbModal)
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    onVal = MyPath & f.FileName & ".XLS"
    
     GridAudit.ExportToFile onVal, True
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            
            Beep
            Er.Show (vbModal)
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:

End Sub

Private Sub Form_Load()
    Call clearForm
End Sub

Private Sub GridAudit_FilterChange()
On Error GoTo err
   Set cols = GridAudit.Columns
   Dim C As Integer

   C = GridAudit.col
   GridAudit.HoldFields
   rsGrid.filter = getFilter()
   If rsGrid.EOF = True Then
      For Each col In GridAudit.Columns
         col.FilterText = ""
      Next col
      rsGrid.filter = ""
   End If
   GridAudit.col = C
   GridAudit.EditActive = True
   Exit Sub
err:
   For Each col In cols
      col.FilterText = ""
   Next col
   Resume Next
End Sub

Private Function getFilter() As String
   On Error GoTo err

On Error GoTo err
Dim tmp As String
Dim n As Integer

   For Each col In cols
      If Trim(col.FilterText) <> "" Then
         n = n + 1
         If n > 1 Then
            tmp = tmp & " AND "
         End If
            tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp

   Exit Function
err:
   Resume Next
End Function

Private Sub GridAudit_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 3 Then rsGrid.Sort = GridAudit.Columns(ColIndex).DataField & " " & Urut
4        GridAudit.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
End Sub

Private Sub opField_Click(Index As Integer)
    Call cmdAudit_Click
    Select Case Index
        Case 0
            If opTipe(0).value Then
                lblAudit = "Display Jasa (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opTipe(1).value Then
                lblAudit = "Display Jasa (Service >> Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            End If
        Case 1
            If opTipe(0).value Then
                lblAudit = "Display Part (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opTipe(1).value Then
                lblAudit = "Display Part (Service >> Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            End If
        Case 2
            If opTipe(0).value Then
                lblAudit = "Display Sublet (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opTipe(1).value Then
                lblAudit = "Display Sublet (Service >> Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            End If
        Case 3
            If opTipe(0).value Then
                lblAudit = "Display Jasa+Sublet (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opTipe(1).value Then
                lblAudit = "Display Jasa+Sublet (Service >> Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            End If
    End Select
End Sub

Private Sub opTipe_Click(Index As Integer)
    Call cmdAudit_Click
    
    Select Case Index
        Case 0
            If opField(0).value Then
                lblAudit = "Display Jasa (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opField(1).value Then
                lblAudit = "Display Part (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opField(2).value Then
                lblAudit = "Display Sublet (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opField(3).value Then
                lblAudit = "Display Jasa+Sublet (Ledger  >>  Service) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            End If
        Case 1
            If opField(0).value Then
                lblAudit = "Display Jasa (Service  >>  Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opField(1).value Then
                lblAudit = "Display Part (Service  >>  Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opField(2).value Then
                lblAudit = "Display Sublet (Service  >>  Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            ElseIf opField(3).value Then
                lblAudit = "Display Jasa+Sublet (Service  >>  Ledger) ... Tanggal : " & Format(DtDOAwal, "DD/MMMM/YYYY") & " - " & Format(DtDOAkhir, "DD/MMMM/YYYY")
            End If
    End Select
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index).Mask = ""
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If

    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    Select Case KeyCode
        Case vbKeyF12
            Select Case Index
                Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
                    NisDb2 = Left(NisDb, InStr(NisDb, "-")) & year(DtDOAwal) & Format(Month(DtDOAwal), "00")
                    gMainKeyCari = Text(Index)
                    gSearchOption = "KODEJNRL"
                    frmCariAcess.Show (vbModal)
            End Select
    End Select
End Sub
