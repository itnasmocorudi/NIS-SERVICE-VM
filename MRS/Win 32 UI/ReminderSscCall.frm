VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form ReminderSscCall 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   1500
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   Picture         =   "ReminderSscCall.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Picture         =   "ReminderSscCall.frx":EEF6
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton cmdRecord 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Record"
      Height          =   732
      Left            =   10560
      Picture         =   "ReminderSscCall.frx":F686
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   7440
      Width           =   735
   End
   Begin VB.Frame FrameGRP 
      BorderStyle     =   0  'None
      Caption         =   ".:: SELECTED ITEM REMAINDER::."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4815
      Left            =   480
      TabIndex        =   30
      Top             =   1080
      Width           =   11655
      Begin VB.TextBox txtKodePos 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   47
         Top             =   1035
         Width           =   1575
      End
      Begin VB.TextBox txtKota 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   46
         Top             =   1035
         Width           =   2655
      End
      Begin VB.TextBox txtNoEngine 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   45
         Top             =   1395
         Width           =   2655
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   44
         Top             =   675
         Width           =   6015
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   43
         Top             =   315
         Width           =   6015
      End
      Begin VB.TextBox txtNoPol 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   42
         Top             =   1755
         Width           =   1575
      End
      Begin VB.TextBox txtNoChassis 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   41
         Top             =   1755
         Width           =   2655
      End
      Begin VB.TextBox txtGSM 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         MaxLength       =   100
         TabIndex        =   40
         Top             =   1035
         Width           =   1815
      End
      Begin VB.TextBox txtTelpon 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   9600
         MaxLength       =   100
         TabIndex        =   39
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox txtCDMA 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         MaxLength       =   100
         TabIndex        =   38
         Top             =   1395
         Width           =   1815
      End
      Begin VB.Timer Timer1 
         Interval        =   150
         Left            =   9360
         Top             =   4200
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   37
         Top             =   3720
         Width           =   7215
      End
      Begin VB.TextBox txtId_Warranty 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   36
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtMDL 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   35
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtWarranty 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   34
         Top             =   3360
         Width           =   7215
      End
      Begin VB.TextBox txtModel 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4800
         MaxLength       =   100
         TabIndex        =   33
         Top             =   3000
         Width           =   4455
      End
      Begin TDBDate6Ctl.TDBDate DtMulai 
         Height          =   315
         Left            =   3240
         TabIndex        =   48
         Top             =   4200
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         Calendar        =   "ReminderSscCall.frx":FCF3
         Caption         =   "ReminderSscCall.frx":FE1F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ReminderSscCall.frx":FE8B
         Keys            =   "ReminderSscCall.frx":FEA9
         Spin            =   "ReminderSscCall.frx":FF07
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   0
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
      Begin TDBDate6Ctl.TDBDate DtSelesai 
         Height          =   315
         Left            =   5400
         TabIndex        =   49
         Top             =   4200
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         Calendar        =   "ReminderSscCall.frx":FF2F
         Caption         =   "ReminderSscCall.frx":1005B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ReminderSscCall.frx":100C7
         Keys            =   "ReminderSscCall.frx":100E5
         Spin            =   "ReminderSscCall.frx":10143
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   0
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
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   3840
         X2              =   7560
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   ".:: ITEM SSC WARRANTY ::."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3840
         TabIndex        =   82
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   81
         Top             =   1035
         Width           =   900
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   6360
         TabIndex        =   80
         Top             =   1035
         Width           =   75
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   79
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   78
         Top             =   1035
         Width           =   75
      End
      Begin VB.Line Line21 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11880
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   77
         Top             =   1395
         Width           =   75
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Engine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   76
         Top             =   1395
         Width           =   990
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Masa Berlaku Waranty SSC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   75
         Top             =   4230
         Width           =   2550
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   6360
         TabIndex        =   74
         Top             =   1755
         Width           =   75
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   73
         Top             =   1755
         Width           =   855
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama  Cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   72
         Top             =   315
         Width           =   1110
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   71
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   70
         Top             =   315
         Width           =   75
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   69
         Top             =   675
         Width           =   75
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   68
         Top             =   1755
         Width           =   75
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Chassis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   67
         Top             =   1755
         Width           =   1065
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   9480
         TabIndex        =   66
         Top             =   1035
         Width           =   75
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   9480
         TabIndex        =   65
         Top             =   675
         Width           =   75
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8640
         TabIndex        =   64
         Top             =   1035
         Width           =   435
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8640
         TabIndex        =   63
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   9480
         TabIndex        =   62
         Top             =   1395
         Width           =   75
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8640
         TabIndex        =   61
         Top             =   1395
         Width           =   600
      End
      Begin VB.Line Line20 
         BorderColor     =   &H000000FF&
         X1              =   8640
         X2              =   11400
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: TELPON ::."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8640
         TabIndex        =   60
         Top             =   315
         Width           =   2775
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3000
         TabIndex        =   59
         Top             =   4200
         Width           =   75
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   5040
         X2              =   5280
         Y1              =   4335
         Y2              =   4335
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11880
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lblBerlaku 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berlaku Selamanya"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7320
         TabIndex        =   58
         Top             =   4230
         Width           =   1800
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   57
         Top             =   3720
         Width           =   1110
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   56
         Top             =   3720
         Width           =   75
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   55
         Top             =   3000
         Width           =   75
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   54
         Top             =   2640
         Width           =   75
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   53
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id Warranty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   52
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   51
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   50
         Top             =   3360
         Width           =   75
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         X1              =   240
         X2              =   11880
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label40 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   ".:: SELECTED CUSTOMER AS USER ::."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3960
         TabIndex        =   32
         Top             =   0
         Width           =   4095
      End
      Begin VB.Label lblMerk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   2520
         TabIndex        =   31
         Top             =   360
         Width           =   75
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   3840
         X2              =   7560
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   480
      TabIndex        =   20
      Top             =   3520
      Width           =   11655
      Begin VB.ComboBox cboJam 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6855
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   525
         Width           =   735
      End
      Begin VB.ComboBox cboMenit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8055
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   525
         Width           =   735
      End
      Begin TDBNumber6Ctl.TDBNumber txtHargaJasa 
         Height          =   330
         Left            =   7200
         TabIndex        =   23
         Top             =   120
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   582
         Calculator      =   "ReminderSscCall.frx":1016B
         Caption         =   "ReminderSscCall.frx":1018B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ReminderSscCall.frx":101F7
         Keys            =   "ReminderSscCall.frx":10215
         Spin            =   "ReminderSscCall.frx":1025F
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1397882885
         MinValueVT      =   1112342533
      End
      Begin TDBNumber6Ctl.TDBNumber txtHargaPart 
         Height          =   330
         Left            =   9480
         TabIndex        =   24
         Top             =   120
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   582
         Calculator      =   "ReminderSscCall.frx":10287
         Caption         =   "ReminderSscCall.frx":102A7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ReminderSscCall.frx":10313
         Keys            =   "ReminderSscCall.frx":10331
         Spin            =   "ReminderSscCall.frx":1037B
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1397882885
         MinValueVT      =   1112342533
      End
      Begin VB.Shape Shape6 
         Height          =   345
         Left            =   8040
         Top             =   510
         Width           =   765
      End
      Begin VB.Shape Shape5 
         Height          =   345
         Left            =   6840
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   6720
         TabIndex        =   28
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   6720
         TabIndex        =   27
         Top             =   510
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimasi Harga        Rp.                              + Rp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4920
         TabIndex        =   26
         Top             =   120
         Width           =   4530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimasi Waktu                    Jam              Menit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4920
         TabIndex        =   25
         Top             =   540
         Width           =   4500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   480
      TabIndex        =   13
      Top             =   500
      Width           =   11655
      Begin VB.TextBox txtKetTelp 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   14
         Top             =   80
         Width           =   5055
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   2040
         TabIndex        =   15
         Top             =   90
         Width           =   4455
         Begin VB.OptionButton OptTelp 
            Caption         =   "Diterima"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton OptTelp 
            Caption         =   "Tidak Diterima"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2160
            TabIndex        =   16
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   19
         Top             =   135
         Width           =   75
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirm Telp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   1245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   3000
         Y2              =   3000
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   5880
      Width           =   11655
      Begin VB.TextBox txtPesan2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   149
         TabIndex        =   2
         Top             =   1320
         Width           =   7935
      End
      Begin VB.OptionButton OptApp 
         Caption         =   "Ya"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox txtApp 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         MaxLength       =   149
         TabIndex        =   4
         Top             =   360
         Width           =   7935
      End
      Begin VB.CommandButton cmdBooking 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Booking"
         Enabled         =   0   'False
         Height          =   732
         Left            =   2520
         Picture         =   "ReminderSscCall.frx":103A3
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox txtPesan1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   149
         TabIndex        =   1
         Top             =   960
         Width           =   7935
      End
      Begin VB.OptionButton OptApp 
         Caption         =   "Tidak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   6
         Top             =   0
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   9960
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lCall 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9960
         TabIndex        =   84
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appoinment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1680
         TabIndex        =   11
         Top             =   45
         Width           =   75
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lCall1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   10080
         TabIndex        =   10
         Top             =   0
         Width           =   120
      End
      Begin VB.Label lCall2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10095
         TabIndex        =   9
         Top             =   15
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   75
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   10440
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   10320
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "   Promosikan Bengkel - Sistem Appointment (sarankan pelanggan untuk melakukan appointment)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   83
      Top             =   7920
      Width           =   9855
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   10320
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM REMINDER SSC - CALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   29
      Top             =   120
      Width           =   10215
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   615
      X2              =   4080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1680
      X2              =   4080
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2655
      X2              =   4080
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11685
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000F&
      Height          =   8415
      Left            =   480
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "ReminderSscCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private rsItem As ADODB.Recordset
Private objSSC As New ClsSSC
Private Rec As Boolean

Public Property Let rsItemCall(value As ADODB.Recordset)
   Set rsItem = value
End Property

Private Sub cmdBooking_Click()
    With frmBooking
        Dim nDpn1 As String
        Dim nTgh1 As String
        Dim nBlk1 As String
        Call pecahNoPol(txtNoPol.Text, nDpn1, nTgh1, nBlk1)
        .Text(0) = Trim(nDpn1)
        .Text(56) = Trim(nTgh1)
        .Text(57) = Trim(nBlk1)
        
        .Show vbModal
    End With
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Rec = False
   Call BindToControlSSC(rsItem)
   Call RecordItem
   Me.OptApp(0).Enabled = False
   Me.OptApp(1).Enabled = False
End Sub

Private Sub BindToControlSSC(rs As ADODB.Recordset)
   With Me
      .txtNama.Text = rs.Fields(0).value
      .txtAlamat.Text = rs.Fields(1).value
      .txtKota.Text = rs.Fields(2).value
      .txtKodePos.Text = rs.Fields(3).value
      .txtNoPol.Text = rs.Fields(4).value
      .txtNoEngine.Text = rs.Fields(5).value
      .txtNoChassis.Text = rs.Fields(6).value
      .txtTelpon.Text = rs.Fields(7).value
      .txtGSM.Text = rs.Fields(8).value
      .txtCDMA.Text = rs.Fields(9).value
      .txtId_Warranty.Text = rs.Fields(10).value
      .txtMDL.Text = rs.Fields(11).value
      .txtModel.Text = rs.Fields(12).value
      .txtWarranty.Text = rs.Fields(13).value
      .txtKeterangan.Text = rs.Fields(14).value
      .DtMulai.value = Format(rs.Fields(15).value, "DD/MM/YYYY")
      .DtSelesai.value = Format(rs.Fields(16).value, "DD/MM/YYYY")
   End With
   
   If Me.DtMulai.value = Me.DtSelesai.value Then
      Me.lblBerlaku.Caption = "Lifetime Warranty"
   Else
      Dim hari As Double
      hari = DateDiff("m", DtMulai.value, DtSelesai.value)
      Me.lblBerlaku.Caption = "Berlaku " & hari & " Bulan"
   End If
End Sub

Private Sub OptApp_Click(Index As Integer)
   If OptApp(0).value = True Then
      Me.txtApp.Enabled = True
      Me.cmdBooking.Enabled = False
   Else
      Me.txtApp.Enabled = False
      Me.cmdBooking.Enabled = True
   End If
End Sub

Private Sub OptTelp_Click(Index As Integer)
   If OptTelp(0).value = True Then
      Me.txtKetTelp.Enabled = True
      Me.OptApp(0).Enabled = False
      Me.OptApp(1).Enabled = False
      Me.OptApp(0).value = False
      Me.OptApp(1).value = False
   Else
      Me.txtKetTelp.Enabled = False
      Me.OptApp(0).Enabled = True
      Me.OptApp(1).Enabled = True
   End If
End Sub

Private Sub Timer1_Timer()
   Dim R, G, B
   R = 255 * Rnd
   G = 255 * Rnd
   B = 255 * Rnd
   Me.lblBerlaku.ForeColor = RGB(R, G, B)
End Sub

Private Sub RecordItem()
   Dim rsN As ADODB.Recordset
   Set rsN = objSSC.FetchRecordItem(txtNoChassis.Text, txtId_Warranty.Text, "CL")
   If rsN.RecordCount = 0 Then
      lCall.Caption = "CALL 1"
   Else
      lCall.Caption = "CALL " & rsN.RecordCount
      txtKetTelp.Text = NullText(rsN.Fields(3).value)
      txtPesan1.Text = NullText(rsN.Fields(6).value)
      txtPesan2.Text = NullText(rsN.Fields(7).value)
      txtApp.Text = NullText(rsN.Fields(10).value)
   End If
End Sub

Private Sub cmdRecord_Click()
   If MsgBox("Record Data SSC - " & lCall.Caption & " ?", vbQuestion + vbYesNo, "SSC") = vbYes Then
      Call objSSC.AddNew_LogRemainderSSC_Call(IIf(OptTelp(1).value = True, 1, 0), txtKetTelp.Text, "CL", _
                           Trim(txtNoChassis.Text), Trim(txtPesan1.Text), Trim(txtPesan2.Text), Trim(txtId_Warranty.Text), _
                           IIf(OptApp(1).value = True, "Y", "N"), Trim(txtApp.Text))
      If MsgBox("Keluar dari Remainder SSC - Call ?", vbQuestion + vbYesNo, "SSC") = vbYes Then
         Unload Me
      End If
      Rec = True
   End If
End Sub



Private Sub txtPesan1_Change()
   If Len(txtPesan1.Text) > 60 Then
      txtPesan2.SetFocus
   End If
End Sub

Private Sub txtPesan1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub
