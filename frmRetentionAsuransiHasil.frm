VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRetentionAsuransiHasil 
   Caption         =   "Retention Asuransi"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   13890
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Filter Data"
      Height          =   1935
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   10935
      Begin VB.ComboBox cbTipe 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmRetentionAsuransiHasil.frx":0000
         Left            =   1545
         List            =   "frmRetentionAsuransiHasil.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc AdoRetail 
         Height          =   375
         Left            =   9600
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbAsuransi 
         Height          =   375
         Left            =   7200
         TabIndex        =   15
         Top             =   720
         Width           =   2295
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBDate6Ctl.TDBDate Tgl1 
         Height          =   330
         Left            =   1545
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRetentionAsuransiHasil.frx":0041
         Caption         =   "frmRetentionAsuransiHasil.frx":016D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRetentionAsuransiHasil.frx":01D9
         Keys            =   "frmRetentionAsuransiHasil.frx":01F7
         Spin            =   "frmRetentionAsuransiHasil.frx":0255
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
      Begin TDBDate6Ctl.TDBDate Tgl2 
         Height          =   330
         Left            =   3465
         TabIndex        =   4
         Top             =   720
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRetentionAsuransiHasil.frx":027D
         Caption         =   "frmRetentionAsuransiHasil.frx":03A9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRetentionAsuransiHasil.frx":0415
         Keys            =   "frmRetentionAsuransiHasil.frx":0433
         Spin            =   "frmRetentionAsuransiHasil.frx":0491
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
      Begin TDBDate6Ctl.TDBDate tglpolish1 
         Height          =   330
         Left            =   1545
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRetentionAsuransiHasil.frx":04B9
         Caption         =   "frmRetentionAsuransiHasil.frx":05E5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRetentionAsuransiHasil.frx":0651
         Keys            =   "frmRetentionAsuransiHasil.frx":066F
         Spin            =   "frmRetentionAsuransiHasil.frx":06CD
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
      Begin TDBDate6Ctl.TDBDate tglpolish2 
         Height          =   330
         Left            =   3465
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRetentionAsuransiHasil.frx":06F5
         Caption         =   "frmRetentionAsuransiHasil.frx":0821
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRetentionAsuransiHasil.frx":088D
         Keys            =   "frmRetentionAsuransiHasil.frx":08AB
         Spin            =   "frmRetentionAsuransiHasil.frx":0909
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
      Begin TDBDate6Ctl.TDBDate tglReminder1 
         Height          =   330
         Left            =   1545
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRetentionAsuransiHasil.frx":0931
         Caption         =   "frmRetentionAsuransiHasil.frx":0A5D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRetentionAsuransiHasil.frx":0AC9
         Keys            =   "frmRetentionAsuransiHasil.frx":0AE7
         Spin            =   "frmRetentionAsuransiHasil.frx":0B45
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
      Begin TDBDate6Ctl.TDBDate tglReminder2 
         Height          =   330
         Left            =   3465
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRetentionAsuransiHasil.frx":0B6D
         Caption         =   "frmRetentionAsuransiHasil.frx":0C99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRetentionAsuransiHasil.frx":0D05
         Keys            =   "frmRetentionAsuransiHasil.frx":0D23
         Spin            =   "frmRetentionAsuransiHasil.frx":0D81
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbRetail 
         Height          =   375
         Left            =   7200
         TabIndex        =   17
         Top             =   1080
         Width           =   2295
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAsuransi 
         Height          =   375
         Left            =   9600
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin ucXPButton.XPButton cmdRef 
         Height          =   495
         Left            =   9600
         TabIndex        =   21
         Top             =   1320
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&r"
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
         MICON           =   "frmRetentionAsuransiHasil.frx":0DA9
         PICN            =   "frmRetentionAsuransiHasil.frx":0DC5
         PICH            =   "frmRetentionAsuransiHasil.frx":1341
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aktifitas"
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
         Left            =   840
         TabIndex        =   20
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Retail"
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
         Left            =   6105
         TabIndex        =   18
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asuransi"
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
         Left            =   6405
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Reminder"
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
         Left            =   135
         TabIndex        =   14
         Top             =   1515
         Width           =   1290
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3225
         TabIndex        =   13
         Top             =   1515
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Polis Aktif"
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
         TabIndex        =   10
         Top             =   1155
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3225
         TabIndex        =   9
         Top             =   1155
         Width           =   60
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Retail"
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
         Left            =   405
         TabIndex        =   6
         Top             =   795
         Width           =   1020
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3225
         TabIndex        =   5
         Top             =   795
         Width           =   60
      End
   End
   Begin ucXPButton.XPButton cmdKeluar 
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   ""
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRetentionAsuransiHasil.frx":1978
      PICN            =   "frmRetentionAsuransiHasil.frx":1994
      PICH            =   "frmRetentionAsuransiHasil.frx":82C7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBGrid70.TDBGrid gridAsuransi 
      Height          =   4095
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7223
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=3281"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   14933984
      RowDividerColor =   14933984
      RowSubDividerColor=   14933984
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=88,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=84,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=85,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=86,.parent=17"
      _StyleDefs(88)  =   "Named:id=33:Normal"
      _StyleDefs(89)  =   ":id=33,.parent=0"
      _StyleDefs(90)  =   "Named:id=34:Heading"
      _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(92)  =   ":id=34,.wraptext=-1"
      _StyleDefs(93)  =   "Named:id=35:Footing"
      _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(95)  =   "Named:id=36:Selected"
      _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=37:Caption"
      _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(99)  =   "Named:id=38:HighlightRow"
      _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(101) =   "Named:id=39:EvenRow"
      _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&"
      _StyleDefs(103) =   "Named:id=40:OddRow"
      _StyleDefs(104) =   ":id=40,.parent=33"
      _StyleDefs(105) =   "Named:id=41:RecordSelector"
      _StyleDefs(106) =   ":id=41,.parent=34"
      _StyleDefs(107) =   "Named:id=42:FilterBar"
      _StyleDefs(108) =   ":id=42,.parent=33"
      _StyleDefs(109) =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(110) =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(111) =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(112) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(113) =   "Named:id=87:SudahDitagih"
      _StyleDefs(114) =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(115) =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(116) =   ":id=87,.fontname=Tahoma"
      _StyleDefs(117) =   "Named:id=83:Booking"
      _StyleDefs(118) =   ":id=83,.parent=87,.fgcolor=&HFF00&"
   End
   Begin ucXPButton.XPButton CmdCall 
      Height          =   495
      Left            =   9030
      TabIndex        =   22
      Top             =   6360
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&f"
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
      MICON           =   "frmRetentionAsuransiHasil.frx":EF7E
      PICN            =   "frmRetentionAsuransiHasil.frx":EF9A
      PICH            =   "frmRetentionAsuransiHasil.frx":129E1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdExport 
      Height          =   495
      Left            =   7920
      TabIndex        =   23
      Top             =   6360
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
      MICON           =   "frmRetentionAsuransiHasil.frx":16653
      PICN            =   "frmRetentionAsuransiHasil.frx":1666F
      PICH            =   "frmRetentionAsuransiHasil.frx":16C04
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdValidasi 
      Height          =   495
      Left            =   8280
      TabIndex        =   24
      Top             =   7200
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&f"
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
      MICON           =   "frmRetentionAsuransiHasil.frx":1723F
      PICN            =   "frmRetentionAsuransiHasil.frx":1725B
      PICH            =   "frmRetentionAsuransiHasil.frx":1AE41
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7575
      Left            =   0
      Picture         =   "frmRetentionAsuransiHasil.frx":1EADE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRetentionAsuransiHasil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmRetentionAsuransiHasil
' DateTime  : 01/10/2020 10:06
' Author    : Rudi Kristanto
' Purpose   :
'---------------------------------------------------------------------------------------

 
Private rsTemp      As ADODB.Recordset
Private rsGrid      As New ADODB.Recordset
Private dealerBP    As String
Private nopol       As String
Private isCall      As Boolean
Dim cols            As TrueOleDBGrid70.Columns
Private Urut        As String


Private Sub cmdCall_Click()
    If rsGrid.RecordCount > 0 Then
        rsGrid.filter = "terlibat = true"
        If rsGrid.RecordCount > 0 Then
            If rsGrid.RecordCount = 1 Then
                'Dim fr1 As New frmRetensiAsuransiCall
                'fr1.Top = Me.Top
'                fr1.MyData = rsGrid
'                fr1.show
               ' If cbTipe.Text = "Upgrade Customer" Then
                    frmRetensiAsuransiCallHasil.MyIdAsuransi = IIf(NulSID(rsGrid!kode_asuransi) = "", 0, NulSID(rsGrid!kode_asuransi))
                'Else
                 '   frmRetensiAsuransiCallHasil.MyIdAsuransi = 0
                'End If
                frmRetensiAsuransiCallHasil.myAkt = NulSID(cbTipe.Text)
                frmRetensiAsuransiCallHasil.MyData = rsGrid
                frmRetensiAsuransiCallHasil.show 'vbModal, Me
                'Call cmdRef_Click
                
                isCall = True
            Else
                MsgBox "Pilih salah satu Data yang akan direminder", vbInformation
            End If
        Else
            MsgBox "Pilih Data yang akan direminder", vbInformation
        End If
    End If
kosongan:
    'rsGrid.filter = ""
End Sub

Private Sub cmdExport_Click()
1         On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsGrid.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Reminder Retensi Asuransi Kosong !"
5             Beep
6             Er.show (vbModal)
7             Exit Sub
8         End If
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".XLS"
16         gridAsuransi.ExportToFile onVal, True
17                Er.HeadCap = "Export to excel !"
18                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
19                Beep
20                Er.show (vbModal)
21         Exit Sub
Err_Export:
22          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
23          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdRef_Click()
    
    isCall = False
    If cbTipe.Text = "New Customer" Then
        Call refreshDataNewCustomer
    ElseIf cbTipe.Text = "Upgrade Customer" Then
        Call refreshDataUpgrade
    ElseIf cbTipe.Text = "Renew Customer" Then
        Call refreshDataRenew
    Else
        MsgBox "Silahkan pilih aktivitas dahulu.", vbInformation
    End If
End Sub

Private Sub Form_Activate()
    If isCall Then
        rsGrid.filter = ""
        Call cmdRef_Click
    End If
End Sub

Private Sub Form_Load()
    If Get_ConnectionProfil_Bengkel99 = False Then
        MsgBox "Ada Kesalahan Setting ke HO," & Chr(13) & "Cek VPN dan konektivitas ke HO" & Chr(13) & "(Lakukan Routing IP)", vbCritical + vbOKOnly
        Unload Me
    Else
        Call refreshCombo
    End If
End Sub

Private Sub refreshDataNewCustomer()
    Dim kriteriaRitel       As String
    Dim kriteriaReminder    As String
    
    If Tgl1 <> "" And Tgl2 <> "" Then
        kriteriaRitel = " and tgl_notadebet between '" & Format(Tgl1, "mm/dd/yyyy") & "' and '" & Format(Tgl2, "mm/dd/yyyy") & "'"
    End If
    
    If tglReminder1 <> "" And tglReminder2 <> "" Then
        kriteriaReminder = " (b.tglFU between '" & Format(tglReminder1, "mm/dd/yyyy") & "' and '" & Format(tglReminder2, "mm/dd/yyyy") & " 23:59:59') "
    End If
    sSQL = "SELECT a.nama_pemesan, a.alamat1+' '+a.alamat2 as alamat,a.Telp_rumah as Telepon, a.hp1, a.hp2, a.norangka,a.Nama_Faktur as NamaTT, ltrim(replace((Alamat1_Faktur + ' ' + Alamat2_Faktur + ' ' + Alamat3_Faktur),'  ','')) as AlamatTT, Kotamadya as KotaTT, a.HP1 as HP1TT, a.HP2 as HP2TT, c.namakategory, d.namaind, e.nomesin, e.kodetahun,e.kodekategori, "
    sSQL = sSQL & " f.kode_leasing, a.kode_asuransi, (select kodedealer+'|'+NamaProfil from [" & Log.LogServ & "].ProfilSPK.dbo.profil) as kode, a.tglpdd as PddFinal,a.kode_sales as kodesales, s.namasales, syarat_bayar,a.spk_no, Isnull(u.NoPolisi,'') NoPolisi, Isnull(OSS.masaBerlakuAwal,'1900-01-01') MasaBerlakuAwal, Isnull(OSS.masaBerlakuAkhir,'1900-01-01') MasaBerlakuAkhir,isnull(b.kodereq,'') kodereq"
    sSQL = sSQL & " FROM [" & Log.LogServ & "].SPK" & Right(gDb, 2) & ".DBO.SPK a"
    sSQL = sSQL & " left join [" & Log.LogServ & "].bengkel" & Right(gDb, 2) & ".dbo.FUAsuransi b on a.norangka = b.norangka and b.aktifitas = 'NEW'"
    sSQL = sSQL & " left join [" & Log.LogServ & "].spk" & Right(gDb, 2) & ".dbo.inkategory c on a.kode_type = c.kodekategory"
    sSQL = sSQL & " left join [" & Log.LogServ & "].spk" & Right(gDb, 2) & ".dbo.InWarnaExt d on a.kode_warna = d.kodewarna"
    sSQL = sSQL & " left join [" & Log.LogServ & "].spk" & Right(gDb, 2) & ".dbo.inmst e on a.norangka = e.kodeitem"
    sSQL = sSQL & " left join [" & Log.LogServ & "].spk" & Right(gDb, 2) & ".dbo.leasing f on a.spk_no = f.spk_no"
    sSQL = sSQL & " Left Join"
    sSQL = sSQL & " ("
    sSQL = sSQL & "     select nospk, norangka, masaBerlakuAwal, masaBerlakuAkhir , jenisAsuransi, jenisAsuransi2, jenisAsuransi3, jenisAsuransi4, jenisAsuransi5,"
    sSQL = sSQL & "     jenisAsuransi6, jenisAsuransi7 , 0 jenisasuransi8"
    sSQL = sSQL & "     from oss..Asuransi"
    sSQL = sSQL & " ) OSS on OSS.NoSPK = a.spk_no and OSS.noRangka = a.norangka"
    sSQL = sSQL & " left join singleidori..UnitCustomer u On a.NoRangka = u.NoRangka "
    sSQL = sSQL & " left join [" & Log.LogServ & "].spk" & Right(gDb, 2) & ".dbo.arsales s with (nolock) on a.kode_sales = s.kodesales"
    sSQL = sSQL & " where isnull(b.kodereq,'') <> '' and syarat_bayar = 0 and isnull(a.kode_asuransi,'') = '' and tgl_notadebet is not null"
    sSQL = sSQL & kriteriaRitel ''" and tgl_notadebet between '08/01/2020' and '08/31/2020'"
    sSQL = sSQL & " and  " & IIf(kriteriaReminder = "", "1=1", kriteriaReminder) & " " ''" and (b.tglFU >= '08/01/2020' or isnull(b.tglFU,'01/01/1900') = '01/01/1900')"

    Set rsTemp = objO.FetchDataHO(sSQL)
    Call doCreateJob
    
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            With rsGrid
                .AddNew
                !nama_pemesan = NulSID(rsTemp!nama_pemesan)
                !Alamat = NulSID(rsTemp!Alamat)
                !noHP1 = NulSID(rsTemp!hp1)
                !noHP2 = NulSID(rsTemp!hp2)
                !noRangka = NulSID(rsTemp!noRangka)
                !namakategory = NulSID(rsTemp!namakategory)
                !warna = NulSID(rsTemp!namaind)
                !noMesin = NulSID(rsTemp!noMesin)
                !kodetahun = NulSID(rsTemp!kodetahun)
                !kode_leasing = NulSID(rsTemp!kode_leasing)
                !kode_asuransi = NulSID(rsTemp!kode_asuransi)
                !kode = NulSID(rsTemp!kode)
                !NoPolisi = NulSID(rsTemp!NoPolisi)
                !DealerPenjual = NulSID(rsTemp!kode)
                !kodeKATEGori = NulSID(rsTemp!kodeKATEGori)
                !pddfinal = NulSID(rsTemp!pddfinal)
                !kodeSales = NulSID(rsTemp!kodeSales)
                !namasales = NulSID(rsTemp!namasales)
                !NoSPK = NulSID(rsTemp!SPK_No)
                !NAMATT = NulSID(rsTemp!NAMATT)
                !alamattt = NulSID(rsTemp!alamattt)
                !kotatt = NulSID(rsTemp!kotatt)
                !TeleponTT = Left(NulSID(rsTemp!telepon), 20)
                !HP1TT = NulSID(rsTemp!HP1TT)
                !HP2TT = NulSID(rsTemp!HP1TT)
                !MASABERLAKUAWAL = NulSID(rsTemp("MasaBerlakuAwal"))
                !MASABERLAKUAkhir = NulSID(rsTemp("MasaBerlakuAkhir"))
                !KodeReQ = NulSID(rsTemp!KodeReQ)
                !terlibat = False
                .Update
            End With
            rsTemp.MoveNext
        Loop
    End If
    
    Call HeaderColom
End Sub


Private Sub refreshDataUpgrade()
    Dim kriteriaRitel       As String
    Dim kriteriaReminder    As String
    Dim kriteriaPolis       As String
    
    If NulSID(tglpolish1) = "" Or NulSID(tglpolish2) = "" Then
        kriteriaPolis = " and Akhir between '" & Format(isnow, "mm/dd/yyyy") & "' and '" & Format(DateAdd("D", 30, Format(isnow, "dd/mm/yyyy")), "mm/dd/yyyy") & "'"
    Else
        kriteriaPolis = " and dateadd(d,-30,Akhir) between '" & Format(tglpolish1, "mm/dd/yyyy") & "' and '" & Format(tglpolish2, "mm/dd/yyyy") & "'"
    End If
    
'    If NulSID(tglpolish1) = "" Or NulSID(tglpolish2) = "" Then
'        MsgBox "Tanggal Polish harus diisi.", vbInformation
'        Exit Sub
'    End If
    If Tgl1 <> "" And Tgl2 <> "" Then
        kriteriaRitel = " and a.tanggal between '" & Format(Tgl1, "mm/dd/yyyy") & "' and '" & Format(Tgl2, "mm/dd/yyyy") & " 23:59:59'"
    End If
     
    sSQL = "select distinct"
    sSQL = sSQL & " n.NamaCust1+' '+n.namacust2+' '+n.NamaCust3 NamaPembeli, n.Alamat+' '+n.Kelurahan+' '+n.Kecamatan+' '+n.Kota+' '+n.Propinsi AlamatPembeli, n.HP1,n.hp2,"
    sSQL = sSQL & " n1.NamaCust1+' '+n1.namacust2+' '+n1.NamaCust3 NamaPemilik, n1.Alamat+' '+n1.Kelurahan+' '+n1.Kecamatan+' '+n1.Kota+' '+n1.Propinsi AlamatPemilik, n1.HP1 as hp1pemilik, n1.hp2 hp2pemilik,"
    sSQL = sSQL & " a.norangka NoKa, u.nopolisi, d.[desc], u.color, u.noengine, U.Thmobil, l.namaLeasing, m.NamaAsuransi, asr.KodeAsuransi, ASR.masaBerlakuAwal, Ins.masaberlakuAkhir,"
    sSQL = sSQL & " case ASR.Periode when 'Tahun 1' then ASR.jenisAsuransi else '' end TipeAsr1, case ASR.Periode when 'Tahun 2' then ASR.jenisAsuransi else '' end TipeAsr2,"
    sSQL = sSQL & " case ASR.Periode when 'Tahun 3' then ASR.jenisAsuransi else '' end TipeAsr3, case ASR.Periode when 'Tahun 4' then ASR.jenisAsuransi else '' end TipeAsr4,"
    sSQL = sSQL & " case ASR.Periode when 'Tahun 5' then ASR.jenisAsuransi else '' end TipeAsr5, case ASR.Periode when 'Tahun 6' then ASR.jenisAsuransi else '' end TipeAsr6,"
    sSQL = sSQL & " case ASR.Periode when 'Tahun 7' then ASR.jenisAsuransi else '' end TipeAsr7, case ASR.Periode when 'Tahun 8' then ASR.jenisAsuransi else '' end TipeAsr8,"
    sSQL = sSQL & " t.FS_NM_DEALER,isnull(fu.kodereq,'') as kodereq"
    sSQL = sSQL & " from SPK99..Summary a"
    sSQL = sSQL & " left join spk99..Leasing L on a.SPK_No = l.NoSPK and a.NoRangka = l.noRangka"
    sSQL = sSQL & " left join ("
    sSQL = sSQL & " select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, masaberlakuawal, masaberlakuakhir, tenorAsuransi, jenisasuransi,jenisAsuransi2, jenisAsuransi3, jenisAsuransi4, jenisAsuransi5, jenisAsuransi6, jenisasuransi7, jenisasuransi8"
    sSQL = sSQL & "     from spk99..Asuransi"
    sSQL = sSQL & "     Union"
    sSQL = sSQL & "     select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, noPolis, masaBerlakuAwal, masaBerlakuAkhir, tenorAsuransi, jenisasuransi,jenisAsuransi2, jenisAsuransi3, jenisAsuransi4, jenisAsuransi5, jenisAsuransi6, jenisasuransi7, '' jenisasuransi8"
    sSQL = sSQL & "     from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "     Union"
    sSQL = sSQL & "     select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, masaBerlakuAwal, masaBerlakuAkhir, tenorAsuransi, case jenisasuransi when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi2 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi3 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi4 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi5 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi6 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi7 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, '' jenisasuransi8"
    sSQL = sSQL & "     from SPK99..AsuransiUpload"
    sSQL = sSQL & " ) Ins on a.NoRangka = Ins.noRangka"
    sSQL = sSQL & " left join ("
    sSQL = sSQL & "         select * from ("
    sSQL = sSQL & "         select * from ("
    sSQL = sSQL & "         select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, 'Tahun 1' Periode, masaberlakuawal, case when tenorasuransi >0 then DATEADD(year,1,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisasuransi"
    sSQL = sSQL & "         from spk99..Asuransi"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, 'Tahun 2' Periode, masaberlakuawal, case when tenorasuransi >1 then DATEADD(year,2,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisasuransi2"
    sSQL = sSQL & "         from spk99..Asuransi"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, 'Tahun 3' Periode, masaberlakuawal, case when tenorasuransi >2 then DATEADD(year,3,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisasuransi3"
    sSQL = sSQL & "         from spk99..Asuransi"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, 'Tahun 4' Periode, masaberlakuawal, case when tenorasuransi >3 then DATEADD(year,4,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisasuransi4"
    sSQL = sSQL & "         from spk99..Asuransi"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, 'Tahun 5' Periode, masaberlakuawal, case when tenorasuransi >4 then DATEADD(year,5,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisasuransi5"
    sSQL = sSQL & "         from spk99..Asuransi"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, 'Tahun 6' Periode, masaberlakuawal, case when tenorasuransi >5 then DATEADD(year,6,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisasuransi6"
    sSQL = sSQL & "         from spk99..Asuransi"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, 'Tahun 7' Periode, masaberlakuawal, case when tenorasuransi >6 then DATEADD(year,7,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisasuransi7"
    sSQL = sSQL & "         from spk99..Asuransi"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, 'Tahun 8' Periode, masaberlakuawal, case when tenorasuransi >7 then DATEADD(year,8,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisasuransi8"
    sSQL = sSQL & "         from spk99..Asuransi"
    sSQL = sSQL & "         ) S where s.Akhir <> '01/01/1900'"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select * from ("
    sSQL = sSQL & "         select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, nopolis, 'Tahun 1' Periode, masaBerlakuAwal,case when tenorasuransi >0 then DATEADD(year,1,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi"
    sSQL = sSQL & "         from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi,nopolis, 'Tahun 2' Periode, masaBerlakuAwal,case when tenorasuransi >1 then DATEADD(year,2,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi2"
    sSQL = sSQL & "         from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, nopolis, 'Tahun 3' Periode, masaBerlakuAwal,case when tenorasuransi >2 then DATEADD(year,3,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi3"
    sSQL = sSQL & "         from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, nopolis, 'Tahun 4' Periode, masaBerlakuAwal,case when tenorasuransi >3 then DATEADD(year,4,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi4"
    sSQL = sSQL & "         from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, nopolis, 'Tahun 5' Periode, masaBerlakuAwal,case when tenorasuransi >4 then DATEADD(year,5,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi5"
    sSQL = sSQL & "         from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, nopolis, 'Tahun 6' Periode, masaBerlakuAwal,case when tenorasuransi >5 then DATEADD(year,6,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi6"
    sSQL = sSQL & "         from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, nopolis, 'Tahun 7' Periode, masaBerlakuAwal,case when tenorasuransi >6 then DATEADD(year,7,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi7"
    sSQL = sSQL & "         from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, nopolis, 'Tahun 8' Periode, masaBerlakuAwal,case when tenorasuransi >7 then DATEADD(year,8,masaberlakuawal) else '' end Akhir, tenorAsuransi, '' jenisAsuransi8"
    sSQL = sSQL & "         from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "         ) O"
    sSQL = sSQL & "         where o.Akhir <> '01/01/1900'"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select * from ("
    sSQL = sSQL & "         select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, 'Tahun 1' Periode, masaBerlakuAwal, case when tenorasuransi >0 then DATEADD(year,1,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi"
    sSQL = sSQL & "         from SPK99..AsuransiUpload"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, 'Tahun 2' Periode, masaBerlakuAwal, case when tenorasuransi >1 then DATEADD(year,2,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi2"
    sSQL = sSQL & "         from SPK99..AsuransiUpload"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, 'Tahun 3' Periode, masaBerlakuAwal, case when tenorasuransi >2 then DATEADD(year,3,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi3"
    sSQL = sSQL & "         from SPK99..AsuransiUpload"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, 'Tahun 4' Periode, masaBerlakuAwal, case when tenorasuransi >3 then DATEADD(year,4,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi4"
    sSQL = sSQL & "         from SPK99..AsuransiUpload"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, 'Tahun 5' Periode, masaBerlakuAwal, case when tenorasuransi >4 then DATEADD(year,5,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi5"
    sSQL = sSQL & "         from SPK99..AsuransiUpload"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, 'Tahun 6' Periode, masaBerlakuAwal, case when tenorasuransi >5 then DATEADD(year,6,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi6"
    sSQL = sSQL & "         from SPK99..AsuransiUpload"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, 'Tahun 7' Periode, masaBerlakuAwal, case when tenorasuransi >6 then DATEADD(year,7,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi7"
    sSQL = sSQL & "         from SPK99..AsuransiUpload"
    sSQL = sSQL & "         Union"
    sSQL = sSQL & "         select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, 'Tahun 8' Periode, masaBerlakuAwal, case when tenorasuransi >7 then DATEADD(year,8,masaberlakuawal) else '' end Akhir, tenorAsuransi, jenisAsuransi8"
    sSQL = sSQL & "         from SPK99..AsuransiUpload"
    sSQL = sSQL & "         ) U"
    sSQL = sSQL & "         where u.Akhir <> '01/01/1900'"
    sSQL = sSQL & "         ) P"
    sSQL = sSQL & " ) ASR on a.NoRangka = ASR.noRangka"
    sSQL = sSQL & " left join oss..MstAsuransi M on asr.KodeAsuransi = m.KodeAsuransi"
    sSQL = sSQL & " left join SingleIDOri..UnitCustomer U on a.NoRangka = U.NoRangka"
    sSQL = sSQL & " left join SingleIDOri..NewCustomer N on U.IDCustSales = n.IDCust"
    sSQL = sSQL & " left join SingleIDOri..NewCustomer N1 on U.IDCustService = n1.IDCust"
    sSQL = sSQL & " left join bengkel99..model D on u.mdl = D.mdl"
    sSQL = sSQL & " left join bengkel99..t_dealer T on a.KodeDealer = t.KODECABANG and INTERNAL = 1"
    sSQL = sSQL & " left join [" & Log.LogServ & "].bengkel" & Right(gDb, 2) & ".dbo.FUAsuransi fu on a.norangka = fu.norangka and fu.aktifitas = 'UPGRADE'"
    sSQL = sSQL & " where isnull(fu.kodereq,'') <> '' and TipeTr = 'SPK' and Status = 1 and a.KodeDealer = '" & Log.kodeCabang & "'"
    ''sSQL = sSQL & " and dateadd(d,-30,Akhir) between '" & Format(tglpolish1, "mm/dd/yyyy") & "' and '" & Format(tglpolish2, "mm/dd/yyyy") & "'"
    sSQL = sSQL & kriteriaPolis
    sSQL = sSQL & " order by a.NoRangka"

    Set rsTemp = objO.FetchDataHO(sSQL)
    Call doCreateJob
    
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            With rsGrid
                .AddNew
                !nama_pemesan = NulSID(rsTemp!namapembeli)
                !Alamat = NulSID(rsTemp!Alamatpembeli)
                !noHP1 = NulSID(rsTemp!hp1)
                !noHP2 = NulSID(rsTemp!hp2)
                !noRangka = NulSID(rsTemp!noKa)
                !namakategory = NulSID(rsTemp!desc)
                !warna = NulSID(rsTemp!Color)
                !noMesin = NulSID(rsTemp!NoEngine)
                !kodetahun = NulSID(rsTemp!thMobil)
                !kode_leasing = "" 'NulSID(rsTemp!kode_leasing)
                !kode_asuransi = NulSID(rsTemp!KodeAsuransi)
                !kode = Log.kodeCabang ' NulSID(rsTemp!kode)
                !NoPolisi = NulSID(rsTemp!NoPolisi)
                !DealerPenjual = Log.kodeCabang 'NulSID(rsTemp!kode)
                !kodeKATEGori = "" 'NulSID(rsTemp!kodeKATEGori)
                !pddfinal = "01/01/1900" 'NulSID(rsTemp!pddfinal)
                !kodeSales = "" 'NulSID(rsTemp!kodeSales)
                !namasales = "" 'NulSID(rsTemp!namasales)
                !NoSPK = "" 'NulSID(rsTemp!spk_no)
                !NAMATT = NulSID(rsTemp!namapemilik)
                !alamattt = NulSID(rsTemp!alamatpemilik)
                !kotatt = "" 'NulSID(rsTemp!kotatt)
                !TeleponTT = "" 'Left(NulSID(rsTemp!telepon), 20)
                !HP1TT = NulSID(rsTemp!hp1pemilik)
                !HP2TT = NulSID(rsTemp!hp2pemilik)
                !MASABERLAKUAWAL = NulSID(rsTemp("MasaBerlakuAwal"))
                !MASABERLAKUAkhir = NulSID(rsTemp("MasaBerlakuAkhir"))
                !KodeReQ = NulSID(rsTemp!KodeReQ)
                !terlibat = False
                .Update
            End With
            rsTemp.MoveNext
        Loop
    End If
    
    Call HeaderColom
End Sub

Private Sub refreshCombo()
    AdoAsuransi.ConnectionString = SQLConHO
    AdoAsuransi.RecordSource = "select KodeAsuransi,NamaAsuransi,WilayahAsuransi from OSS..mstAsuransi where oss = 1"
    cbAsuransi.DataSourceList = AdoAsuransi
    cbAsuransi.DataFieldList = "NamaAsuransi"
    cbAsuransi.Columns(1).Width = 2400
    cbAsuransi.Columns(2).Width = 2400
    cbAsuransi.Columns(1).Caption = "Nama Asuransi"
    cbAsuransi.Columns(2).Caption = "Wilayah Asuransi"
    cbAsuransi.Columns(0).Visible = False
    
    AdoRetail.ConnectionString = SQLConHO
    AdoRetail.RecordSource = "select KodeDealer,NamaDealer from ProfilSPK..DealerBD where  Sales <> ''"
    cbRetail.DataSourceList = AdoRetail
    cbRetail.DataFieldList = "KodeDealer"
    cbRetail.Columns(1).Width = 2400
    cbRetail.Columns(1).Caption = "Nama Dealer"
    cbRetail.Columns(0).Visible = False
End Sub

Private Sub createHeadAsuransi()
    
   On Error GoTo createHeadAsuransi_Error
    With gridAsuransi
        .Columns(0).Width = 3000
        .Columns(0).Caption = "Nama Pemesan"
        .Columns(0).DataField = "nama_pemesan"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 3000
        .Columns(1).Caption = "Alamat Pemesan"
        .Columns(1).DataField = "Alamat"
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 1000
        .Columns(2).Caption = "No Telp"
        .Columns(2).DataField = "noHP"
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 1500
        .Columns(3).Caption = "No Rangka"
        .Columns(3).DataField = "noRangka"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 1500
        .Columns(4).Caption = "Nama Kategory"
        .Columns(4).DataField = "namakategory"
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 1000
        .Columns(5).Caption = "Warna"
        .Columns(5).DataField = "warna"
        .Splits(0).Columns(5).Locked = True
        
    End With
    

   On Error GoTo 0
   Exit Sub

createHeadAsuransi_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure createHeadAsuransi of Form frmRetentionAsuransiHasil"


End Sub

Private Sub doCreateJob()
   On Error GoTo doCreateJob_Error

    If rsGrid.State Then rsGrid.Close
    Set rsGrid = Nothing
    With rsGrid
        .Fields.Append "terlibat", adBoolean
        .Fields.Append "nama_pemesan", adVarChar, 50
        .Fields.Append "Alamat", adVarChar, 250
        .Fields.Append "noHP1", adVarChar, 50
        .Fields.Append "noHP2", adVarChar, 50
        .Fields.Append "noRangka", adVarChar, 30
        .Fields.Append "namakategory", adVarChar, 50
        .Fields.Append "warna", adVarChar, 50
        .Fields.Append "nomesin", adVarChar, 50
        .Fields.Append "kodetahun", adVarChar, 50
        .Fields.Append "kode_leasing", adVarChar, 50
        .Fields.Append "kode_asuransi", adVarChar, 50
        .Fields.Append "kode", adVarChar, 50            ' TANDA JIKA DIBEBANKAN KE CUSTOMER
        .Fields.Append "NoPolisi", adVarChar, 15
        .Fields.Append "DealerPenjual", adVarChar, 30     '11
        .Fields.Append "Kodekategori", adVarChar, 50      '12
        .Fields.Append "PddFinal", adVarChar, 50          '13
        .Fields.Append "KodeSales", adVarChar, 10         '14
        .Fields.Append "Namasales", adVarChar, 100        '15
        .Fields.Append "NoSPK", adVarChar, 15             '16
        .Fields.Append "HP1", adVarChar, 20               '17
        .Fields.Append "HP2", adVarChar, 20               '18
        .Fields.Append "NamaTT", adVarChar, 200           '19
        .Fields.Append "AlamatTT", adVarChar, 320         '20
        .Fields.Append "KotaTT", adVarChar, 100           '20
        .Fields.Append "TeleponTT", adVarChar, 20         '21
        .Fields.Append "HP1TT", adVarChar, 20             '22
        .Fields.Append "HP2TT", adVarChar, 20             '23
        .Fields.Append "masaberlakuawal", adVarChar, 20   '24
        .Fields.Append "masaberlakuakhir", adVarChar, 20  '25
        .Fields.Append "nopolish", adVarChar, 50
        .Fields.Append "kodereq", adVarChar, 50
        .Open
    End With
    Set gridAsuransi.DataSource = rsGrid
   On Error GoTo 0
   Exit Sub

doCreateJob_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure doCreateJob of Form frmRetentionAsuransiHasil"
    
End Sub


Private Sub HeaderColom()
    With gridAsuransi
        .Columns(0).Caption = "Terlibat"
        .Columns(0).Width = 1000
        .Columns(0).Alignment = dbgCenter
        .Columns(0).DataField = "terlibat"
        .Splits(0).Columns(0).Locked = False
        .Splits(0).Columns(0).ValueItems.Presentation = dbgCheckBox
                   
        .Columns(1).Caption = "Nama Customer"
        .Columns(1).DataField = "nama_pemesan"
        .Columns(1).Width = 2000
        .Columns(1).Alignment = dbgLeft
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Caption = "Alamat"
        .Columns(2).Width = 4000
        .Columns(2).DataField = "Alamat"
        .Columns(2).Alignment = dbgLeft
        .Splits(0).Columns(2).Locked = True
          
        .Columns(3).Caption = "No HP"
        .Columns(3).Width = 1000
        .Columns(3).DataField = "noHP1"
'        .Columns(3).Alignment = dbgRight
'        .Columns(3).NumberFormat = "#,#0"
        .Columns(3).Locked = False
         
        .Columns(4).Caption = "NIK"
        .Columns(4).Width = 500
        .Columns(4).DataField = "noHP2"
        .Columns(4).Visible = False
                   
        .Columns(5).Caption = "No Rangka"
        .Columns(5).Width = 2000
        .Columns(5).DataField = "noRangka"
        .Columns(5).Visible = True
                
        .Columns(6).Caption = "Model Kendaraan"
        .Columns(6).Width = 2000
        .Columns(6).DataField = "namakategory"
        .Columns(6).Visible = True
                
        .Columns(7).Caption = "Warna"
        .Columns(7).Width = 1500
        .Columns(7).DataField = "warna"
        .Columns(7).Visible = True
                
        .Columns(8).Caption = "Kode Dealer"
        .Columns(8).Width = 2000
        .Columns(8).DataField = "kode"
        .Columns(8).Visible = True
        
        
        .Columns(9).DataField = "nomesin"
        .Columns(9).Visible = False
        
        .Columns(10).DataField = "kodetahun"
        .Columns(10).Visible = False
        
        .Columns(11).DataField = "kode_leasing"
        .Columns(11).Visible = False
        
        .Columns(12).DataField = "kode_asuransi"
        .Columns(12).Visible = False
                
    End With
End Sub

 Private Sub refreshDataRenew()
    Dim kriteriaRitel       As String
    Dim kriteriaReminder    As String
    If tglpolish1 = "" Or tglpolish2 = "" Then
        MsgBox "Tanggal Polish harus diisi.", vbInformation
        Exit Sub
    End If
    If Tgl1 <> "" And Tgl2 <> "" Then
        kriteriaRitel = " and a.tanggal between '" & Format(Tgl1, "mm/dd/yyyy") & "' and '" & Format(Tgl2, "mm/dd/yyyy") & " 23:59:59'"
    End If
     
    sSQL = "select distinct "
    sSQL = sSQL & " n.NamaCust1+' '+n.namacust2+' '+n.NamaCust3 NamaPembeli, n.Alamat+' '+n.Kelurahan+' '+n.Kecamatan+' '+n.Kota+' '+n.Propinsi AlamatPembeli, n.HP1,n.hp2,"
    sSQL = sSQL & " n1.NamaCust1+' '+n1.namacust2+' '+n1.NamaCust3 NamaPemilik, n1.Alamat+' '+n1.Kelurahan+' '+n1.Kecamatan+' '+n1.Kota+' '+n1.Propinsi AlamatPemilik, n1.HP1 as hp1pemilik,n1.hp2 hp2Pemilik,"
    sSQL = sSQL & " a.norangka NoKa, u.nopolisi, d.[desc], u.color, u.noengine, U.Thmobil, l.namaLeasing, m.NamaAsuransi, asr.KodeAsuransi, ASR.masaBerlakuAwal, ASR.masaBerlakuAkhir,"
    sSQL = sSQL & " case ASR.jenisAsuransi when 2 then 'All Risk' when 1 then 'TLO' else '' end TipeAsr1 , case ASR.jenisAsuransi2 when 2 then 'All Risk' when 1 then 'TLO' else '' end TipeAsr2,"
    sSQL = sSQL & " case ASR.jenisAsuransi3 when 2 then 'All Risk' when 1 then 'TLO' else '' end TipeAsr3, case ASR.jenisAsuransi4 when 2 then 'All Risk' when 1 then 'TLO' else '' end TipeAsr4,"
    sSQL = sSQL & " case ASR.jenisAsuransi5 when 2 then 'All Risk' when 1 then 'TLO' else '' end TipeAsr5, case ASR.jenisAsuransi6 when 2 then 'All Risk' when 1 then 'TLO' else '' end TipeAsr6,"
    sSQL = sSQL & " case ASR.jenisAsuransi7 when 2 then 'All Risk' when 1 then 'TLO' else '' end TipeAsr7, case ASR.jenisAsuransi8 when 2 then 'All Risk' when 1 then 'TLO' else '' end TipeAsr8,"
    sSQL = sSQL & " t.FS_NM_DEALER, isnull(fu.kodereq,'') as kodereq"
    sSQL = sSQL & " from SPK99..Summary a"
    sSQL = sSQL & " left join [" & Log.LogServ & "].bengkel" & Right(gDb, 2) & ".dbo.FUAsuransi b on a.norangka = b.norangka and b.aktifitas = 'NEW'"
    sSQL = sSQL & " left join spk99..Leasing L on a.SPK_No = l.NoSPK and a.NoRangka = l.noRangka"
    sSQL = sSQL & " left join ("
    sSQL = sSQL & "     select 'SPK' sumber, norangka, nospk, kodeasuransi, nopolis, masaberlakuawal, masaberlakuakhir, tenorAsuransi, jenisasuransi,jenisAsuransi2, jenisAsuransi3, jenisAsuransi4, jenisAsuransi5, jenisAsuransi6, jenisasuransi7, jenisasuransi8"
    sSQL = sSQL & "     from spk99..Asuransi"
    sSQL = sSQL & "     Union"
    sSQL = sSQL & "     select 'OSS' sumber, noRangka, NoSPK, KodeAsuransi, noPolis, masaBerlakuAwal, masaBerlakuAkhir, tenorAsuransi, jenisasuransi,jenisAsuransi2, jenisAsuransi3, jenisAsuransi4, jenisAsuransi5, jenisAsuransi6, jenisasuransi7, '' jenisasuransi8"
    sSQL = sSQL & "     from oss..Asuransi where statusAsuransi in ('DISETUJUI')"
    sSQL = sSQL & "     Union"
    sSQL = sSQL & "     select 'Upload' sumber, noRangka, '' NoSPK, KodeAsuransi, noPolis, masaBerlakuAwal, masaBerlakuAkhir, tenorAsuransi, case jenisasuransi when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi2 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi3 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi4 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi5 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi6 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, case jenisasuransi7 when 'Allrisk' then 2 when 'TLO' then 1 else 0 end, '' jenisasuransi8"
    sSQL = sSQL & "     from SPK99..AsuransiUpload"
    sSQL = sSQL & " ) ASR on a.NoRangka = ASR.noRangka"
    sSQL = sSQL & " left join oss..MstAsuransi M on asr.KodeAsuransi = m.KodeAsuransi"
    sSQL = sSQL & " left join SingleIDOri..UnitCustomer U on a.NoRangka = U.NoRangka"
    sSQL = sSQL & " left join SingleIDOri..NewCustomer N on U.IDCustSales = n.IDCust"
    sSQL = sSQL & " left join SingleIDOri..NewCustomer N1 on U.IDCustService = n1.IDCust"
    sSQL = sSQL & " left join bengkel99..model D on u.mdl = D.mdl"
    sSQL = sSQL & " left join bengkel99..t_dealer T on a.KodeDealer = t.KODECABANG and INTERNAL = 1"
    sSQL = sSQL & " left join [" & Log.LogServ & "].bengkel" & Right(gDb, 2) & ".dbo.FUAsuransi fu on a.norangka = fu.norangka and fu.aktifitas = 'RENEW'"
    sSQL = sSQL & " where isnull(fu.kodereq,'') <> '' and TipeTr = 'SPK' and Status = 1 and a.KodeDealer = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " and dateadd(d,-30,asr.masaBerlakuAkhir) between '" & Format(tglpolish1, "mm/dd/yyyy") & "' and '" & Format(tglpolish2, "mm/dd/yyyy") & "'"
    sSQL = sSQL & " order by a.NoRangka"

    Set rsTemp = objO.FetchDataHO(sSQL)
    Call doCreateJob
    
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            With rsGrid
                .AddNew
                !nama_pemesan = NulSID(rsTemp!namapembeli)
                !Alamat = NulSID(rsTemp!Alamatpembeli)
                !noHP1 = NulSID(rsTemp!hp1)
                !noHP2 = NulSID(rsTemp!hp2)
                !noRangka = NulSID(rsTemp!noKa)
                !namakategory = NulSID(rsTemp!desc)
                !warna = NulSID(rsTemp!Color)
                !noMesin = NulSID(rsTemp!NoEngine)
                !kodetahun = NulSID(rsTemp!thMobil)
                !kode_leasing = "" 'NulSID(rsTemp!kode_leasing)
                !kode_asuransi = NulSID(rsTemp!KodeAsuransi)
                !kode = Log.kodeCabang ' NulSID(rsTemp!kode)
                !NoPolisi = NulSID(rsTemp!NoPolisi)
                !DealerPenjual = Log.kodeCabang 'NulSID(rsTemp!kode)
                !kodeKATEGori = "" 'NulSID(rsTemp!kodeKATEGori)
                !pddfinal = "01/01/1900" 'NulSID(rsTemp!pddfinal)
                !kodeSales = "" 'NulSID(rsTemp!kodeSales)
                !namasales = "" 'NulSID(rsTemp!namasales)
                !NoSPK = "" 'NulSID(rsTemp!spk_no)
                !NAMATT = NulSID(rsTemp!namapemilik)
                !alamattt = NulSID(rsTemp!alamatpemilik)
                !kotatt = "" 'NulSID(rsTemp!kotatt)
                !TeleponTT = "" 'Left(NulSID(rsTemp!telepon), 20)
                !HP1TT = NulSID(rsTemp!hp1pemilik)
                !HP2TT = NulSID(rsTemp!hp2pemilik)
                !MASABERLAKUAWAL = NulSID(rsTemp("MasaBerlakuAwal"))
                !MASABERLAKUAkhir = NulSID(rsTemp("MasaBerlakuAkhir"))
                !KodeReQ = NulSID(rsTemp!KodeReQ)
                !terlibat = False
                .Update
            End With
            rsTemp.MoveNext
        Loop
    End If
    
    Call HeaderColom
End Sub

Private Sub GridAsuransi_FilterChange()
On Error GoTo errHandler
    Set cols = gridAsuransi.Columns
    Dim C As Integer
    
    C = gridAsuransi.col
    gridAsuransi.HoldFields
    rsGrid.filter = getFilter()
    If rsGrid.EOF Then
        For Each col In gridAsuransi.Columns
            col.FilterText = ""
        Next col
        rsGrid.filter = ""
    End If
    gridAsuransi.col = C
    gridAsuransi.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub


Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        
        End If
    Next col
    getFilter = tmp
End Function

Private Sub gridAsuransi_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex > 0 And ColIndex <= 12 Then rsGrid.Sort = gridAsuransi.Columns(ColIndex).DataField & " " & Urut
4        gridAsuransi.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")

End Sub
