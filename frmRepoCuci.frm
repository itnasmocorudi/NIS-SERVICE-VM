VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRepoCuci 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Harian Mekanik"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H000000FF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5160
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1395
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5160
      Width           =   1440
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCetak 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8325
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5205
      TabIndex        =   19
      Top             =   5190
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   4095
      TabIndex        =   18
      Top             =   5175
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      BTYPE           =   8
      TX              =   "Printer :"
      ENAB            =   0   'False
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
      MICON           =   "frmRepoCuci.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   7575
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   7455
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   6000
         Top             =   360
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbStallCuci 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1080
         Width           =   3135
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   5530
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "Pilih Semua Stall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   450
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   794
         _Version        =   327682
         Appearance      =   1
      End
      Begin TDBDate6Ctl.TDBDate Tgl1 
         Height          =   330
         Left            =   1680
         TabIndex        =   14
         Top             =   120
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoCuci.frx":001C
         Caption         =   "frmRepoCuci.frx":0148
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoCuci.frx":01B4
         Keys            =   "frmRepoCuci.frx":01D2
         Spin            =   "frmRepoCuci.frx":0230
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
         Left            =   3360
         TabIndex        =   15
         Top             =   120
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoCuci.frx":0258
         Caption         =   "frmRepoCuci.frx":0384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoCuci.frx":03F0
         Keys            =   "frmRepoCuci.frx":040E
         Spin            =   "frmRepoCuci.frx":046C
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
      Begin TDBDate6Ctl.TDBDate TglCuci1 
         Height          =   330
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoCuci.frx":0494
         Caption         =   "frmRepoCuci.frx":05C0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoCuci.frx":062C
         Keys            =   "frmRepoCuci.frx":064A
         Spin            =   "frmRepoCuci.frx":06A8
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
      Begin TDBDate6Ctl.TDBDate TglCuci2 
         Height          =   330
         Left            =   3360
         TabIndex        =   17
         Top             =   600
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoCuci.frx":06D0
         Caption         =   "frmRepoCuci.frx":07FC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoCuci.frx":0868
         Keys            =   "frmRepoCuci.frx":0886
         Spin            =   "frmRepoCuci.frx":08E4
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Stall"
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
         Left            =   645
         TabIndex        =   10
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label2 
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
         Left            =   3240
         TabIndex        =   9
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Cuci"
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
         Left            =   525
         TabIndex        =   8
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal WO"
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
         Left            =   555
         TabIndex        =   3
         Top             =   240
         Width           =   885
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
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   0
         X2              =   7440
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   661
   End
   Begin ucXPButton.XPButton tblSimpan 
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   3960
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&p"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmRepoCuci.frx":090C
      PICN            =   "frmRepoCuci.frx":0928
      PICH            =   "frmRepoCuci.frx":0D7F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton tblKeluar 
      Height          =   495
      Left            =   6720
      TabIndex        =   13
      Top             =   3960
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
      MICON           =   "frmRepoCuci.frx":1269
      PICN            =   "frmRepoCuci.frx":1285
      PICH            =   "frmRepoCuci.frx":1759
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "LAPORAN UNIT CUCI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   7455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   480
      X2              =   7920
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label LblKnow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE WAIT  (^_^) ..."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label prosen 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "     %"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3600
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmRepoCuci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          : REPORT CUCI | RUDI


Dim rsTemp As ADODB.Recordset
Dim Report As Object 'New rptRepoCuci
'Dim ReportPengering As New rptRepoCuciPengering

Dim doubleJob As Boolean
Dim resetJob As Boolean
Dim resetRate As Boolean
Dim vNamaTabel As String
Dim objRepo As New clsRepo
Public isPengering As Boolean
Public strSql As String

Dim totalWaktuTC As Double
Dim totalWaktuPC As Double
Dim totalWaktuTotal As Double
Dim totalWaktuTK As Double
Dim totalWaktuPK As Double

Dim avgWaktuTC As Single
Dim avgWaktuPC As Single
Dim avgWaktuTotal As Single
Dim avgWaktuTK As Double
Dim avgWaktuPK As Double

Dim kriteria As String


Private Sub hapusForm()
    Tgl1 = isnow
    Tgl2 = isnow
    TglCuci1 = isnow
    TglCuci2 = isnow
    LblKnow.Visible = False
    ProgressBar1.Visible = False
    prosen.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub isiCombo()
    Adodc1.ConnectionString = SQLCONBengkel
    Adodc1.RecordSource = "SELECT NoStall,NamaStall FROM qty_stallwashing WHERE tipestall = 'Cuci' AND AKTIF = '1' AND ISDEL = '0'"
    cbStallCuci.DataSourceList = Adodc1
    cbStallCuci.DataFieldList = "NoStall"
    cbStallCuci.DataFieldToDisplay = "NamaStall"
    cbStallCuci.Columns(0).Width = 1200
    cbStallCuci.Columns(1).Width = 4000
    cbStallCuci.Columns(0).Caption = "No Stall"
    cbStallCuci.Columns(1).Caption = "Nama Stall"
End Sub

Private Function createTableCuci(vNamaTable As String) As Boolean
Dim rsHlp As ADODB.Recordset
    createTableCuci = False
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)

    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ( " _
         & "[NOWO][VARCHAR](20) NOT NULL DEFAULT '', " _
         & "[NOPOLISI][VARCHAR](20) NOT NULL DEFAULT '', " _
         & "[SELESAIQC][VARCHAR](30) NOT NULL DEFAULT '01/01/1900 00:00', " _
         & "[MULAICUCI][VARCHAR](30) NOT NULL DEFAULT '01/01/1900 00:00', " _
         & "[SELESAICUCI][VARCHAR](30) NOT NULL DEFAULT '01/01/1900 00:00', " _
         & "[LTTUNGGUCUCI][VARCHAR](5) NOT NULL DEFAULT '0', " _
         & "[LTPROSESCUCI][VARCHAR](5) NOT NULL DEFAULT '0', " _
         & "[LTTOTAL][VARCHAR](5) NOT NULL DEFAULT '0'"
    Call objO.UpdateData(sSQL)
    
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTableCuci = True
    End If
End Function


Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset
    createTemp = False
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)

    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ( " _
         & "[NO][DECIMAL](19) NOT NULL DEFAULT '', " _
         & "[NOWO][VARCHAR](10) NOT NULL DEFAULT '', " _
         & "[TANGGAL][DATETIME] NOT NULL DEFAULT '01/01/3000', " _
         & "[NOSTALL][CHAR](10) NOT NULL DEFAULT '', " _
         & "[START][CHAR](5) NOT NULL DEFAULT '', " _
         & "[STATUS][CHAR](1) NOT NULL DEFAULT '', " _
         & "[SELESAI][CHAR](5) NOT NULL DEFAULT '', " _
         & "[MECH][CHAR](5) NOT NULL DEFAULT '', " _
         & "[MECHNAME][CHAR](30) NOT NULL DEFAULT '', " _
         & "[FOREMAN][CHAR](5) NOT NULL DEFAULT '', " _
         & "[MENIT][DECIMAL](18,2) NOT NULL DEFAULT 0, " _
         & "[NOPOLISI][VARCHAR](15) NOT NULL DEFAULT '', " _
         & "[NORANGKA][VARCHAR](25) NOT NULL DEFAULT '', " _
         & "[RATE1][DECIMAL](18,2) NOT NULL DEFAULT 0, " _
         & "[RATE2][DECIMAL](18,2) NOT NULL DEFAULT 0, " _
         & "[GROUP][CHAR](5) NOT NULL DEFAULT '', " _
         & "[JOB][CHAR](100) NOT NULL DEFAULT '', " _
         & "[REASON][CHAR](500) NOT NULL DEFAULT '') ON [PRIMARY]"
    Call objO.UpdateData(sSQL)
    
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    End If
End Function

Private Sub IsiTempRepo(Ket As String, vNamaTabel As String, Txt As String)
Dim rsDouble As ADODB.Recordset
Dim btkTabel As Boolean
Dim woLama As String
Dim startLama As String
Dim stallLama As String
Dim jamKrj As Double
Dim rate1 As Double
Dim rate2 As Double
Dim nCount As Integer
Dim nProg As Integer
Dim strSql As String
    
    woLama = ""
    startLama = ""
    Dim rteJMech As Double
    
    If Txt <> "" Then
        Txt = " AND " & Txt
    End If
    
    LblKnow.Visible = True
    DoEvents
    Set rsTemp = objRepo.searchWorkTech(Format(Tgl1, "MM/DD/YYYY"), Format(Tgl2, "MM/DD/YYYY"), Ket, Txt)
    If Not rsTemp.EOF Then
        nCount = 0
        nProg = 0
        btkTabel = createTemp(vNamaTabel)
        Do While Not rsTemp.EOF
            LblKnow.Visible = False
            ProgressBar1.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = rsTemp.RecordCount
            nProg = nProg + 1
            ProgressBar1.value = nProg
            prosen.Caption = str(Round((nProg / rsTemp.RecordCount) * 100, 0)) + " %"
            prosen.Visible = True
            DoEvents
            If btkTabel Then
                If woLama = Trim(rsTemp("NOORDER")) Then
                    rate1 = 0
                    rate2 = 0
                Else
                    rate1 = CDbl(rsTemp("RATE1"))
                    rate2 = CDbl(rsTemp("RATE2"))
                End If
                
                If startLama = Trim(rsTemp("START")) And stallLama = Trim(rsTemp("NOSTALL")) Then
                    jamKrj = 0
                Else
                    jamKrj = CDbl(Round(CDbl(rsTemp("JAMKERJA")) / 60, 2))
                End If
                resetJob = False
                Set rsDouble = objO.FetchData("SELECT NOWO FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NOWO = '" & Trim(rsTemp("NOORDER")) & "' AND JOB = '" & Trim(rsTemp("JOB")) & "' AND MECH = '" & Trim(rsTemp("MECH")) & "' AND TANGGAL = '" & Format(rsTemp("TANGGAL"), "MM/DD/YYYY") & "'")
                If Not rsDouble.EOF Then
                    resetJob = True
                End If
                If rsDouble.State <> adStateClosed Then rsDouble.Close
                
                resetRate = False
                Set rsDouble = objO.FetchData("SELECT TOP 1 MECH, DATEDIFF(n,START,PAUSE) FROM TEMPPROGRESS WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND FLAGP <> 'M' GROUP BY MECH, DATEDIFF(n,START,PAUSE) ORDER BY COUNT(*), DATEDIFF(n,START,PAUSE) DESC")
                If Not rsDouble.EOF Then
                    If Trim(rsDouble("MECH")) <> Trim(rsTemp("MECH")) Then
                        resetRate = True
                    End If
                End If
                If rsDouble.State <> adStateClosed Then rsDouble.Close
                
                doubleJob = False
                'Set rsDouble = objO.FetchData("SELECT NOORDER FROM TEMPPROGRESS WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND TANGGAL < '" & Format(Tgl1, "MM/DD/YYYY") & "' AND FLAGP <> 'M'")
                Set rsDouble = objO.FetchData("SELECT NOORDER FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND TANGGAL < '" & Format(Tgl1, "MM/DD/YYYY") & "' AND FLAGP not in ('M','P')")
                If Not rsDouble.EOF Then
                    doubleJob = True
                End If
                If rsDouble.State <> adStateClosed Then rsDouble.Close
                
                '' ALASAN JOB STOPPED
                Dim rsStop As ADODB.Recordset
                Dim ketStop As String
                ketStop = ""
                If Trim(rsTemp("PAUSE")) <> "" And Trim(rsTemp("FLAGP")) = "P" Then
                    Set rsStop = objO.FetchData("SELECT * FROM PROGRESS WITH(NOLOCK) WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND FLAG = 'P' AND JAM = '" & Trim(rsTemp("PAUSE")) & "'")
                    If Not rsStop.EOF Then
                        ketStop = Trim(Mid(rsStop("KETERANGAN"), InStr(rsStop("KETERANGAN"), "|") + 1, Len(rsStop("KETERANGAN")) - InStr(rsStop("KETERANGAN"), "|")))
                    End If
                    If rsStop.State <> adStateClosed Then rsStop.Close
                End If
                
                If Not resetJob Or jamKrj > 0 Then
                    If mekanik(Trim(rsTemp("MECH"))) Then
                        If jamKrj > 0 Then nCount = nCount + 1
                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (" & CDbl(nCount) & ",'" & Trim(rsTemp("NOORDER")) & "', '" & Format(rsTemp("TANGGAL"), "MM/DD/YYYY") & "', '" & Trim(rsTemp("NOSTALL")) & "', '" & Trim(rsTemp("START")) & "', " _
                                     & "'" & Trim(rsTemp("FLAGP")) & "', '" & Trim(rsTemp("PAUSE")) & "', '" & Trim(rsTemp("MECH")) & "', '" & Trim(rsTemp("NAME")) & "', '" & Trim(rsTemp("FOREMAN")) & "', " & RepRp(Replace(jamKrj, ",", ".")) & ", '" & Trim(rsTemp("NO_POLISI")) & "', " _
                                     & "'" & Trim(rsTemp("NO_RANGKA")) & "', " & RepRp(IIf(resetRate, 0, IIf(doubleJob, 0, IIf(resetJob, 0, Replace(CDbl(rate1), ",", "."))))) & "," & RepRp(IIf(resetRate, 0, IIf(doubleJob, 0, IIf(resetJob, 0, Replace(CDbl(rate2), ",", "."))))) & ", " _
                                     & "'" & IIf(resetJob, "", Trim(rsTemp("GRP"))) & "','" & IIf(resetJob, "", Trim(rsTemp("JOB"))) & "','" & Trim(ketStop) & "')")
                    End If
                End If
                woLama = Trim(rsTemp("NOORDER"))
                startLama = Trim(rsTemp("START"))
                stallLama = Trim(rsTemp("NOSTALL"))
            End If
            rsTemp.MoveNext
        Loop
        Call Cetak(vNamaTabel, Txt)
    Else
        MsgBox "Tidak ada data !!"
        Exit Sub
    End If
End Sub

Private Function mekanik(namaMekanik As String) As Boolean
    Dim rsSearch As ADODB.Recordset
    mekanik = False
    Set rsSearch = objO.FetchData("SELECT MECH FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(namaMekanik) & "'")
    If Not rsSearch.EOF Then
        mekanik = True
    End If
End Function


Private Sub Cetak(vNamaTabel As String, Txt As String)
Dim oRs As Object
Dim Report As New rptRepoHarianTech1
Screen.MousePointer = vbHourglass
    
    Set oRs = objRepo.printRepo(vNamaTabel)
    Report.database.SetDataSource oRs
        
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
   
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
    prosen.Visible = False
    ProgressBar1.Visible = False
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = 200 'asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > 300 Then
        CRViewer1.Width = ScaleWidth - 300 'asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
   ' asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    Set rsTemp = objO.FetchData("SELECT * FROM QTY_StallWashing WHERE TipeStall <> 'Cuci'  AND ISDEL = '0' AND AKTIF = '1'")
    isPengering = False
    If rsTemp.RecordCount > 0 Then
        isPengering = True
    End If
    onTimer = False
    Call hapusForm
    Call isiCombo
    vNamaTabel = "LaporanCuci" & DriveSerialNumber("c:")
End Sub

Private Sub TblKeluar_Click()
    onTimer = True
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim txtTemp As String
Dim Txt As String
    
    boleh = False
    txtTemp = ""
    Txt = ""
    ProgressBar1.Visible = False
    prosen.Visible = False
    
    Call isiRepoCuci
    
End Sub

Private Sub isiRepoCuci()
    Dim hasilCreate As Boolean
    Dim hasilCari As ADODB.Recordset
    Dim nProg As Integer
    hasilCreate = False
    kriteria = ""
        
    If (IIf(IsNull(Tgl1), "", Tgl1)) <> "" Then
        kriteria = kriteria & " AND (A.[DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "') "
    End If
    If (IIf(IsNull(TglCuci1), "", TglCuci1)) <> "" Then
        kriteria = kriteria & " AND (T.TANGGAL >= '" & Format(TglCuci1, "YYYY-MM-DD") & " 00:00" & "' AND T.TANGGAL <= '" & Format(TglCuci2, "YYYY-MM-DD") & " 23:59" & "') "
    End If
    If cbStallCuci <> "Pilih Semua Stall" Then
        If cbStallCuci.Columns(0).value <> "" Then
            kriteria = kriteria & " AND T.NOSTALL = '" & cbStallCuci.Columns(0).value & "'"
        End If
    End If
    
    If isPengering Then
        'jika ada pengering]
'        strSql = "SELECT A.*, " & _
'                "right('0' + convert(varchar,A.LTTUNGGUCUCI / 60),2) + ':' + right('0' + convert(varchar,A.LTTUNGGUCUCI % 60),2) AS TAMPILLTTUNGGUCUCI, " & _
'                "right('0' + convert(varchar,A.LTPROSESCUCI / 60),2) + ':' + right('0' + convert(varchar,A.LTPROSESCUCI % 60),2) AS TAMPILLTPROSESCUCI, " & _
'                "right('0' + convert(varchar,A.LTTUNGGUKERING / 60),2) + ':' + right('0' + convert(varchar,A.LTTUNGGUKERING % 60),2) AS TAMPILLTTUNGGUKERING, " & _
'                "right('0' + convert(varchar,A.LTPROSESKERING / 60),2) + ':' + right('0' + convert(varchar,A.LTPROSESKERING % 60),2) AS TAMPILLTPROSESKERING, " & _
'                "right('0' + convert(varchar,A.LTTOTAL / 60),2) + ':' + right('0' + convert(varchar,A.LTTOTAL % 60),2) AS TAMPILLTTOTAL " & _
'                "FROM ( SELECT distinct A.NO_POLISI,A.NO_WO, " & _
'                "CONVERT(VARCHAR(10),P.TANGGAL,120) + ' ' + CAST(P.JAM AS DATETIME) AS SELESAIQC, " & _
'                "CONVERT(VARCHAR(10),T.TANGGAL,120) + ' ' + (SELECT TOP 1 CAST(START AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER ) AS MULAICUCI, " & _
'                "(SELECT CONVERT(VARCHAR(10),TANGGAL,120) + CAST(FINISH AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER ) AS SELESAICUCI, " & _
'                "CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME) AS MULAIKERING, " & _
'                "CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME) AS SELESAIKERING, " & _
'                "DATEDIFF(MINUTE,CONVERT(VARCHAR(10),P.TANGGAL,120) + CAST(P.JAM AS DATETIME),(SELECT CONVERT(VARCHAR(10),TANGGAL,120) + CAST(START AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER)) AS LTTUNGGUCUCI, " & _
'                "DATEDIFF(MINUTE,(SELECT CONVERT(VARCHAR(10),TANGGAL,120) + CAST(START AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER),CONVERT(VARCHAR(10),T.TANGGAL,120) + (SELECT TOP 1 CAST(FINISH AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER )) AS LTPROSESCUCI, " & _
'                "DATEDIFF(MINUTE,(SELECT CONVERT(VARCHAR(10),TANGGAL,120) + CAST(FINISH AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER ),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME)) AS LTTUNGGUKERING, " & _
'                "DATEDIFF(MINUTE,CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME)) AS LTPROSESKERING, " & _
'                "DATEDIFF(MINUTE,CONVERT(VARCHAR(10),P.TANGGAL,120) + CAST(P.JAM AS DATETIME),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME)) LTTOTAL, " & _
'                "P.jam , T.FINISH, SW.NAMASTALL " & _
'                "FROM TEMPPROGRESSCUCI T " & _
'                "INNER JOIN JOBDTA A ON T.NOORDER = A.NO_WO " & _
'                "INNER JOIN ANTRICUCI AC ON AC.NOORDER = T.NOORDER " & _
'                "INNER JOIN PROGRESS P ON P.NOORDER = T.NOORDER " & _
'                "LEFT OUTER JOIN QTY_STALLWASHING SW ON SW.NOSTALL = T.NOSTALL " & _
'                 "WHERE  T.FLAG = 'F' " & kriteria & "  ) A"
        strSql = "SELECT A.*, " & _
                "right('0' + convert(varchar,A.LTTUNGGUCUCI / 60),2) + ':' + right('0' + convert(varchar,A.LTTUNGGUCUCI % 60),2) AS TAMPILLTTUNGGUCUCI, " & _
                "right('0' + convert(varchar,A.LTPROSESCUCI / 60),2) + ':' + right('0' + convert(varchar,A.LTPROSESCUCI % 60),2) AS TAMPILLTPROSESCUCI, " & _
                "right('0' + convert(varchar,A.LTTUNGGUKERING / 60),2) + ':' + right('0' + convert(varchar,A.LTTUNGGUKERING % 60),2) AS TAMPILLTTUNGGUKERING, " & _
                "right('0' + convert(varchar,A.LTPROSESKERING / 60),2) + ':' + right('0' + convert(varchar,A.LTPROSESKERING % 60),2) AS TAMPILLTPROSESKERING, " & _
                "right('0' + convert(varchar,A.LTTOTAL / 60),2) + ':' + right('0' + convert(varchar,A.LTTOTAL % 60),2) AS TAMPILLTTOTAL " & _
                "FROM ( " & _
                "SELECT distinct A.NO_POLISI,A.NO_WO, (SELECT TOP 1 TANGGAL FROM ANTRICUCI WHERE NOORDER = T.NOORDER AND TIPESTALL = 'CUCI') AS SELESAIQC, " & _
                "CONVERT(VARCHAR(10),T.TANGGAL,120) + ' ' + (SELECT TOP 1 CAST(START AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER ) AS MULAICUCI, " & _
                "(SELECT top 1 CONVERT(VARCHAR(10),TANGGAL,120) + CAST(FINISH AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER ) AS SELESAICUCI, " & _
                "CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME) AS MULAIKERING, CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME) AS SELESAIKERING, " & _
                "DATEDIFF(MINUTE,(SELECT TOP 1 TANGGAL FROM ANTRICUCI WHERE NOORDER = T.NOORDER AND TIPESTALL = 'CUCI'),(SELECT top 1 CONVERT(VARCHAR(10),TANGGAL,120) + CAST(START AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER)) AS LTTUNGGUCUCI, " & _
                "DATEDIFF(MINUTE,(SELECT top 1 CONVERT(VARCHAR(10),TANGGAL,120) + CAST(START AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER), " & _
                "CONVERT(VARCHAR(10),T.TANGGAL,120) + (SELECT TOP 1 CAST(FINISH AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER )) AS LTPROSESCUCI, " & _
                "DATEDIFF(MINUTE,(SELECT top 1 CONVERT(VARCHAR(10),TANGGAL,120) + CAST(FINISH AS DATETIME) FROM TEMPPROGRESSCUCI WHERE FLAG = 'D' AND NOORDER = T.NOORDER ),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME)) AS LTTUNGGUKERING, " & _
                "DATEDIFF(MINUTE,CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME)) AS LTPROSESKERING, " & _
                "DATEDIFF(MINUTE,(SELECT TOP 1 TANGGAL FROM ANTRICUCI WHERE NOORDER = T.NOORDER AND TIPESTALL = 'CUCI'),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME)) LTTOTAL, " & _
                "T.FINISH, SW.NAMASTALL FROM TEMPPROGRESSCUCI T " & _
                "INNER JOIN JOBDTA A ON T.NOORDER = A.NO_WO " & _
                "INNER JOIN ANTRICUCI AC ON AC.NOORDER = T.NOORDER " & _
                "LEFT OUTER JOIN QTY_STALLWASHING SW ON SW.NOSTALL = T.NOSTALL " & _
                "WHERE  T.FLAG = 'F' " & kriteria & "  ) A"
    Else
        '- TANPA PROGRESS
        strSql = "SELECT A.*, " & _
                "right('0' + convert(varchar,A.LTTUNGGUCUCI / 60),2) + ':' + right('0' + convert(varchar,A.LTTUNGGUCUCI % 60),2) AS TAMPILLTTUNGGUCUCI, " & _
                "right('0' + convert(varchar,A.LTPROSESCUCI / 60),2) + ':' + right('0' + convert(varchar,A.LTPROSESCUCI % 60),2) AS TAMPILLTPROSESCUCI, " & _
                "right('0' + convert(varchar,A.LTTOTAL / 60),2) + ':' + right('0' + convert(varchar,A.LTTOTAL % 60),2) AS TAMPILLTTOTAL " & _
                "FROM ( SELECT distinct A.NO_POLISI,A.NO_WO, " & _
                "AC.TANGGAL AS SELESAIQC, " & _
                "CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME) AS MULAICUCI, " & _
                "CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME) AS SELESAICUCI, " & _
                "DATEDIFF(MINUTE,AC.TANGGAL,CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME)) AS LTTUNGGUCUCI, " & _
                "DATEDIFF(MINUTE,CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME)) AS LTPROSESCUCI, " & _
                "DATEDIFF(MINUTE,AC.TANGGAL,CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME)) AS LTTOTAL, " & _
                " T.FINISH, SW.NAMASTALL " & _
                "FROM TEMPPROGRESSCUCI T " & _
                "INNER JOIN JOBDTA A ON T.NOORDER = A.NO_WO " & _
                "INNER JOIN ANTRICUCI AC ON AC.NOORDER = T.NOORDER " & _
                "LEFT OUTER JOIN QTY_STALLWASHING SW ON SW.NOSTALL = T.NOSTALL " & _
                "WHERE T.FLAG = 'F' " & kriteria & " ) A"
'
'        strSql = "SELECT A.*, " & _
'                "right('0' + convert(varchar,A.LTTUNGGUCUCI / 60),2) + ':' + right('0' + convert(varchar,A.LTTUNGGUCUCI % 60),2) AS TAMPILLTTUNGGUCUCI, " & _
'                "right('0' + convert(varchar,A.LTPROSESCUCI / 60),2) + ':' + right('0' + convert(varchar,A.LTPROSESCUCI % 60),2) AS TAMPILLTPROSESCUCI, " & _
'                "right('0' + convert(varchar,A.LTTOTAL / 60),2) + ':' + right('0' + convert(varchar,A.LTTOTAL % 60),2) AS TAMPILLTTOTAL " & _
'                "FROM ( SELECT distinct A.NO_POLISI,A.NO_WO, " & _
'                "CONVERT(VARCHAR(10),P.TANGGAL,120) + CAST(P.JAM AS DATETIME) AS SELESAIQC, " & _
'                "CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME) AS MULAICUCI, " & _
'                "CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME) AS SELESAICUCI, " & _
'                "DATEDIFF(MINUTE,CONVERT(VARCHAR(10),P.TANGGAL,120) + CAST(P.JAM AS DATETIME),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME)) AS LTTUNGGUCUCI, " & _
'                "DATEDIFF(MINUTE,CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.START AS DATETIME),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME)) AS LTPROSESCUCI, " & _
'                "DATEDIFF(MINUTE,CONVERT(VARCHAR(10),P.TANGGAL,120) + CAST(P.JAM AS DATETIME),CONVERT(VARCHAR(10),T.TANGGAL,120) + CAST(T.FINISH AS DATETIME)) AS LTTOTAL, " & _
'                "P.jam , T.FINISH, SW.NAMASTALL " & _
'                "FROM TEMPPROGRESSCUCI T " & _
'                "INNER JOIN JOBDTA A ON T.NOORDER = A.NO_WO " & _
'                "INNER JOIN ANTRICUCI AC ON AC.NOORDER = T.NOORDER " & _
'                "INNER JOIN PROGRESS P ON P.NOORDER = T.NOORDER " & _
'                "LEFT OUTER JOIN QTY_STALLWASHING SW ON SW.NOSTALL = T.NOSTALL " & _
'                "WHERE T.FLAG = 'F' " & kriteria & " ) A"
    End If
    Set hasilCari = objO.FetchData(strSql)
    If hasilCari.RecordCount > 0 Then
        'Set oRs = hasilCari 'objO.FetchData(sSQL)
        hasilCari.MoveFirst
        totalWaktuTC = 0
        totalWaktuPC = 0
        totalWaktuTotal = 0
        totalWaktuTK = 0
        totalWaktuPK = 0
        
        avgWaktuPC = 0
        avgWaktuPK = 0
        avgWaktuTC = 0
        avgWaktuTK = 0
        avgWaktuTotal = 0
'        Dim Report As New rptRepoCuci
'        Dim ReportPengering As New rptRepoCuciPengering

        If isPengering Then
            Set Report = New rptRepoCuciPengering
            Do While Not hasilCari.EOF
                totalWaktuTC = totalWaktuTC + CDbl(hasilCari("LTTUNGGUCUCI"))
                totalWaktuPC = totalWaktuPC + CDbl(hasilCari("LTPROSESCUCI"))
                totalWaktuTotal = totalWaktuTotal + CDbl(hasilCari("LTTOTAL"))
                totalWaktuTK = totalWaktuTK + CDbl(hasilCari("LTTUNGGUKERING"))
                totalWaktuPK = totalWaktuPK + CDbl(hasilCari("LTPROSESKERING"))
                
                hasilCari.MoveNext
            Loop
            
            Report.database.SetDataSource hasilCari
            Report.txtTotTC.SetText IIf(totalWaktuTC > 60, Format(Val(totalWaktuTC / 60), "0#") & ":" & Format(Val(totalWaktuTC Mod 60), "0#"), "00:" & Format(Val(totalWaktuTC Mod 60), "0#"))
            Report.txtTotalPC.SetText IIf(totalWaktuPC > 60, Format(Val(totalWaktuPC / 60), "0#") & ":" & Format(Val(totalWaktuPC Mod 60), "0#"), "00:" & Format(Val(totalWaktuPC Mod 60), "0#"))
            Report.txtTotalLT.SetText IIf(totalWaktuTotal > 60, Format(Val(totalWaktuTotal / 60), "0#") & ":" & Format(Val(totalWaktuTotal Mod 60), "0#"), "00:" & Format(Val(totalWaktuTotal Mod 60), "0#"))
            Report.txtTotTK.SetText IIf(totalWaktuTK > 60, Format(Val(totalWaktuTK / 60), "0#") & ":" & Format(Val(totalWaktuTK Mod 60), "0#"), "00:" & Format(Val(totalWaktuPK Mod 60), "0#"))
            Report.txtTotPK.SetText IIf(totalWaktuPK > 60, Format(Val(totalWaktuPK / 60), "0#") & ":" & Format(Val(totalWaktuPK Mod 60), "0#"), "00:" & Format(Val(totalWaktuPK Mod 60), "0#"))
            
            avgWaktuPC = totalWaktuPC / hasilCari.RecordCount
            avgWaktuTC = totalWaktuTC / hasilCari.RecordCount
            avgWaktuTotal = totalWaktuTotal / hasilCari.RecordCount
            avgWaktuTK = totalWaktuTK / hasilCari.RecordCount
            avgWaktuPK = totalWaktuPK / hasilCari.RecordCount
            
            Report.txtAVGPC.SetText IIf(avgWaktuPC > 60, Format(Val(avgWaktuPC / 60), "0#") & ":" & Format(Val(avgWaktuPC Mod 60), "0#"), "00:" & Format(Val(avgWaktuPC Mod 60), "0#"))
            Report.txtAVGTC.SetText IIf(avgWaktuTC > 60, Format(Val(avgWaktuTC / 60), "0#") & ":" & Format(Val(avgWaktuTC Mod 60), "0#"), "00:" & Format(Val(avgWaktuTC Mod 60), "0#"))
            Report.txtAVGTotal.SetText IIf(avgWaktuTotal > 60, Format(Val(avgWaktuTotal / 60), "0#") & ":" & Format(Val(avgWaktuTotal Mod 60), "0#"), "00:" & Format(Val(avgWaktuTotal Mod 60), "0#"))
            Report.txtAvgTK.SetText IIf(avgWaktuTK > 60, Format(Val(avgWaktuTK / 60), "0#") & ":" & Format(Val(avgWaktuTK Mod 60), "0#"), "00:" & Format(Val(avgWaktuTK Mod 60), "0#"))
            Report.txtAvgPK.SetText IIf(avgWaktuPK > 60, Format(Val(avgWaktuPK / 60), "0#") & ":" & Format(Val(avgWaktuPK Mod 60), "0#"), "00:" & Format(Val(avgWaktuPK Mod 60), "0#"))
            
            Report.txtPerusahaan.SetText gPerush
            Report.txtAlamat.SetText gAlamat
            Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
            Report.txtSA.SetText gSA
            Report.txtTotalUnit.SetText hasilCari.RecordCount
            
            With CRViewer1
                .ReportSource = Report
                .ViewReport
                .Zoom 100
                .Visible = True
            End With
            
            'nWait.Visible = False
            'DoEvents
            cmdTutup.Visible = True
            cmdSettingPrinter.Visible = True
            Screen.MousePointer = vbDefault
            Call DisableExport
            Call setLokasiTombol
        Else
            Set Report = New rptRepoCuci
            Do While Not hasilCari.EOF
                totalWaktuTC = totalWaktuTC + CDbl(hasilCari("LTTUNGGUCUCI"))
                totalWaktuPC = totalWaktuPC + CDbl(hasilCari("LTPROSESCUCI"))
                totalWaktuTotal = totalWaktuTotal + CDbl(hasilCari("LTTOTAL"))
                
                hasilCari.MoveNext
            Loop
            
            Report.database.SetDataSource hasilCari
            Report.txtTotalTC.SetText IIf(totalWaktuTC > 60, Format(Val(totalWaktuTC / 60), "0#") & ":" & Format(Val(totalWaktuTC Mod 60), "0#"), "00:" & Format(Val(totalWaktuTC Mod 60), "0#"))
            Report.txtTotalPC.SetText IIf(totalWaktuPC > 60, Format(Val(totalWaktuPC / 60), "0#") & ":" & Format(Val(totalWaktuPC Mod 60), "0#"), "00:" & Format(Val(totalWaktuPC Mod 60), "0#"))
            Report.txtTotalLT.SetText IIf(totalWaktuTotal > 60, Format(Val(totalWaktuTotal / 60), "0#") & ":" & Format(Val(totalWaktuTotal Mod 60), "0#"), "00:" & Format(Val(totalWaktuTotal Mod 60), "0#"))
            
            avgWaktuPC = totalWaktuPC / hasilCari.RecordCount
            avgWaktuTC = totalWaktuTC / hasilCari.RecordCount
            avgWaktuTotal = totalWaktuTotal / hasilCari.RecordCount
            
            Report.txtAVGPC.SetText IIf(avgWaktuPC > 60, Format(Val(avgWaktuPC / 60), "0#") & ":" & Format(Val(avgWaktuPC Mod 60), "0#"), "00:" & Format(Val(avgWaktuPC Mod 60), "0#"))
            Report.txtAVGTC.SetText IIf(avgWaktuTC > 60, Format(Val(avgWaktuTC / 60), "0#") & ":" & Format(Val(avgWaktuTC Mod 60), "0#"), "00:" & Format(Val(avgWaktuTC Mod 60), "0#"))
            Report.txtAVGTotal.SetText IIf(avgWaktuTotal > 60, Format(Val(avgWaktuTotal / 60), "0#") & ":" & Format(Val(avgWaktuTotal Mod 60), "0#"), "00:" & Format(Val(avgWaktuTotal Mod 60), "0#"))
            
            Report.txtPerusahaan.SetText gPerush
            Report.txtAlamat.SetText gAlamat
            Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
            Report.txtSA.SetText gSA
            Report.txtTotalUnit.SetText hasilCari.RecordCount
            
            With CRViewer1
                .ReportSource = Report
                .ViewReport
                .Zoom 100
                .Visible = True
            End With
            
            'nWait.Visible = False
            'DoEvents
            cmdTutup.Visible = True
            cmdSettingPrinter.Visible = True
            Screen.MousePointer = vbDefault
            Call DisableExport
            Call setLokasiTombol
        End If
    Else
        MsgBox "Tidak ada data yang ditampilkan", vbInformation
    End If
    
End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 345
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdCetak.Top = CRViewer1.Top + 345
    cmdCetak.Left = cboPrinter.Left + cboPrinter.Width + 30
End Sub

Private Sub DisableExport()
    Dim objPrinter As Printer
    Set objPrinter = Printer
    If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
        cboPrinter = ""
    Else
        cboPrinter = objPrinter.DeviceName
    End If
    cboPrinter.Clear
    
    For Each pr In Printers
        If InStr(UCase(pr.DeviceName), "PDF") = 0 Then cboPrinter.AddItem pr.DeviceName
    Next pr
    If Log.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        CRViewer1.EnablePrintButton = True
        cmdExport.Visible = False
        cboPrinter.Visible = False
        cmdCetak.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdCetak_Click()
    For Each pr In Printers
        If UCase(pr.DeviceName) = UCase(cboPrinter) Then
            GoTo Cetak
        End If
    Next pr
    MsgBox "Printer tidak ditemukan Pilih Printer dahulu", vbInformation, "Proses Cetak Gagal"
    cboPrinter.SetFocus
    Exit Sub
    
Cetak:
    Dim Psize As Integer
    Dim POr As Integer
    
    Psize = Report.PaperSize
    POr = Report.PaperOrientation
    
    Set Printer = Printers(cboPrinter.ListIndex)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    Report.PaperOrientation = POr
    Report.PaperSize = Psize
    Report.PrintOut
    
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Salah
    Dim NamaFile As String
    NamaFile = ""
    
    Dim f As New FormExplorer
    f.Show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    
    If Log.PdfOnly = 1 Then
         
        NamaFile = MyPath & f.FileName
        
        If UCase(Right(Trim(f.FileName), 5)) = ".XLSX" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 5)
        If UCase(Right(Trim(f.FileName), 4)) = ".XLS" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 4)
         
        
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
        Report.ExportOptions.RTFExportAllPages = True
        Report.ExportOptions.FormatType = crEFTExactRichText
        Report.Export False
        MsgBox NamaFile & ".Doc" & "  ... Data sudah di Export !", vbInformation
        Exit Sub
        With msg
            .PassingMsg = NamaFile & ".Doc" & "  ... Data sudah di Export !"
            .passingTipe = 2
            .Show vbModal
        End With
        Exit Sub
    ElseIf onExpo = 0 Then
        Report.Export
    End If
    Exit Sub
Salah:
    If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
    MsgBox "Err.Line : " & Erl & "Err.Numb : " & err.Number & "Err.Desc. : " & err.Description & vbCr & "Procedure : " & Me.Name & " export"
Keluar:
End Sub




