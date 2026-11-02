VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmInfoPelanggan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Daftar Pelanggan"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdExportCSV 
      Caption         =   "EXPORT CSV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   94
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   6120
      TabIndex        =   84
      Top             =   4080
      Visible         =   0   'False
      Width           =   4575
      Begin VB.ComboBox cbBulanRet 
         Height          =   405
         ItemData        =   "frmInfoPelanggan.frx":0000
         Left            =   1560
         List            =   "frmInfoPelanggan.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   600
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   495
         Left            =   240
         TabIndex        =   91
         Top             =   1560
         Visible         =   0   'False
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H000000FF&
         Caption         =   "BATAL"
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080FFFF&
         Caption         =   "PRINT"
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   1560
         Width           =   1215
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   375
         Left            =   1560
         TabIndex        =   93
         Top             =   1080
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "frmInfoPelanggan.frx":0053
         Caption         =   "frmInfoPelanggan.frx":0073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":00DF
         Keys            =   "frmInfoPelanggan.frx":00FD
         Spin            =   "frmInfoPelanggan.frx":0147
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2500
         MinValue        =   1900
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   2015
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "KRITERIA SB RETENTION"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   90
         Top             =   60
         Width           =   4575
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   4575
      End
      Begin VB.Label lblPengambilanDataRet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   89
         Top             =   2160
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "TAHUN :"
         Height          =   375
         Left            =   120
         TabIndex        =   86
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "BULAN :"
         Height          =   375
         Left            =   120
         TabIndex        =   85
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CUSTOMER MAPPING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   2400
      TabIndex        =   76
      Top             =   2160
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton Command7 
         BackColor       =   &H0000FF00&
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   5760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H000000FF&
         Caption         =   "TUTUP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   120
         Width           =   855
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "frmInfoPelanggan.frx":016F
         TabIndex        =   77
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   14520
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton cmdRetension 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SB Retention"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H000080FF&
      Caption         =   "EXPORT XLS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Frame frameWarna 
      Caption         =   "Pilihan Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   11280
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default (Putih)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Biru"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000FF00&
         Caption         =   "Hijau"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Kuning"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pilih Warna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000000FF&
         Caption         =   "Merah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FramePasif2Range 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   62
      Top             =   840
      Width           =   2775
      Begin TDBNumber6Ctl.TDBNumber blnPasif2Awal 
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":298A
         Caption         =   "frmInfoPelanggan.frx":29AA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":2A16
         Keys            =   "frmInfoPelanggan.frx":2A34
         Spin            =   "frmInfoPelanggan.frx":2A7E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   13
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber blnPasif2Akhir 
         Height          =   255
         Left            =   1560
         TabIndex        =   64
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":2AA6
         Caption         =   "frmInfoPelanggan.frx":2AC6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":2B32
         Keys            =   "frmInfoPelanggan.frx":2B50
         Spin            =   "frmInfoPelanggan.frx":2B9A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   18
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan s.d "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   600
         TabIndex        =   66
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   2040
         TabIndex        =   65
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.Frame framePasif1Range 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   30
      Top             =   480
      Width           =   2775
      Begin TDBNumber6Ctl.TDBNumber blnRangePasif1Awal 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":2BC2
         Caption         =   "frmInfoPelanggan.frx":2BE2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":2C4E
         Keys            =   "frmInfoPelanggan.frx":2C6C
         Spin            =   "frmInfoPelanggan.frx":2CB6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   7
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber blnRangePasif1Akhir 
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":2CDE
         Caption         =   "frmInfoPelanggan.frx":2CFE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":2D6A
         Keys            =   "frmInfoPelanggan.frx":2D88
         Spin            =   "frmInfoPelanggan.frx":2DD2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   12
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan s.d "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   34
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   33
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   1200
      Width           =   615
   End
   Begin VB.Frame FramePasif2Bulan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   69
      Top             =   840
      Width           =   1335
      Begin TDBNumber6Ctl.TDBNumber blnPasif2Bulan 
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":2DFA
         Caption         =   "frmInfoPelanggan.frx":2E1A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":2E86
         Keys            =   "frmInfoPelanggan.frx":2EA4
         Spin            =   "frmInfoPelanggan.frx":2EEE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   71
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "frmInfoPelanggan.frx":2F16
      Left            =   7800
      List            =   "frmInfoPelanggan.frx":2F26
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox ckPasif2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategori Pasif 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   67
      Top             =   840
      Width           =   1815
   End
   Begin VB.Frame FramePasif3Range 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   57
      Top             =   1200
      Width           =   2775
      Begin TDBNumber6Ctl.TDBNumber blnPasif3Awal 
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":2F3A
         Caption         =   "frmInfoPelanggan.frx":2F5A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":2FC6
         Keys            =   "frmInfoPelanggan.frx":2FE4
         Spin            =   "frmInfoPelanggan.frx":302E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   6
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber blnPasif3Akhir 
         Height          =   255
         Left            =   1560
         TabIndex        =   59
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":3056
         Caption         =   "frmInfoPelanggan.frx":3076
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":30E2
         Keys            =   "frmInfoPelanggan.frx":3100
         Spin            =   "frmInfoPelanggan.frx":314A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   6
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   2040
         TabIndex        =   61
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan s.d "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   60
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.CheckBox ckPasif3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategori Pasif 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   56
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "frmInfoPelanggan.frx":3172
      Left            =   7800
      List            =   "frmInfoPelanggan.frx":3182
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame FramePasif3Bulan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8880
      TabIndex        =   52
      Top             =   1200
      Width           =   1335
      Begin TDBNumber6Ctl.TDBNumber blnPasif3Bulan 
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":3196
         Caption         =   "frmInfoPelanggan.frx":31B6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":3222
         Keys            =   "frmInfoPelanggan.frx":3240
         Spin            =   "frmInfoPelanggan.frx":328A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   18
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   600
         TabIndex        =   54
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdWarna 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame framePasifBulan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8880
      TabIndex        =   37
      Top             =   480
      Width           =   1335
      Begin TDBNumber6Ctl.TDBNumber blnPasif1 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":32B2
         Caption         =   "frmInfoPelanggan.frx":32D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":333E
         Keys            =   "frmInfoPelanggan.frx":335C
         Spin            =   "frmInfoPelanggan.frx":33A6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1376261
         Value           =   1
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   39
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmInfoPelanggan.frx":33CE
      Left            =   7800
      List            =   "frmInfoPelanggan.frx":33DE
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox ckPasif1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategori Pasif 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   35
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame FrameAktifRange 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   25
      Top             =   120
      Width           =   2775
      Begin TDBNumber6Ctl.TDBNumber blnRangeAktifAwal 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":33F2
         Caption         =   "frmInfoPelanggan.frx":3412
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":347E
         Keys            =   "frmInfoPelanggan.frx":349C
         Spin            =   "frmInfoPelanggan.frx":34E6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   6
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber blnRangeAktifAkhir 
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":350E
         Caption         =   "frmInfoPelanggan.frx":352E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":359A
         Keys            =   "frmInfoPelanggan.frx":35B8
         Spin            =   "frmInfoPelanggan.frx":3602
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   6
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2040
         TabIndex        =   29
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan s.d "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   27
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.CheckBox ckAktif 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategori Aktif"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data HO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cbOperator 
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
      ItemData        =   "frmInfoPelanggan.frx":362A
      Left            =   7800
      List            =   "frmInfoPelanggan.frx":363A
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtAlamat 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   1200
      Width           =   3495
   End
   Begin TrueOleDBGrid70.TDBGrid gridTagih 
      Height          =   5535
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   9763
      _LayoutType     =   4
      _RowHeight      =   25
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
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
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).DataField=   ""
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).DataField=   ""
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).DataField=   ""
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).DataField=   ""
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).DataField=   ""
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).DataField=   ""
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).DataField=   ""
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(42)._VlistStyle=   0
      Columns(42)._MaxComboItems=   5
      Columns(42).DataField=   ""
      Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(43)._VlistStyle=   0
      Columns(43)._MaxComboItems=   5
      Columns(43).DataField=   ""
      Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(44)._VlistStyle=   0
      Columns(44)._MaxComboItems=   5
      Columns(44).DataField=   ""
      Columns(44)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(45)._VlistStyle=   0
      Columns(45)._MaxComboItems=   5
      Columns(45).DataField=   ""
      Columns(45)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(46)._VlistStyle=   0
      Columns(46)._MaxComboItems=   5
      Columns(46).DataField=   ""
      Columns(46)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(47)._VlistStyle=   0
      Columns(47)._MaxComboItems=   5
      Columns(47).DataField=   ""
      Columns(47)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(48)._VlistStyle=   0
      Columns(48)._MaxComboItems=   5
      Columns(48).DataField=   ""
      Columns(48)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(49)._VlistStyle=   0
      Columns(49)._MaxComboItems=   5
      Columns(49).DataField=   ""
      Columns(49)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(50)._VlistStyle=   0
      Columns(50)._MaxComboItems=   5
      Columns(50).DataField=   ""
      Columns(50)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(51)._VlistStyle=   0
      Columns(51)._MaxComboItems=   5
      Columns(51).DataField=   ""
      Columns(51)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(52)._VlistStyle=   0
      Columns(52)._MaxComboItems=   5
      Columns(52).DataField=   ""
      Columns(52)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(53)._VlistStyle=   0
      Columns(53)._MaxComboItems=   5
      Columns(53).DataField=   ""
      Columns(53)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(54)._VlistStyle=   0
      Columns(54)._MaxComboItems=   5
      Columns(54).DataField=   ""
      Columns(54)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(55)._VlistStyle=   0
      Columns(55)._MaxComboItems=   5
      Columns(55).DataField=   ""
      Columns(55)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(56)._VlistStyle=   0
      Columns(56)._MaxComboItems=   5
      Columns(56).DataField=   ""
      Columns(56)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(57)._VlistStyle=   0
      Columns(57)._MaxComboItems=   5
      Columns(57).DataField=   ""
      Columns(57)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(58)._VlistStyle=   0
      Columns(58)._MaxComboItems=   5
      Columns(58).DataField=   ""
      Columns(58)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(59)._VlistStyle=   0
      Columns(59)._MaxComboItems=   5
      Columns(59).DataField=   ""
      Columns(59)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(60)._VlistStyle=   0
      Columns(60)._MaxComboItems=   5
      Columns(60).DataField=   ""
      Columns(60)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(61)._VlistStyle=   0
      Columns(61)._MaxComboItems=   5
      Columns(61).DataField=   ""
      Columns(61)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(62)._VlistStyle=   0
      Columns(62)._MaxComboItems=   5
      Columns(62).DataField=   ""
      Columns(62)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(63)._VlistStyle=   0
      Columns(63)._MaxComboItems=   5
      Columns(63).DataField=   ""
      Columns(63)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(64)._VlistStyle=   0
      Columns(64)._MaxComboItems=   5
      Columns(64).DataField=   ""
      Columns(64)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(65)._VlistStyle=   0
      Columns(65)._MaxComboItems=   5
      Columns(65).DataField=   ""
      Columns(65)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(66)._VlistStyle=   0
      Columns(66)._MaxComboItems=   5
      Columns(66).DataField=   ""
      Columns(66)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(67)._VlistStyle=   0
      Columns(67)._MaxComboItems=   5
      Columns(67).DataField=   ""
      Columns(67)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(68)._VlistStyle=   0
      Columns(68)._MaxComboItems=   5
      Columns(68).DataField=   ""
      Columns(68)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(69)._VlistStyle=   0
      Columns(69)._MaxComboItems=   5
      Columns(69).DataField=   ""
      Columns(69)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(70)._VlistStyle=   0
      Columns(70)._MaxComboItems=   5
      Columns(70).DataField=   ""
      Columns(70)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(71)._VlistStyle=   0
      Columns(71)._MaxComboItems=   5
      Columns(71).DataField=   ""
      Columns(71)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(72)._VlistStyle=   0
      Columns(72)._MaxComboItems=   5
      Columns(72).DataField=   ""
      Columns(72)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(73)._VlistStyle=   0
      Columns(73)._MaxComboItems=   5
      Columns(73).DataField=   ""
      Columns(73)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(74)._VlistStyle=   0
      Columns(74)._MaxComboItems=   5
      Columns(74).DataField=   ""
      Columns(74)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(75)._VlistStyle=   0
      Columns(75)._MaxComboItems=   5
      Columns(75).DataField=   ""
      Columns(75)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(76)._VlistStyle=   0
      Columns(76)._MaxComboItems=   5
      Columns(76).DataField=   ""
      Columns(76)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(77)._VlistStyle=   0
      Columns(77)._MaxComboItems=   5
      Columns(77).DataField=   ""
      Columns(77)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(78)._VlistStyle=   0
      Columns(78)._MaxComboItems=   5
      Columns(78).DataField=   ""
      Columns(78)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(79)._VlistStyle=   0
      Columns(79)._MaxComboItems=   5
      Columns(79).DataField=   ""
      Columns(79)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(80)._VlistStyle=   0
      Columns(80)._MaxComboItems=   5
      Columns(80).DataField=   ""
      Columns(80)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(81)._VlistStyle=   0
      Columns(81)._MaxComboItems=   5
      Columns(81).DataField=   ""
      Columns(81)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(82)._VlistStyle=   0
      Columns(82)._MaxComboItems=   5
      Columns(82).DataField=   ""
      Columns(82)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(83)._VlistStyle=   0
      Columns(83)._MaxComboItems=   5
      Columns(83).DataField=   ""
      Columns(83)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(84)._VlistStyle=   0
      Columns(84)._MaxComboItems=   5
      Columns(84).DataField=   ""
      Columns(84)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(85)._VlistStyle=   0
      Columns(85)._MaxComboItems=   5
      Columns(85).DataField=   ""
      Columns(85)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(86)._VlistStyle=   0
      Columns(86)._MaxComboItems=   5
      Columns(86).DataField=   ""
      Columns(86)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(87)._VlistStyle=   0
      Columns(87)._MaxComboItems=   5
      Columns(87).DataField=   ""
      Columns(87)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(88)._VlistStyle=   0
      Columns(88)._MaxComboItems=   5
      Columns(88).DataField=   ""
      Columns(88)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(89)._VlistStyle=   0
      Columns(89)._MaxComboItems=   5
      Columns(89).DataField=   ""
      Columns(89)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(90)._VlistStyle=   0
      Columns(90)._MaxComboItems=   5
      Columns(90).DataField=   ""
      Columns(90)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(91)._VlistStyle=   0
      Columns(91)._MaxComboItems=   5
      Columns(91).DataField=   ""
      Columns(91)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(92)._VlistStyle=   0
      Columns(92)._MaxComboItems=   5
      Columns(92).DataField=   ""
      Columns(92)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(93)._VlistStyle=   0
      Columns(93)._MaxComboItems=   5
      Columns(93).DataField=   ""
      Columns(93)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(94)._VlistStyle=   0
      Columns(94)._MaxComboItems=   5
      Columns(94).DataField=   ""
      Columns(94)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(95)._VlistStyle=   0
      Columns(95)._MaxComboItems=   5
      Columns(95).DataField=   ""
      Columns(95)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(96)._VlistStyle=   0
      Columns(96)._MaxComboItems=   5
      Columns(96).DataField=   ""
      Columns(96)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(97)._VlistStyle=   0
      Columns(97)._MaxComboItems=   5
      Columns(97).DataField=   ""
      Columns(97)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(98)._VlistStyle=   0
      Columns(98)._MaxComboItems=   5
      Columns(98).DataField=   ""
      Columns(98)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(99)._VlistStyle=   0
      Columns(99)._MaxComboItems=   5
      Columns(99).DataField=   ""
      Columns(99)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(100)._VlistStyle=   0
      Columns(100)._MaxComboItems=   5
      Columns(100).DataField=   ""
      Columns(100)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(101)._VlistStyle=   0
      Columns(101)._MaxComboItems=   5
      Columns(101).DataField=   ""
      Columns(101)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(102)._VlistStyle=   0
      Columns(102)._MaxComboItems=   5
      Columns(102).DataField=   ""
      Columns(102)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(103)._VlistStyle=   0
      Columns(103)._MaxComboItems=   5
      Columns(103).DataField=   ""
      Columns(103)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(104)._VlistStyle=   0
      Columns(104)._MaxComboItems=   5
      Columns(104).DataField=   ""
      Columns(104)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   105
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   2
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=105"
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
      Splits(0)._ColumnProps(53)=   "Column(13).Width=3281"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(69)=   "Column(17).Width=3281"
      Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=3175"
      Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(73)=   "Column(18).Width=3281"
      Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=3175"
      Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(77)=   "Column(19).Width=3281"
      Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=3175"
      Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(81)=   "Column(20).Width=3281"
      Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=3175"
      Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(85)=   "Column(21).Width=3281"
      Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=3175"
      Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(89)=   "Column(22).Width=3281"
      Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=3175"
      Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(93)=   "Column(23).Width=3281"
      Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=3175"
      Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(97)=   "Column(24).Width=3281"
      Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=3175"
      Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(101)=   "Column(25).Width=3281"
      Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=3175"
      Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(105)=   "Column(26).Width=3281"
      Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=3175"
      Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(109)=   "Column(27).Width=3281"
      Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=3175"
      Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(113)=   "Column(28).Width=3281"
      Splits(0)._ColumnProps(114)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(28)._WidthInPix=3175"
      Splits(0)._ColumnProps(116)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(117)=   "Column(29).Width=3281"
      Splits(0)._ColumnProps(118)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(29)._WidthInPix=3175"
      Splits(0)._ColumnProps(120)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(121)=   "Column(30).Width=3281"
      Splits(0)._ColumnProps(122)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(30)._WidthInPix=3175"
      Splits(0)._ColumnProps(124)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(125)=   "Column(31).Width=3281"
      Splits(0)._ColumnProps(126)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(31)._WidthInPix=3175"
      Splits(0)._ColumnProps(128)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(129)=   "Column(32).Width=3281"
      Splits(0)._ColumnProps(130)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(32)._WidthInPix=3175"
      Splits(0)._ColumnProps(132)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(133)=   "Column(33).Width=3281"
      Splits(0)._ColumnProps(134)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(33)._WidthInPix=3175"
      Splits(0)._ColumnProps(136)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(137)=   "Column(34).Width=3281"
      Splits(0)._ColumnProps(138)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(34)._WidthInPix=3175"
      Splits(0)._ColumnProps(140)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(141)=   "Column(35).Width=3281"
      Splits(0)._ColumnProps(142)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(143)=   "Column(35)._WidthInPix=3175"
      Splits(0)._ColumnProps(144)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(145)=   "Column(36).Width=3281"
      Splits(0)._ColumnProps(146)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(147)=   "Column(36)._WidthInPix=3175"
      Splits(0)._ColumnProps(148)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(149)=   "Column(37).Width=3281"
      Splits(0)._ColumnProps(150)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(37)._WidthInPix=3175"
      Splits(0)._ColumnProps(152)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(153)=   "Column(38).Width=3281"
      Splits(0)._ColumnProps(154)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(155)=   "Column(38)._WidthInPix=3175"
      Splits(0)._ColumnProps(156)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(157)=   "Column(39).Width=3281"
      Splits(0)._ColumnProps(158)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(159)=   "Column(39)._WidthInPix=3175"
      Splits(0)._ColumnProps(160)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(161)=   "Column(40).Width=3281"
      Splits(0)._ColumnProps(162)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(40)._WidthInPix=3175"
      Splits(0)._ColumnProps(164)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(165)=   "Column(41).Width=3281"
      Splits(0)._ColumnProps(166)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(167)=   "Column(41)._WidthInPix=3175"
      Splits(0)._ColumnProps(168)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(169)=   "Column(42).Width=3281"
      Splits(0)._ColumnProps(170)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(171)=   "Column(42)._WidthInPix=3175"
      Splits(0)._ColumnProps(172)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(173)=   "Column(43).Width=3281"
      Splits(0)._ColumnProps(174)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(43)._WidthInPix=3175"
      Splits(0)._ColumnProps(176)=   "Column(43).Order=44"
      Splits(0)._ColumnProps(177)=   "Column(44).Width=3281"
      Splits(0)._ColumnProps(178)=   "Column(44).DividerColor=0"
      Splits(0)._ColumnProps(179)=   "Column(44)._WidthInPix=3175"
      Splits(0)._ColumnProps(180)=   "Column(44).Order=45"
      Splits(0)._ColumnProps(181)=   "Column(45).Width=3281"
      Splits(0)._ColumnProps(182)=   "Column(45).DividerColor=0"
      Splits(0)._ColumnProps(183)=   "Column(45)._WidthInPix=3175"
      Splits(0)._ColumnProps(184)=   "Column(45).Order=46"
      Splits(0)._ColumnProps(185)=   "Column(46).Width=3281"
      Splits(0)._ColumnProps(186)=   "Column(46).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(46)._WidthInPix=3175"
      Splits(0)._ColumnProps(188)=   "Column(46).Order=47"
      Splits(0)._ColumnProps(189)=   "Column(47).Width=3281"
      Splits(0)._ColumnProps(190)=   "Column(47).DividerColor=0"
      Splits(0)._ColumnProps(191)=   "Column(47)._WidthInPix=3175"
      Splits(0)._ColumnProps(192)=   "Column(47).Order=48"
      Splits(0)._ColumnProps(193)=   "Column(48).Width=3281"
      Splits(0)._ColumnProps(194)=   "Column(48).DividerColor=0"
      Splits(0)._ColumnProps(195)=   "Column(48)._WidthInPix=3175"
      Splits(0)._ColumnProps(196)=   "Column(48).Order=49"
      Splits(0)._ColumnProps(197)=   "Column(49).Width=3281"
      Splits(0)._ColumnProps(198)=   "Column(49).DividerColor=0"
      Splits(0)._ColumnProps(199)=   "Column(49)._WidthInPix=3175"
      Splits(0)._ColumnProps(200)=   "Column(49).Order=50"
      Splits(0)._ColumnProps(201)=   "Column(50).Width=3281"
      Splits(0)._ColumnProps(202)=   "Column(50).DividerColor=0"
      Splits(0)._ColumnProps(203)=   "Column(50)._WidthInPix=3175"
      Splits(0)._ColumnProps(204)=   "Column(50).Order=51"
      Splits(0)._ColumnProps(205)=   "Column(51).Width=3281"
      Splits(0)._ColumnProps(206)=   "Column(51).DividerColor=0"
      Splits(0)._ColumnProps(207)=   "Column(51)._WidthInPix=3175"
      Splits(0)._ColumnProps(208)=   "Column(51).Order=52"
      Splits(0)._ColumnProps(209)=   "Column(52).Width=3281"
      Splits(0)._ColumnProps(210)=   "Column(52).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(52)._WidthInPix=3175"
      Splits(0)._ColumnProps(212)=   "Column(52).Order=53"
      Splits(0)._ColumnProps(213)=   "Column(53).Width=3281"
      Splits(0)._ColumnProps(214)=   "Column(53).DividerColor=0"
      Splits(0)._ColumnProps(215)=   "Column(53)._WidthInPix=3175"
      Splits(0)._ColumnProps(216)=   "Column(53).Order=54"
      Splits(0)._ColumnProps(217)=   "Column(54).Width=3281"
      Splits(0)._ColumnProps(218)=   "Column(54).DividerColor=0"
      Splits(0)._ColumnProps(219)=   "Column(54)._WidthInPix=3175"
      Splits(0)._ColumnProps(220)=   "Column(54).Order=55"
      Splits(0)._ColumnProps(221)=   "Column(55).Width=3281"
      Splits(0)._ColumnProps(222)=   "Column(55).DividerColor=0"
      Splits(0)._ColumnProps(223)=   "Column(55)._WidthInPix=3175"
      Splits(0)._ColumnProps(224)=   "Column(55).Order=56"
      Splits(0)._ColumnProps(225)=   "Column(56).Width=3281"
      Splits(0)._ColumnProps(226)=   "Column(56).DividerColor=0"
      Splits(0)._ColumnProps(227)=   "Column(56)._WidthInPix=3175"
      Splits(0)._ColumnProps(228)=   "Column(56).Order=57"
      Splits(0)._ColumnProps(229)=   "Column(57).Width=3281"
      Splits(0)._ColumnProps(230)=   "Column(57).DividerColor=0"
      Splits(0)._ColumnProps(231)=   "Column(57)._WidthInPix=3175"
      Splits(0)._ColumnProps(232)=   "Column(57).Order=58"
      Splits(0)._ColumnProps(233)=   "Column(58).Width=3281"
      Splits(0)._ColumnProps(234)=   "Column(58).DividerColor=0"
      Splits(0)._ColumnProps(235)=   "Column(58)._WidthInPix=3175"
      Splits(0)._ColumnProps(236)=   "Column(58).Order=59"
      Splits(0)._ColumnProps(237)=   "Column(59).Width=3281"
      Splits(0)._ColumnProps(238)=   "Column(59).DividerColor=0"
      Splits(0)._ColumnProps(239)=   "Column(59)._WidthInPix=3175"
      Splits(0)._ColumnProps(240)=   "Column(59).Order=60"
      Splits(0)._ColumnProps(241)=   "Column(60).Width=3281"
      Splits(0)._ColumnProps(242)=   "Column(60).DividerColor=0"
      Splits(0)._ColumnProps(243)=   "Column(60)._WidthInPix=3175"
      Splits(0)._ColumnProps(244)=   "Column(60).Order=61"
      Splits(0)._ColumnProps(245)=   "Column(61).Width=3281"
      Splits(0)._ColumnProps(246)=   "Column(61).DividerColor=0"
      Splits(0)._ColumnProps(247)=   "Column(61)._WidthInPix=3175"
      Splits(0)._ColumnProps(248)=   "Column(61).Order=62"
      Splits(0)._ColumnProps(249)=   "Column(62).Width=3281"
      Splits(0)._ColumnProps(250)=   "Column(62).DividerColor=0"
      Splits(0)._ColumnProps(251)=   "Column(62)._WidthInPix=3175"
      Splits(0)._ColumnProps(252)=   "Column(62).Order=63"
      Splits(0)._ColumnProps(253)=   "Column(63).Width=3281"
      Splits(0)._ColumnProps(254)=   "Column(63).DividerColor=0"
      Splits(0)._ColumnProps(255)=   "Column(63)._WidthInPix=3175"
      Splits(0)._ColumnProps(256)=   "Column(63).Order=64"
      Splits(0)._ColumnProps(257)=   "Column(64).Width=3281"
      Splits(0)._ColumnProps(258)=   "Column(64).DividerColor=0"
      Splits(0)._ColumnProps(259)=   "Column(64)._WidthInPix=3175"
      Splits(0)._ColumnProps(260)=   "Column(64).Order=65"
      Splits(0)._ColumnProps(261)=   "Column(65).Width=3281"
      Splits(0)._ColumnProps(262)=   "Column(65).DividerColor=0"
      Splits(0)._ColumnProps(263)=   "Column(65)._WidthInPix=3175"
      Splits(0)._ColumnProps(264)=   "Column(65).Order=66"
      Splits(0)._ColumnProps(265)=   "Column(66).Width=3281"
      Splits(0)._ColumnProps(266)=   "Column(66).DividerColor=0"
      Splits(0)._ColumnProps(267)=   "Column(66)._WidthInPix=3175"
      Splits(0)._ColumnProps(268)=   "Column(66).Order=67"
      Splits(0)._ColumnProps(269)=   "Column(67).Width=3281"
      Splits(0)._ColumnProps(270)=   "Column(67).DividerColor=0"
      Splits(0)._ColumnProps(271)=   "Column(67)._WidthInPix=3175"
      Splits(0)._ColumnProps(272)=   "Column(67).Order=68"
      Splits(0)._ColumnProps(273)=   "Column(68).Width=3281"
      Splits(0)._ColumnProps(274)=   "Column(68).DividerColor=0"
      Splits(0)._ColumnProps(275)=   "Column(68)._WidthInPix=3175"
      Splits(0)._ColumnProps(276)=   "Column(68).Order=69"
      Splits(0)._ColumnProps(277)=   "Column(69).Width=3281"
      Splits(0)._ColumnProps(278)=   "Column(69).DividerColor=0"
      Splits(0)._ColumnProps(279)=   "Column(69)._WidthInPix=3175"
      Splits(0)._ColumnProps(280)=   "Column(69).Order=70"
      Splits(0)._ColumnProps(281)=   "Column(70).Width=3281"
      Splits(0)._ColumnProps(282)=   "Column(70).DividerColor=0"
      Splits(0)._ColumnProps(283)=   "Column(70)._WidthInPix=3175"
      Splits(0)._ColumnProps(284)=   "Column(70).Order=71"
      Splits(0)._ColumnProps(285)=   "Column(71).Width=3281"
      Splits(0)._ColumnProps(286)=   "Column(71).DividerColor=0"
      Splits(0)._ColumnProps(287)=   "Column(71)._WidthInPix=3175"
      Splits(0)._ColumnProps(288)=   "Column(71).Order=72"
      Splits(0)._ColumnProps(289)=   "Column(72).Width=3281"
      Splits(0)._ColumnProps(290)=   "Column(72).DividerColor=0"
      Splits(0)._ColumnProps(291)=   "Column(72)._WidthInPix=3175"
      Splits(0)._ColumnProps(292)=   "Column(72).Order=73"
      Splits(0)._ColumnProps(293)=   "Column(73).Width=3281"
      Splits(0)._ColumnProps(294)=   "Column(73).DividerColor=0"
      Splits(0)._ColumnProps(295)=   "Column(73)._WidthInPix=3175"
      Splits(0)._ColumnProps(296)=   "Column(73).Order=74"
      Splits(0)._ColumnProps(297)=   "Column(74).Width=3281"
      Splits(0)._ColumnProps(298)=   "Column(74).DividerColor=0"
      Splits(0)._ColumnProps(299)=   "Column(74)._WidthInPix=3175"
      Splits(0)._ColumnProps(300)=   "Column(74).Order=75"
      Splits(0)._ColumnProps(301)=   "Column(75).Width=3281"
      Splits(0)._ColumnProps(302)=   "Column(75).DividerColor=0"
      Splits(0)._ColumnProps(303)=   "Column(75)._WidthInPix=3175"
      Splits(0)._ColumnProps(304)=   "Column(75).Order=76"
      Splits(0)._ColumnProps(305)=   "Column(76).Width=3281"
      Splits(0)._ColumnProps(306)=   "Column(76).DividerColor=0"
      Splits(0)._ColumnProps(307)=   "Column(76)._WidthInPix=3175"
      Splits(0)._ColumnProps(308)=   "Column(76).Order=77"
      Splits(0)._ColumnProps(309)=   "Column(77).Width=3281"
      Splits(0)._ColumnProps(310)=   "Column(77).DividerColor=0"
      Splits(0)._ColumnProps(311)=   "Column(77)._WidthInPix=3175"
      Splits(0)._ColumnProps(312)=   "Column(77).Order=78"
      Splits(0)._ColumnProps(313)=   "Column(78).Width=3281"
      Splits(0)._ColumnProps(314)=   "Column(78).DividerColor=0"
      Splits(0)._ColumnProps(315)=   "Column(78)._WidthInPix=3175"
      Splits(0)._ColumnProps(316)=   "Column(78).Order=79"
      Splits(0)._ColumnProps(317)=   "Column(79).Width=3281"
      Splits(0)._ColumnProps(318)=   "Column(79).DividerColor=0"
      Splits(0)._ColumnProps(319)=   "Column(79)._WidthInPix=3175"
      Splits(0)._ColumnProps(320)=   "Column(79).Order=80"
      Splits(0)._ColumnProps(321)=   "Column(80).Width=3281"
      Splits(0)._ColumnProps(322)=   "Column(80).DividerColor=0"
      Splits(0)._ColumnProps(323)=   "Column(80)._WidthInPix=3175"
      Splits(0)._ColumnProps(324)=   "Column(80).Order=81"
      Splits(0)._ColumnProps(325)=   "Column(81).Width=3281"
      Splits(0)._ColumnProps(326)=   "Column(81).DividerColor=0"
      Splits(0)._ColumnProps(327)=   "Column(81)._WidthInPix=3175"
      Splits(0)._ColumnProps(328)=   "Column(81).Order=82"
      Splits(0)._ColumnProps(329)=   "Column(82).Width=3281"
      Splits(0)._ColumnProps(330)=   "Column(82).DividerColor=0"
      Splits(0)._ColumnProps(331)=   "Column(82)._WidthInPix=3175"
      Splits(0)._ColumnProps(332)=   "Column(82).Order=83"
      Splits(0)._ColumnProps(333)=   "Column(83).Width=3281"
      Splits(0)._ColumnProps(334)=   "Column(83).DividerColor=0"
      Splits(0)._ColumnProps(335)=   "Column(83)._WidthInPix=3175"
      Splits(0)._ColumnProps(336)=   "Column(83).Order=84"
      Splits(0)._ColumnProps(337)=   "Column(84).Width=3281"
      Splits(0)._ColumnProps(338)=   "Column(84).DividerColor=0"
      Splits(0)._ColumnProps(339)=   "Column(84)._WidthInPix=3175"
      Splits(0)._ColumnProps(340)=   "Column(84).Order=85"
      Splits(0)._ColumnProps(341)=   "Column(85).Width=3281"
      Splits(0)._ColumnProps(342)=   "Column(85).DividerColor=0"
      Splits(0)._ColumnProps(343)=   "Column(85)._WidthInPix=3175"
      Splits(0)._ColumnProps(344)=   "Column(85).Order=86"
      Splits(0)._ColumnProps(345)=   "Column(86).Width=3281"
      Splits(0)._ColumnProps(346)=   "Column(86).DividerColor=0"
      Splits(0)._ColumnProps(347)=   "Column(86)._WidthInPix=3175"
      Splits(0)._ColumnProps(348)=   "Column(86).Order=87"
      Splits(0)._ColumnProps(349)=   "Column(87).Width=3281"
      Splits(0)._ColumnProps(350)=   "Column(87).DividerColor=0"
      Splits(0)._ColumnProps(351)=   "Column(87)._WidthInPix=3175"
      Splits(0)._ColumnProps(352)=   "Column(87).Order=88"
      Splits(0)._ColumnProps(353)=   "Column(88).Width=3281"
      Splits(0)._ColumnProps(354)=   "Column(88).DividerColor=0"
      Splits(0)._ColumnProps(355)=   "Column(88)._WidthInPix=3175"
      Splits(0)._ColumnProps(356)=   "Column(88).Order=89"
      Splits(0)._ColumnProps(357)=   "Column(89).Width=3281"
      Splits(0)._ColumnProps(358)=   "Column(89).DividerColor=0"
      Splits(0)._ColumnProps(359)=   "Column(89)._WidthInPix=3175"
      Splits(0)._ColumnProps(360)=   "Column(89).Order=90"
      Splits(0)._ColumnProps(361)=   "Column(90).Width=3281"
      Splits(0)._ColumnProps(362)=   "Column(90).DividerColor=0"
      Splits(0)._ColumnProps(363)=   "Column(90)._WidthInPix=3175"
      Splits(0)._ColumnProps(364)=   "Column(90).Order=91"
      Splits(0)._ColumnProps(365)=   "Column(91).Width=3281"
      Splits(0)._ColumnProps(366)=   "Column(91).DividerColor=0"
      Splits(0)._ColumnProps(367)=   "Column(91)._WidthInPix=3175"
      Splits(0)._ColumnProps(368)=   "Column(91).Order=92"
      Splits(0)._ColumnProps(369)=   "Column(92).Width=3281"
      Splits(0)._ColumnProps(370)=   "Column(92).DividerColor=0"
      Splits(0)._ColumnProps(371)=   "Column(92)._WidthInPix=3175"
      Splits(0)._ColumnProps(372)=   "Column(92).Order=93"
      Splits(0)._ColumnProps(373)=   "Column(93).Width=3281"
      Splits(0)._ColumnProps(374)=   "Column(93).DividerColor=0"
      Splits(0)._ColumnProps(375)=   "Column(93)._WidthInPix=3175"
      Splits(0)._ColumnProps(376)=   "Column(93).Order=94"
      Splits(0)._ColumnProps(377)=   "Column(94).Width=3281"
      Splits(0)._ColumnProps(378)=   "Column(94).DividerColor=0"
      Splits(0)._ColumnProps(379)=   "Column(94)._WidthInPix=3175"
      Splits(0)._ColumnProps(380)=   "Column(94).Order=95"
      Splits(0)._ColumnProps(381)=   "Column(95).Width=3281"
      Splits(0)._ColumnProps(382)=   "Column(95).DividerColor=0"
      Splits(0)._ColumnProps(383)=   "Column(95)._WidthInPix=3175"
      Splits(0)._ColumnProps(384)=   "Column(95).Order=96"
      Splits(0)._ColumnProps(385)=   "Column(96).Width=3281"
      Splits(0)._ColumnProps(386)=   "Column(96).DividerColor=0"
      Splits(0)._ColumnProps(387)=   "Column(96)._WidthInPix=3175"
      Splits(0)._ColumnProps(388)=   "Column(96).Order=97"
      Splits(0)._ColumnProps(389)=   "Column(97).Width=3281"
      Splits(0)._ColumnProps(390)=   "Column(97).DividerColor=0"
      Splits(0)._ColumnProps(391)=   "Column(97)._WidthInPix=3175"
      Splits(0)._ColumnProps(392)=   "Column(97).Order=98"
      Splits(0)._ColumnProps(393)=   "Column(98).Width=3281"
      Splits(0)._ColumnProps(394)=   "Column(98).DividerColor=0"
      Splits(0)._ColumnProps(395)=   "Column(98)._WidthInPix=3175"
      Splits(0)._ColumnProps(396)=   "Column(98).Order=99"
      Splits(0)._ColumnProps(397)=   "Column(99).Width=3281"
      Splits(0)._ColumnProps(398)=   "Column(99).DividerColor=0"
      Splits(0)._ColumnProps(399)=   "Column(99)._WidthInPix=3175"
      Splits(0)._ColumnProps(400)=   "Column(99).Order=100"
      Splits(0)._ColumnProps(401)=   "Column(100).Width=3281"
      Splits(0)._ColumnProps(402)=   "Column(100).DividerColor=0"
      Splits(0)._ColumnProps(403)=   "Column(100)._WidthInPix=3175"
      Splits(0)._ColumnProps(404)=   "Column(100).Order=101"
      Splits(0)._ColumnProps(405)=   "Column(101).Width=3281"
      Splits(0)._ColumnProps(406)=   "Column(101).DividerColor=0"
      Splits(0)._ColumnProps(407)=   "Column(101)._WidthInPix=3175"
      Splits(0)._ColumnProps(408)=   "Column(101).Order=102"
      Splits(0)._ColumnProps(409)=   "Column(102).Width=3281"
      Splits(0)._ColumnProps(410)=   "Column(102).DividerColor=0"
      Splits(0)._ColumnProps(411)=   "Column(102)._WidthInPix=3175"
      Splits(0)._ColumnProps(412)=   "Column(102).Order=103"
      Splits(0)._ColumnProps(413)=   "Column(103).Width=3281"
      Splits(0)._ColumnProps(414)=   "Column(103).DividerColor=0"
      Splits(0)._ColumnProps(415)=   "Column(103)._WidthInPix=3175"
      Splits(0)._ColumnProps(416)=   "Column(103).Order=104"
      Splits(0)._ColumnProps(417)=   "Column(104).Width=3281"
      Splits(0)._ColumnProps(418)=   "Column(104).DividerColor=0"
      Splits(0)._ColumnProps(419)=   "Column(104)._WidthInPix=3175"
      Splits(0)._ColumnProps(420)=   "Column(104).Order=105"
      Splits(1)._UserFlags=   0
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).AlternatingRowStyle=   -1  'True
      Splits(1).DividerColor=   14933984
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=105"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(1)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(1)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(1)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(9)=   "Column(2).Width=3281"
      Splits(1)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
      Splits(1)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(13)=   "Column(3).Width=3281"
      Splits(1)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
      Splits(1)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(17)=   "Column(4).Width=3281"
      Splits(1)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
      Splits(1)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(21)=   "Column(5).Width=3281"
      Splits(1)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
      Splits(1)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(25)=   "Column(6).Width=3281"
      Splits(1)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
      Splits(1)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(29)=   "Column(7).Width=3281"
      Splits(1)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
      Splits(1)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(33)=   "Column(8).Width=3281"
      Splits(1)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
      Splits(1)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(37)=   "Column(9).Width=3281"
      Splits(1)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
      Splits(1)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(41)=   "Column(10).Width=3281"
      Splits(1)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
      Splits(1)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(45)=   "Column(11).Width=3281"
      Splits(1)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
      Splits(1)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(49)=   "Column(12).Width=3281"
      Splits(1)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
      Splits(1)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(53)=   "Column(13).Width=3281"
      Splits(1)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
      Splits(1)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(57)=   "Column(14).Width=3281"
      Splits(1)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
      Splits(1)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(61)=   "Column(15).Width=3281"
      Splits(1)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
      Splits(1)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(65)=   "Column(16).Width=3281"
      Splits(1)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
      Splits(1)._ColumnProps(68)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(69)=   "Column(17).Width=3281"
      Splits(1)._ColumnProps(70)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(71)=   "Column(17)._WidthInPix=3175"
      Splits(1)._ColumnProps(72)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(73)=   "Column(18).Width=3281"
      Splits(1)._ColumnProps(74)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(75)=   "Column(18)._WidthInPix=3175"
      Splits(1)._ColumnProps(76)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(77)=   "Column(19).Width=3281"
      Splits(1)._ColumnProps(78)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(79)=   "Column(19)._WidthInPix=3175"
      Splits(1)._ColumnProps(80)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(81)=   "Column(20).Width=3281"
      Splits(1)._ColumnProps(82)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(20)._WidthInPix=3175"
      Splits(1)._ColumnProps(84)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(85)=   "Column(21).Width=3281"
      Splits(1)._ColumnProps(86)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(87)=   "Column(21)._WidthInPix=3175"
      Splits(1)._ColumnProps(88)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(89)=   "Column(22).Width=3281"
      Splits(1)._ColumnProps(90)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(91)=   "Column(22)._WidthInPix=3175"
      Splits(1)._ColumnProps(92)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(93)=   "Column(23).Width=3281"
      Splits(1)._ColumnProps(94)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(95)=   "Column(23)._WidthInPix=3175"
      Splits(1)._ColumnProps(96)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(97)=   "Column(24).Width=3281"
      Splits(1)._ColumnProps(98)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(99)=   "Column(24)._WidthInPix=3175"
      Splits(1)._ColumnProps(100)=   "Column(24).Order=25"
      Splits(1)._ColumnProps(101)=   "Column(25).Width=3281"
      Splits(1)._ColumnProps(102)=   "Column(25).DividerColor=0"
      Splits(1)._ColumnProps(103)=   "Column(25)._WidthInPix=3175"
      Splits(1)._ColumnProps(104)=   "Column(25).Order=26"
      Splits(1)._ColumnProps(105)=   "Column(26).Width=3281"
      Splits(1)._ColumnProps(106)=   "Column(26).DividerColor=0"
      Splits(1)._ColumnProps(107)=   "Column(26)._WidthInPix=3175"
      Splits(1)._ColumnProps(108)=   "Column(26).Order=27"
      Splits(1)._ColumnProps(109)=   "Column(27).Width=3281"
      Splits(1)._ColumnProps(110)=   "Column(27).DividerColor=0"
      Splits(1)._ColumnProps(111)=   "Column(27)._WidthInPix=3175"
      Splits(1)._ColumnProps(112)=   "Column(27).Order=28"
      Splits(1)._ColumnProps(113)=   "Column(28).Width=3281"
      Splits(1)._ColumnProps(114)=   "Column(28).DividerColor=0"
      Splits(1)._ColumnProps(115)=   "Column(28)._WidthInPix=3175"
      Splits(1)._ColumnProps(116)=   "Column(28).Order=29"
      Splits(1)._ColumnProps(117)=   "Column(29).Width=3281"
      Splits(1)._ColumnProps(118)=   "Column(29).DividerColor=0"
      Splits(1)._ColumnProps(119)=   "Column(29)._WidthInPix=3175"
      Splits(1)._ColumnProps(120)=   "Column(29).Order=30"
      Splits(1)._ColumnProps(121)=   "Column(30).Width=3281"
      Splits(1)._ColumnProps(122)=   "Column(30).DividerColor=0"
      Splits(1)._ColumnProps(123)=   "Column(30)._WidthInPix=3175"
      Splits(1)._ColumnProps(124)=   "Column(30).Order=31"
      Splits(1)._ColumnProps(125)=   "Column(31).Width=3281"
      Splits(1)._ColumnProps(126)=   "Column(31).DividerColor=0"
      Splits(1)._ColumnProps(127)=   "Column(31)._WidthInPix=3175"
      Splits(1)._ColumnProps(128)=   "Column(31).Order=32"
      Splits(1)._ColumnProps(129)=   "Column(32).Width=3281"
      Splits(1)._ColumnProps(130)=   "Column(32).DividerColor=0"
      Splits(1)._ColumnProps(131)=   "Column(32)._WidthInPix=3175"
      Splits(1)._ColumnProps(132)=   "Column(32).Order=33"
      Splits(1)._ColumnProps(133)=   "Column(33).Width=3281"
      Splits(1)._ColumnProps(134)=   "Column(33).DividerColor=0"
      Splits(1)._ColumnProps(135)=   "Column(33)._WidthInPix=3175"
      Splits(1)._ColumnProps(136)=   "Column(33).Order=34"
      Splits(1)._ColumnProps(137)=   "Column(34).Width=3281"
      Splits(1)._ColumnProps(138)=   "Column(34).DividerColor=0"
      Splits(1)._ColumnProps(139)=   "Column(34)._WidthInPix=3175"
      Splits(1)._ColumnProps(140)=   "Column(34).Order=35"
      Splits(1)._ColumnProps(141)=   "Column(35).Width=3281"
      Splits(1)._ColumnProps(142)=   "Column(35).DividerColor=0"
      Splits(1)._ColumnProps(143)=   "Column(35)._WidthInPix=3175"
      Splits(1)._ColumnProps(144)=   "Column(35).Order=36"
      Splits(1)._ColumnProps(145)=   "Column(36).Width=3281"
      Splits(1)._ColumnProps(146)=   "Column(36).DividerColor=0"
      Splits(1)._ColumnProps(147)=   "Column(36)._WidthInPix=3175"
      Splits(1)._ColumnProps(148)=   "Column(36).Order=37"
      Splits(1)._ColumnProps(149)=   "Column(37).Width=3281"
      Splits(1)._ColumnProps(150)=   "Column(37).DividerColor=0"
      Splits(1)._ColumnProps(151)=   "Column(37)._WidthInPix=3175"
      Splits(1)._ColumnProps(152)=   "Column(37).Order=38"
      Splits(1)._ColumnProps(153)=   "Column(38).Width=3281"
      Splits(1)._ColumnProps(154)=   "Column(38).DividerColor=0"
      Splits(1)._ColumnProps(155)=   "Column(38)._WidthInPix=3175"
      Splits(1)._ColumnProps(156)=   "Column(38).Order=39"
      Splits(1)._ColumnProps(157)=   "Column(39).Width=3281"
      Splits(1)._ColumnProps(158)=   "Column(39).DividerColor=0"
      Splits(1)._ColumnProps(159)=   "Column(39)._WidthInPix=3175"
      Splits(1)._ColumnProps(160)=   "Column(39).Order=40"
      Splits(1)._ColumnProps(161)=   "Column(40).Width=3281"
      Splits(1)._ColumnProps(162)=   "Column(40).DividerColor=0"
      Splits(1)._ColumnProps(163)=   "Column(40)._WidthInPix=3175"
      Splits(1)._ColumnProps(164)=   "Column(40).Order=41"
      Splits(1)._ColumnProps(165)=   "Column(41).Width=3281"
      Splits(1)._ColumnProps(166)=   "Column(41).DividerColor=0"
      Splits(1)._ColumnProps(167)=   "Column(41)._WidthInPix=3175"
      Splits(1)._ColumnProps(168)=   "Column(41).Order=42"
      Splits(1)._ColumnProps(169)=   "Column(42).Width=3281"
      Splits(1)._ColumnProps(170)=   "Column(42).DividerColor=0"
      Splits(1)._ColumnProps(171)=   "Column(42)._WidthInPix=3175"
      Splits(1)._ColumnProps(172)=   "Column(42).Order=43"
      Splits(1)._ColumnProps(173)=   "Column(43).Width=3281"
      Splits(1)._ColumnProps(174)=   "Column(43).DividerColor=0"
      Splits(1)._ColumnProps(175)=   "Column(43)._WidthInPix=3175"
      Splits(1)._ColumnProps(176)=   "Column(43).Order=44"
      Splits(1)._ColumnProps(177)=   "Column(44).Width=3281"
      Splits(1)._ColumnProps(178)=   "Column(44).DividerColor=0"
      Splits(1)._ColumnProps(179)=   "Column(44)._WidthInPix=3175"
      Splits(1)._ColumnProps(180)=   "Column(44).Order=45"
      Splits(1)._ColumnProps(181)=   "Column(45).Width=3281"
      Splits(1)._ColumnProps(182)=   "Column(45).DividerColor=0"
      Splits(1)._ColumnProps(183)=   "Column(45)._WidthInPix=3175"
      Splits(1)._ColumnProps(184)=   "Column(45).Order=46"
      Splits(1)._ColumnProps(185)=   "Column(46).Width=3281"
      Splits(1)._ColumnProps(186)=   "Column(46).DividerColor=0"
      Splits(1)._ColumnProps(187)=   "Column(46)._WidthInPix=3175"
      Splits(1)._ColumnProps(188)=   "Column(46).Order=47"
      Splits(1)._ColumnProps(189)=   "Column(47).Width=3281"
      Splits(1)._ColumnProps(190)=   "Column(47).DividerColor=0"
      Splits(1)._ColumnProps(191)=   "Column(47)._WidthInPix=3175"
      Splits(1)._ColumnProps(192)=   "Column(47).Order=48"
      Splits(1)._ColumnProps(193)=   "Column(48).Width=3281"
      Splits(1)._ColumnProps(194)=   "Column(48).DividerColor=0"
      Splits(1)._ColumnProps(195)=   "Column(48)._WidthInPix=3175"
      Splits(1)._ColumnProps(196)=   "Column(48).Order=49"
      Splits(1)._ColumnProps(197)=   "Column(49).Width=3281"
      Splits(1)._ColumnProps(198)=   "Column(49).DividerColor=0"
      Splits(1)._ColumnProps(199)=   "Column(49)._WidthInPix=3175"
      Splits(1)._ColumnProps(200)=   "Column(49).Order=50"
      Splits(1)._ColumnProps(201)=   "Column(50).Width=3281"
      Splits(1)._ColumnProps(202)=   "Column(50).DividerColor=0"
      Splits(1)._ColumnProps(203)=   "Column(50)._WidthInPix=3175"
      Splits(1)._ColumnProps(204)=   "Column(50).Order=51"
      Splits(1)._ColumnProps(205)=   "Column(51).Width=3281"
      Splits(1)._ColumnProps(206)=   "Column(51).DividerColor=0"
      Splits(1)._ColumnProps(207)=   "Column(51)._WidthInPix=3175"
      Splits(1)._ColumnProps(208)=   "Column(51).Order=52"
      Splits(1)._ColumnProps(209)=   "Column(52).Width=3281"
      Splits(1)._ColumnProps(210)=   "Column(52).DividerColor=0"
      Splits(1)._ColumnProps(211)=   "Column(52)._WidthInPix=3175"
      Splits(1)._ColumnProps(212)=   "Column(52).Order=53"
      Splits(1)._ColumnProps(213)=   "Column(53).Width=3281"
      Splits(1)._ColumnProps(214)=   "Column(53).DividerColor=0"
      Splits(1)._ColumnProps(215)=   "Column(53)._WidthInPix=3175"
      Splits(1)._ColumnProps(216)=   "Column(53).Order=54"
      Splits(1)._ColumnProps(217)=   "Column(54).Width=3281"
      Splits(1)._ColumnProps(218)=   "Column(54).DividerColor=0"
      Splits(1)._ColumnProps(219)=   "Column(54)._WidthInPix=3175"
      Splits(1)._ColumnProps(220)=   "Column(54).Order=55"
      Splits(1)._ColumnProps(221)=   "Column(55).Width=3281"
      Splits(1)._ColumnProps(222)=   "Column(55).DividerColor=0"
      Splits(1)._ColumnProps(223)=   "Column(55)._WidthInPix=3175"
      Splits(1)._ColumnProps(224)=   "Column(55).Order=56"
      Splits(1)._ColumnProps(225)=   "Column(56).Width=3281"
      Splits(1)._ColumnProps(226)=   "Column(56).DividerColor=0"
      Splits(1)._ColumnProps(227)=   "Column(56)._WidthInPix=3175"
      Splits(1)._ColumnProps(228)=   "Column(56).Order=57"
      Splits(1)._ColumnProps(229)=   "Column(57).Width=3281"
      Splits(1)._ColumnProps(230)=   "Column(57).DividerColor=0"
      Splits(1)._ColumnProps(231)=   "Column(57)._WidthInPix=3175"
      Splits(1)._ColumnProps(232)=   "Column(57).Order=58"
      Splits(1)._ColumnProps(233)=   "Column(58).Width=3281"
      Splits(1)._ColumnProps(234)=   "Column(58).DividerColor=0"
      Splits(1)._ColumnProps(235)=   "Column(58)._WidthInPix=3175"
      Splits(1)._ColumnProps(236)=   "Column(58).Order=59"
      Splits(1)._ColumnProps(237)=   "Column(59).Width=3281"
      Splits(1)._ColumnProps(238)=   "Column(59).DividerColor=0"
      Splits(1)._ColumnProps(239)=   "Column(59)._WidthInPix=3175"
      Splits(1)._ColumnProps(240)=   "Column(59).Order=60"
      Splits(1)._ColumnProps(241)=   "Column(60).Width=3281"
      Splits(1)._ColumnProps(242)=   "Column(60).DividerColor=0"
      Splits(1)._ColumnProps(243)=   "Column(60)._WidthInPix=3175"
      Splits(1)._ColumnProps(244)=   "Column(60).Order=61"
      Splits(1)._ColumnProps(245)=   "Column(61).Width=3281"
      Splits(1)._ColumnProps(246)=   "Column(61).DividerColor=0"
      Splits(1)._ColumnProps(247)=   "Column(61)._WidthInPix=3175"
      Splits(1)._ColumnProps(248)=   "Column(61).Order=62"
      Splits(1)._ColumnProps(249)=   "Column(62).Width=3281"
      Splits(1)._ColumnProps(250)=   "Column(62).DividerColor=0"
      Splits(1)._ColumnProps(251)=   "Column(62)._WidthInPix=3175"
      Splits(1)._ColumnProps(252)=   "Column(62).Order=63"
      Splits(1)._ColumnProps(253)=   "Column(63).Width=3281"
      Splits(1)._ColumnProps(254)=   "Column(63).DividerColor=0"
      Splits(1)._ColumnProps(255)=   "Column(63)._WidthInPix=3175"
      Splits(1)._ColumnProps(256)=   "Column(63).Order=64"
      Splits(1)._ColumnProps(257)=   "Column(64).Width=3281"
      Splits(1)._ColumnProps(258)=   "Column(64).DividerColor=0"
      Splits(1)._ColumnProps(259)=   "Column(64)._WidthInPix=3175"
      Splits(1)._ColumnProps(260)=   "Column(64).Order=65"
      Splits(1)._ColumnProps(261)=   "Column(65).Width=3281"
      Splits(1)._ColumnProps(262)=   "Column(65).DividerColor=0"
      Splits(1)._ColumnProps(263)=   "Column(65)._WidthInPix=3175"
      Splits(1)._ColumnProps(264)=   "Column(65).Order=66"
      Splits(1)._ColumnProps(265)=   "Column(66).Width=3281"
      Splits(1)._ColumnProps(266)=   "Column(66).DividerColor=0"
      Splits(1)._ColumnProps(267)=   "Column(66)._WidthInPix=3175"
      Splits(1)._ColumnProps(268)=   "Column(66).Order=67"
      Splits(1)._ColumnProps(269)=   "Column(67).Width=3281"
      Splits(1)._ColumnProps(270)=   "Column(67).DividerColor=0"
      Splits(1)._ColumnProps(271)=   "Column(67)._WidthInPix=3175"
      Splits(1)._ColumnProps(272)=   "Column(67).Order=68"
      Splits(1)._ColumnProps(273)=   "Column(68).Width=3281"
      Splits(1)._ColumnProps(274)=   "Column(68).DividerColor=0"
      Splits(1)._ColumnProps(275)=   "Column(68)._WidthInPix=3175"
      Splits(1)._ColumnProps(276)=   "Column(68).Order=69"
      Splits(1)._ColumnProps(277)=   "Column(69).Width=3281"
      Splits(1)._ColumnProps(278)=   "Column(69).DividerColor=0"
      Splits(1)._ColumnProps(279)=   "Column(69)._WidthInPix=3175"
      Splits(1)._ColumnProps(280)=   "Column(69).Order=70"
      Splits(1)._ColumnProps(281)=   "Column(70).Width=3281"
      Splits(1)._ColumnProps(282)=   "Column(70).DividerColor=0"
      Splits(1)._ColumnProps(283)=   "Column(70)._WidthInPix=3175"
      Splits(1)._ColumnProps(284)=   "Column(70).Order=71"
      Splits(1)._ColumnProps(285)=   "Column(71).Width=3281"
      Splits(1)._ColumnProps(286)=   "Column(71).DividerColor=0"
      Splits(1)._ColumnProps(287)=   "Column(71)._WidthInPix=3175"
      Splits(1)._ColumnProps(288)=   "Column(71).Order=72"
      Splits(1)._ColumnProps(289)=   "Column(72).Width=3281"
      Splits(1)._ColumnProps(290)=   "Column(72).DividerColor=0"
      Splits(1)._ColumnProps(291)=   "Column(72)._WidthInPix=3175"
      Splits(1)._ColumnProps(292)=   "Column(72).Order=73"
      Splits(1)._ColumnProps(293)=   "Column(73).Width=3281"
      Splits(1)._ColumnProps(294)=   "Column(73).DividerColor=0"
      Splits(1)._ColumnProps(295)=   "Column(73)._WidthInPix=3175"
      Splits(1)._ColumnProps(296)=   "Column(73).Order=74"
      Splits(1)._ColumnProps(297)=   "Column(74).Width=3281"
      Splits(1)._ColumnProps(298)=   "Column(74).DividerColor=0"
      Splits(1)._ColumnProps(299)=   "Column(74)._WidthInPix=3175"
      Splits(1)._ColumnProps(300)=   "Column(74).Order=75"
      Splits(1)._ColumnProps(301)=   "Column(75).Width=3281"
      Splits(1)._ColumnProps(302)=   "Column(75).DividerColor=0"
      Splits(1)._ColumnProps(303)=   "Column(75)._WidthInPix=3175"
      Splits(1)._ColumnProps(304)=   "Column(75).Order=76"
      Splits(1)._ColumnProps(305)=   "Column(76).Width=3281"
      Splits(1)._ColumnProps(306)=   "Column(76).DividerColor=0"
      Splits(1)._ColumnProps(307)=   "Column(76)._WidthInPix=3175"
      Splits(1)._ColumnProps(308)=   "Column(76).Order=77"
      Splits(1)._ColumnProps(309)=   "Column(77).Width=3281"
      Splits(1)._ColumnProps(310)=   "Column(77).DividerColor=0"
      Splits(1)._ColumnProps(311)=   "Column(77)._WidthInPix=3175"
      Splits(1)._ColumnProps(312)=   "Column(77).Order=78"
      Splits(1)._ColumnProps(313)=   "Column(78).Width=3281"
      Splits(1)._ColumnProps(314)=   "Column(78).DividerColor=0"
      Splits(1)._ColumnProps(315)=   "Column(78)._WidthInPix=3175"
      Splits(1)._ColumnProps(316)=   "Column(78).Order=79"
      Splits(1)._ColumnProps(317)=   "Column(79).Width=3281"
      Splits(1)._ColumnProps(318)=   "Column(79).DividerColor=0"
      Splits(1)._ColumnProps(319)=   "Column(79)._WidthInPix=3175"
      Splits(1)._ColumnProps(320)=   "Column(79).Order=80"
      Splits(1)._ColumnProps(321)=   "Column(80).Width=3281"
      Splits(1)._ColumnProps(322)=   "Column(80).DividerColor=0"
      Splits(1)._ColumnProps(323)=   "Column(80)._WidthInPix=3175"
      Splits(1)._ColumnProps(324)=   "Column(80).Order=81"
      Splits(1)._ColumnProps(325)=   "Column(81).Width=3281"
      Splits(1)._ColumnProps(326)=   "Column(81).DividerColor=0"
      Splits(1)._ColumnProps(327)=   "Column(81)._WidthInPix=3175"
      Splits(1)._ColumnProps(328)=   "Column(81).Order=82"
      Splits(1)._ColumnProps(329)=   "Column(82).Width=3281"
      Splits(1)._ColumnProps(330)=   "Column(82).DividerColor=0"
      Splits(1)._ColumnProps(331)=   "Column(82)._WidthInPix=3175"
      Splits(1)._ColumnProps(332)=   "Column(82).Order=83"
      Splits(1)._ColumnProps(333)=   "Column(83).Width=3281"
      Splits(1)._ColumnProps(334)=   "Column(83).DividerColor=0"
      Splits(1)._ColumnProps(335)=   "Column(83)._WidthInPix=3175"
      Splits(1)._ColumnProps(336)=   "Column(83).Order=84"
      Splits(1)._ColumnProps(337)=   "Column(84).Width=3281"
      Splits(1)._ColumnProps(338)=   "Column(84).DividerColor=0"
      Splits(1)._ColumnProps(339)=   "Column(84)._WidthInPix=3175"
      Splits(1)._ColumnProps(340)=   "Column(84).Order=85"
      Splits(1)._ColumnProps(341)=   "Column(85).Width=3281"
      Splits(1)._ColumnProps(342)=   "Column(85).DividerColor=0"
      Splits(1)._ColumnProps(343)=   "Column(85)._WidthInPix=3175"
      Splits(1)._ColumnProps(344)=   "Column(85).Order=86"
      Splits(1)._ColumnProps(345)=   "Column(86).Width=3281"
      Splits(1)._ColumnProps(346)=   "Column(86).DividerColor=0"
      Splits(1)._ColumnProps(347)=   "Column(86)._WidthInPix=3175"
      Splits(1)._ColumnProps(348)=   "Column(86).Order=87"
      Splits(1)._ColumnProps(349)=   "Column(87).Width=3281"
      Splits(1)._ColumnProps(350)=   "Column(87).DividerColor=0"
      Splits(1)._ColumnProps(351)=   "Column(87)._WidthInPix=3175"
      Splits(1)._ColumnProps(352)=   "Column(87).Order=88"
      Splits(1)._ColumnProps(353)=   "Column(88).Width=3281"
      Splits(1)._ColumnProps(354)=   "Column(88).DividerColor=0"
      Splits(1)._ColumnProps(355)=   "Column(88)._WidthInPix=3175"
      Splits(1)._ColumnProps(356)=   "Column(88).Order=89"
      Splits(1)._ColumnProps(357)=   "Column(89).Width=3281"
      Splits(1)._ColumnProps(358)=   "Column(89).DividerColor=0"
      Splits(1)._ColumnProps(359)=   "Column(89)._WidthInPix=3175"
      Splits(1)._ColumnProps(360)=   "Column(89).Order=90"
      Splits(1)._ColumnProps(361)=   "Column(90).Width=3281"
      Splits(1)._ColumnProps(362)=   "Column(90).DividerColor=0"
      Splits(1)._ColumnProps(363)=   "Column(90)._WidthInPix=3175"
      Splits(1)._ColumnProps(364)=   "Column(90).Order=91"
      Splits(1)._ColumnProps(365)=   "Column(91).Width=3281"
      Splits(1)._ColumnProps(366)=   "Column(91).DividerColor=0"
      Splits(1)._ColumnProps(367)=   "Column(91)._WidthInPix=3175"
      Splits(1)._ColumnProps(368)=   "Column(91).Order=92"
      Splits(1)._ColumnProps(369)=   "Column(92).Width=3281"
      Splits(1)._ColumnProps(370)=   "Column(92).DividerColor=0"
      Splits(1)._ColumnProps(371)=   "Column(92)._WidthInPix=3175"
      Splits(1)._ColumnProps(372)=   "Column(92).Order=93"
      Splits(1)._ColumnProps(373)=   "Column(93).Width=3281"
      Splits(1)._ColumnProps(374)=   "Column(93).DividerColor=0"
      Splits(1)._ColumnProps(375)=   "Column(93)._WidthInPix=3175"
      Splits(1)._ColumnProps(376)=   "Column(93).Order=94"
      Splits(1)._ColumnProps(377)=   "Column(94).Width=3281"
      Splits(1)._ColumnProps(378)=   "Column(94).DividerColor=0"
      Splits(1)._ColumnProps(379)=   "Column(94)._WidthInPix=3175"
      Splits(1)._ColumnProps(380)=   "Column(94).Order=95"
      Splits(1)._ColumnProps(381)=   "Column(95).Width=3281"
      Splits(1)._ColumnProps(382)=   "Column(95).DividerColor=0"
      Splits(1)._ColumnProps(383)=   "Column(95)._WidthInPix=3175"
      Splits(1)._ColumnProps(384)=   "Column(95).Order=96"
      Splits(1)._ColumnProps(385)=   "Column(96).Width=3281"
      Splits(1)._ColumnProps(386)=   "Column(96).DividerColor=0"
      Splits(1)._ColumnProps(387)=   "Column(96)._WidthInPix=3175"
      Splits(1)._ColumnProps(388)=   "Column(96).Order=97"
      Splits(1)._ColumnProps(389)=   "Column(97).Width=3281"
      Splits(1)._ColumnProps(390)=   "Column(97).DividerColor=0"
      Splits(1)._ColumnProps(391)=   "Column(97)._WidthInPix=3175"
      Splits(1)._ColumnProps(392)=   "Column(97).Order=98"
      Splits(1)._ColumnProps(393)=   "Column(98).Width=3281"
      Splits(1)._ColumnProps(394)=   "Column(98).DividerColor=0"
      Splits(1)._ColumnProps(395)=   "Column(98)._WidthInPix=3175"
      Splits(1)._ColumnProps(396)=   "Column(98).Order=99"
      Splits(1)._ColumnProps(397)=   "Column(99).Width=3281"
      Splits(1)._ColumnProps(398)=   "Column(99).DividerColor=0"
      Splits(1)._ColumnProps(399)=   "Column(99)._WidthInPix=3175"
      Splits(1)._ColumnProps(400)=   "Column(99).Order=100"
      Splits(1)._ColumnProps(401)=   "Column(100).Width=3281"
      Splits(1)._ColumnProps(402)=   "Column(100).DividerColor=0"
      Splits(1)._ColumnProps(403)=   "Column(100)._WidthInPix=3175"
      Splits(1)._ColumnProps(404)=   "Column(100).Order=101"
      Splits(1)._ColumnProps(405)=   "Column(101).Width=3281"
      Splits(1)._ColumnProps(406)=   "Column(101).DividerColor=0"
      Splits(1)._ColumnProps(407)=   "Column(101)._WidthInPix=3175"
      Splits(1)._ColumnProps(408)=   "Column(101).Order=102"
      Splits(1)._ColumnProps(409)=   "Column(102).Width=3281"
      Splits(1)._ColumnProps(410)=   "Column(102).DividerColor=0"
      Splits(1)._ColumnProps(411)=   "Column(102)._WidthInPix=3175"
      Splits(1)._ColumnProps(412)=   "Column(102).Order=103"
      Splits(1)._ColumnProps(413)=   "Column(103).Width=3281"
      Splits(1)._ColumnProps(414)=   "Column(103).DividerColor=0"
      Splits(1)._ColumnProps(415)=   "Column(103)._WidthInPix=3175"
      Splits(1)._ColumnProps(416)=   "Column(103).Order=104"
      Splits(1)._ColumnProps(417)=   "Column(104).Width=3281"
      Splits(1)._ColumnProps(418)=   "Column(104).DividerColor=0"
      Splits(1)._ColumnProps(419)=   "Column(104)._WidthInPix=3175"
      Splits(1)._ColumnProps(420)=   "Column(104).Order=105"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(24)  =   "Splits(0).Style:id=79,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=93,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=80,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=81,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=82,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=84,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=83,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=85,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=86,.parent=9,.bgcolor=&H80FFFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=92,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=94,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=95,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=522,.parent=79"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=519,.parent=80"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=520,.parent=81"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=521,.parent=83"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=509,.parent=79"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=506,.parent=80"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=507,.parent=81"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=508,.parent=83"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=501,.parent=79"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=498,.parent=80"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=499,.parent=81"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=500,.parent=83"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=493,.parent=79"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=490,.parent=80"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=491,.parent=81"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=492,.parent=83"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=485,.parent=79"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=482,.parent=80"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=483,.parent=81"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=484,.parent=83"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=477,.parent=79"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=474,.parent=80"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=475,.parent=81"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=476,.parent=83"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=469,.parent=79"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=466,.parent=80"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=467,.parent=81"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=468,.parent=83"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=461,.parent=79"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=458,.parent=80"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=459,.parent=81"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=460,.parent=83"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=453,.parent=79"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=450,.parent=80"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=451,.parent=81"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=452,.parent=83"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=445,.parent=79"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=442,.parent=80"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=443,.parent=81"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=444,.parent=83"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=437,.parent=79"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=434,.parent=80"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=435,.parent=81"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=436,.parent=83"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=429,.parent=79"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=426,.parent=80"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=427,.parent=81"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=428,.parent=83"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=421,.parent=79"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=418,.parent=80"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=419,.parent=81"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=420,.parent=83"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=413,.parent=79"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=410,.parent=80"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=411,.parent=81"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=412,.parent=83"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=405,.parent=79"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=402,.parent=80"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=403,.parent=81"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=404,.parent=83"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=397,.parent=79"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=394,.parent=80"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=395,.parent=81"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=396,.parent=83"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=389,.parent=79"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=386,.parent=80"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=387,.parent=81"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=388,.parent=83"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=381,.parent=79"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=378,.parent=80"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=379,.parent=81"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=380,.parent=83"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=373,.parent=79"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=370,.parent=80"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=371,.parent=81"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=372,.parent=83"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=365,.parent=79"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=362,.parent=80"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=363,.parent=81"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=364,.parent=83"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=357,.parent=79"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=354,.parent=80"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=355,.parent=81"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=356,.parent=83"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=349,.parent=79"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=346,.parent=80"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=347,.parent=81"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=348,.parent=83"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=341,.parent=79"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=338,.parent=80"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=339,.parent=81"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=340,.parent=83"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=333,.parent=79"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=330,.parent=80"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=331,.parent=81"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=332,.parent=83"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=325,.parent=79"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=322,.parent=80"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=323,.parent=81"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=324,.parent=83"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=317,.parent=79"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=314,.parent=80"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=315,.parent=81"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=316,.parent=83"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=309,.parent=79"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=306,.parent=80"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=307,.parent=81"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=308,.parent=83"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=301,.parent=79"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=298,.parent=80"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=299,.parent=81"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=300,.parent=83"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=293,.parent=79"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=290,.parent=80"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=291,.parent=81"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=292,.parent=83"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=285,.parent=79"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=282,.parent=80"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=283,.parent=81"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=284,.parent=83"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=277,.parent=79"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=274,.parent=80"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=275,.parent=81"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=276,.parent=83"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=269,.parent=79"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=266,.parent=80"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=267,.parent=81"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=268,.parent=83"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=261,.parent=79"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=258,.parent=80"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=259,.parent=81"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=260,.parent=83"
      _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=253,.parent=79"
      _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=250,.parent=80"
      _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=251,.parent=81"
      _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=252,.parent=83"
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=245,.parent=79"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=242,.parent=80"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=243,.parent=81"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=244,.parent=83"
      _StyleDefs(176) =   "Splits(0).Columns(35).Style:id=237,.parent=79"
      _StyleDefs(177) =   "Splits(0).Columns(35).HeadingStyle:id=234,.parent=80"
      _StyleDefs(178) =   "Splits(0).Columns(35).FooterStyle:id=235,.parent=81"
      _StyleDefs(179) =   "Splits(0).Columns(35).EditorStyle:id=236,.parent=83"
      _StyleDefs(180) =   "Splits(0).Columns(36).Style:id=229,.parent=79"
      _StyleDefs(181) =   "Splits(0).Columns(36).HeadingStyle:id=226,.parent=80"
      _StyleDefs(182) =   "Splits(0).Columns(36).FooterStyle:id=227,.parent=81"
      _StyleDefs(183) =   "Splits(0).Columns(36).EditorStyle:id=228,.parent=83"
      _StyleDefs(184) =   "Splits(0).Columns(37).Style:id=221,.parent=79"
      _StyleDefs(185) =   "Splits(0).Columns(37).HeadingStyle:id=218,.parent=80"
      _StyleDefs(186) =   "Splits(0).Columns(37).FooterStyle:id=219,.parent=81"
      _StyleDefs(187) =   "Splits(0).Columns(37).EditorStyle:id=220,.parent=83"
      _StyleDefs(188) =   "Splits(0).Columns(38).Style:id=213,.parent=79"
      _StyleDefs(189) =   "Splits(0).Columns(38).HeadingStyle:id=210,.parent=80"
      _StyleDefs(190) =   "Splits(0).Columns(38).FooterStyle:id=211,.parent=81"
      _StyleDefs(191) =   "Splits(0).Columns(38).EditorStyle:id=212,.parent=83"
      _StyleDefs(192) =   "Splits(0).Columns(39).Style:id=205,.parent=79"
      _StyleDefs(193) =   "Splits(0).Columns(39).HeadingStyle:id=202,.parent=80"
      _StyleDefs(194) =   "Splits(0).Columns(39).FooterStyle:id=203,.parent=81"
      _StyleDefs(195) =   "Splits(0).Columns(39).EditorStyle:id=204,.parent=83"
      _StyleDefs(196) =   "Splits(0).Columns(40).Style:id=197,.parent=79"
      _StyleDefs(197) =   "Splits(0).Columns(40).HeadingStyle:id=194,.parent=80"
      _StyleDefs(198) =   "Splits(0).Columns(40).FooterStyle:id=195,.parent=81"
      _StyleDefs(199) =   "Splits(0).Columns(40).EditorStyle:id=196,.parent=83"
      _StyleDefs(200) =   "Splits(0).Columns(41).Style:id=189,.parent=79"
      _StyleDefs(201) =   "Splits(0).Columns(41).HeadingStyle:id=186,.parent=80"
      _StyleDefs(202) =   "Splits(0).Columns(41).FooterStyle:id=187,.parent=81"
      _StyleDefs(203) =   "Splits(0).Columns(41).EditorStyle:id=188,.parent=83"
      _StyleDefs(204) =   "Splits(0).Columns(42).Style:id=181,.parent=79"
      _StyleDefs(205) =   "Splits(0).Columns(42).HeadingStyle:id=178,.parent=80"
      _StyleDefs(206) =   "Splits(0).Columns(42).FooterStyle:id=179,.parent=81"
      _StyleDefs(207) =   "Splits(0).Columns(42).EditorStyle:id=180,.parent=83"
      _StyleDefs(208) =   "Splits(0).Columns(43).Style:id=173,.parent=79"
      _StyleDefs(209) =   "Splits(0).Columns(43).HeadingStyle:id=170,.parent=80"
      _StyleDefs(210) =   "Splits(0).Columns(43).FooterStyle:id=171,.parent=81"
      _StyleDefs(211) =   "Splits(0).Columns(43).EditorStyle:id=172,.parent=83"
      _StyleDefs(212) =   "Splits(0).Columns(44).Style:id=165,.parent=79"
      _StyleDefs(213) =   "Splits(0).Columns(44).HeadingStyle:id=162,.parent=80"
      _StyleDefs(214) =   "Splits(0).Columns(44).FooterStyle:id=163,.parent=81"
      _StyleDefs(215) =   "Splits(0).Columns(44).EditorStyle:id=164,.parent=83"
      _StyleDefs(216) =   "Splits(0).Columns(45).Style:id=157,.parent=79"
      _StyleDefs(217) =   "Splits(0).Columns(45).HeadingStyle:id=154,.parent=80"
      _StyleDefs(218) =   "Splits(0).Columns(45).FooterStyle:id=155,.parent=81"
      _StyleDefs(219) =   "Splits(0).Columns(45).EditorStyle:id=156,.parent=83"
      _StyleDefs(220) =   "Splits(0).Columns(46).Style:id=149,.parent=79"
      _StyleDefs(221) =   "Splits(0).Columns(46).HeadingStyle:id=146,.parent=80"
      _StyleDefs(222) =   "Splits(0).Columns(46).FooterStyle:id=147,.parent=81"
      _StyleDefs(223) =   "Splits(0).Columns(46).EditorStyle:id=148,.parent=83"
      _StyleDefs(224) =   "Splits(0).Columns(47).Style:id=99,.parent=79"
      _StyleDefs(225) =   "Splits(0).Columns(47).HeadingStyle:id=96,.parent=80"
      _StyleDefs(226) =   "Splits(0).Columns(47).FooterStyle:id=97,.parent=81"
      _StyleDefs(227) =   "Splits(0).Columns(47).EditorStyle:id=98,.parent=83"
      _StyleDefs(228) =   "Splits(0).Columns(48).Style:id=103,.parent=79"
      _StyleDefs(229) =   "Splits(0).Columns(48).HeadingStyle:id=100,.parent=80"
      _StyleDefs(230) =   "Splits(0).Columns(48).FooterStyle:id=101,.parent=81"
      _StyleDefs(231) =   "Splits(0).Columns(48).EditorStyle:id=102,.parent=83"
      _StyleDefs(232) =   "Splits(0).Columns(49).Style:id=107,.parent=79"
      _StyleDefs(233) =   "Splits(0).Columns(49).HeadingStyle:id=104,.parent=80"
      _StyleDefs(234) =   "Splits(0).Columns(49).FooterStyle:id=105,.parent=81"
      _StyleDefs(235) =   "Splits(0).Columns(49).EditorStyle:id=106,.parent=83"
      _StyleDefs(236) =   "Splits(0).Columns(50).Style:id=111,.parent=79"
      _StyleDefs(237) =   "Splits(0).Columns(50).HeadingStyle:id=108,.parent=80"
      _StyleDefs(238) =   "Splits(0).Columns(50).FooterStyle:id=109,.parent=81"
      _StyleDefs(239) =   "Splits(0).Columns(50).EditorStyle:id=110,.parent=83"
      _StyleDefs(240) =   "Splits(0).Columns(51).Style:id=115,.parent=79"
      _StyleDefs(241) =   "Splits(0).Columns(51).HeadingStyle:id=112,.parent=80"
      _StyleDefs(242) =   "Splits(0).Columns(51).FooterStyle:id=113,.parent=81"
      _StyleDefs(243) =   "Splits(0).Columns(51).EditorStyle:id=114,.parent=83"
      _StyleDefs(244) =   "Splits(0).Columns(52).Style:id=119,.parent=79"
      _StyleDefs(245) =   "Splits(0).Columns(52).HeadingStyle:id=116,.parent=80"
      _StyleDefs(246) =   "Splits(0).Columns(52).FooterStyle:id=117,.parent=81"
      _StyleDefs(247) =   "Splits(0).Columns(52).EditorStyle:id=118,.parent=83"
      _StyleDefs(248) =   "Splits(0).Columns(53).Style:id=123,.parent=79"
      _StyleDefs(249) =   "Splits(0).Columns(53).HeadingStyle:id=120,.parent=80"
      _StyleDefs(250) =   "Splits(0).Columns(53).FooterStyle:id=121,.parent=81"
      _StyleDefs(251) =   "Splits(0).Columns(53).EditorStyle:id=122,.parent=83"
      _StyleDefs(252) =   "Splits(0).Columns(54).Style:id=127,.parent=79"
      _StyleDefs(253) =   "Splits(0).Columns(54).HeadingStyle:id=124,.parent=80"
      _StyleDefs(254) =   "Splits(0).Columns(54).FooterStyle:id=125,.parent=81"
      _StyleDefs(255) =   "Splits(0).Columns(54).EditorStyle:id=126,.parent=83"
      _StyleDefs(256) =   "Splits(0).Columns(55).Style:id=133,.parent=79"
      _StyleDefs(257) =   "Splits(0).Columns(55).HeadingStyle:id=128,.parent=80"
      _StyleDefs(258) =   "Splits(0).Columns(55).FooterStyle:id=129,.parent=81"
      _StyleDefs(259) =   "Splits(0).Columns(55).EditorStyle:id=130,.parent=83"
      _StyleDefs(260) =   "Splits(0).Columns(56).Style:id=137,.parent=79"
      _StyleDefs(261) =   "Splits(0).Columns(56).HeadingStyle:id=134,.parent=80"
      _StyleDefs(262) =   "Splits(0).Columns(56).FooterStyle:id=135,.parent=81"
      _StyleDefs(263) =   "Splits(0).Columns(56).EditorStyle:id=136,.parent=83"
      _StyleDefs(264) =   "Splits(0).Columns(57).Style:id=141,.parent=79"
      _StyleDefs(265) =   "Splits(0).Columns(57).HeadingStyle:id=138,.parent=80"
      _StyleDefs(266) =   "Splits(0).Columns(57).FooterStyle:id=139,.parent=81"
      _StyleDefs(267) =   "Splits(0).Columns(57).EditorStyle:id=140,.parent=83"
      _StyleDefs(268) =   "Splits(0).Columns(58).Style:id=530,.parent=79"
      _StyleDefs(269) =   "Splits(0).Columns(58).HeadingStyle:id=527,.parent=80"
      _StyleDefs(270) =   "Splits(0).Columns(58).FooterStyle:id=528,.parent=81"
      _StyleDefs(271) =   "Splits(0).Columns(58).EditorStyle:id=529,.parent=83"
      _StyleDefs(272) =   "Splits(0).Columns(59).Style:id=145,.parent=79"
      _StyleDefs(273) =   "Splits(0).Columns(59).HeadingStyle:id=142,.parent=80"
      _StyleDefs(274) =   "Splits(0).Columns(59).FooterStyle:id=143,.parent=81"
      _StyleDefs(275) =   "Splits(0).Columns(59).EditorStyle:id=144,.parent=83"
      _StyleDefs(276) =   "Splits(0).Columns(60).Style:id=538,.parent=79"
      _StyleDefs(277) =   "Splits(0).Columns(60).HeadingStyle:id=535,.parent=80"
      _StyleDefs(278) =   "Splits(0).Columns(60).FooterStyle:id=536,.parent=81"
      _StyleDefs(279) =   "Splits(0).Columns(60).EditorStyle:id=537,.parent=83"
      _StyleDefs(280) =   "Splits(0).Columns(61).Style:id=546,.parent=79"
      _StyleDefs(281) =   "Splits(0).Columns(61).HeadingStyle:id=543,.parent=80"
      _StyleDefs(282) =   "Splits(0).Columns(61).FooterStyle:id=544,.parent=81"
      _StyleDefs(283) =   "Splits(0).Columns(61).EditorStyle:id=545,.parent=83"
      _StyleDefs(284) =   "Splits(0).Columns(62).Style:id=554,.parent=79"
      _StyleDefs(285) =   "Splits(0).Columns(62).HeadingStyle:id=551,.parent=80"
      _StyleDefs(286) =   "Splits(0).Columns(62).FooterStyle:id=552,.parent=81"
      _StyleDefs(287) =   "Splits(0).Columns(62).EditorStyle:id=553,.parent=83"
      _StyleDefs(288) =   "Splits(0).Columns(63).Style:id=562,.parent=79"
      _StyleDefs(289) =   "Splits(0).Columns(63).HeadingStyle:id=559,.parent=80"
      _StyleDefs(290) =   "Splits(0).Columns(63).FooterStyle:id=560,.parent=81"
      _StyleDefs(291) =   "Splits(0).Columns(63).EditorStyle:id=561,.parent=83"
      _StyleDefs(292) =   "Splits(0).Columns(64).Style:id=570,.parent=79"
      _StyleDefs(293) =   "Splits(0).Columns(64).HeadingStyle:id=567,.parent=80"
      _StyleDefs(294) =   "Splits(0).Columns(64).FooterStyle:id=568,.parent=81"
      _StyleDefs(295) =   "Splits(0).Columns(64).EditorStyle:id=569,.parent=83"
      _StyleDefs(296) =   "Splits(0).Columns(65).Style:id=578,.parent=79"
      _StyleDefs(297) =   "Splits(0).Columns(65).HeadingStyle:id=575,.parent=80"
      _StyleDefs(298) =   "Splits(0).Columns(65).FooterStyle:id=576,.parent=81"
      _StyleDefs(299) =   "Splits(0).Columns(65).EditorStyle:id=577,.parent=83"
      _StyleDefs(300) =   "Splits(0).Columns(66).Style:id=586,.parent=79"
      _StyleDefs(301) =   "Splits(0).Columns(66).HeadingStyle:id=583,.parent=80"
      _StyleDefs(302) =   "Splits(0).Columns(66).FooterStyle:id=584,.parent=81"
      _StyleDefs(303) =   "Splits(0).Columns(66).EditorStyle:id=585,.parent=83"
      _StyleDefs(304) =   "Splits(0).Columns(67).Style:id=594,.parent=79"
      _StyleDefs(305) =   "Splits(0).Columns(67).HeadingStyle:id=591,.parent=80"
      _StyleDefs(306) =   "Splits(0).Columns(67).FooterStyle:id=592,.parent=81"
      _StyleDefs(307) =   "Splits(0).Columns(67).EditorStyle:id=593,.parent=83"
      _StyleDefs(308) =   "Splits(0).Columns(68).Style:id=602,.parent=79"
      _StyleDefs(309) =   "Splits(0).Columns(68).HeadingStyle:id=599,.parent=80"
      _StyleDefs(310) =   "Splits(0).Columns(68).FooterStyle:id=600,.parent=81"
      _StyleDefs(311) =   "Splits(0).Columns(68).EditorStyle:id=601,.parent=83"
      _StyleDefs(312) =   "Splits(0).Columns(69).Style:id=610,.parent=79"
      _StyleDefs(313) =   "Splits(0).Columns(69).HeadingStyle:id=607,.parent=80"
      _StyleDefs(314) =   "Splits(0).Columns(69).FooterStyle:id=608,.parent=81"
      _StyleDefs(315) =   "Splits(0).Columns(69).EditorStyle:id=609,.parent=83"
      _StyleDefs(316) =   "Splits(0).Columns(70).Style:id=618,.parent=79"
      _StyleDefs(317) =   "Splits(0).Columns(70).HeadingStyle:id=615,.parent=80"
      _StyleDefs(318) =   "Splits(0).Columns(70).FooterStyle:id=616,.parent=81"
      _StyleDefs(319) =   "Splits(0).Columns(70).EditorStyle:id=617,.parent=83"
      _StyleDefs(320) =   "Splits(0).Columns(71).Style:id=626,.parent=79"
      _StyleDefs(321) =   "Splits(0).Columns(71).HeadingStyle:id=623,.parent=80"
      _StyleDefs(322) =   "Splits(0).Columns(71).FooterStyle:id=624,.parent=81"
      _StyleDefs(323) =   "Splits(0).Columns(71).EditorStyle:id=625,.parent=83"
      _StyleDefs(324) =   "Splits(0).Columns(72).Style:id=634,.parent=79"
      _StyleDefs(325) =   "Splits(0).Columns(72).HeadingStyle:id=631,.parent=80"
      _StyleDefs(326) =   "Splits(0).Columns(72).FooterStyle:id=632,.parent=81"
      _StyleDefs(327) =   "Splits(0).Columns(72).EditorStyle:id=633,.parent=83"
      _StyleDefs(328) =   "Splits(0).Columns(73).Style:id=642,.parent=79"
      _StyleDefs(329) =   "Splits(0).Columns(73).HeadingStyle:id=639,.parent=80"
      _StyleDefs(330) =   "Splits(0).Columns(73).FooterStyle:id=640,.parent=81"
      _StyleDefs(331) =   "Splits(0).Columns(73).EditorStyle:id=641,.parent=83"
      _StyleDefs(332) =   "Splits(0).Columns(74).Style:id=650,.parent=79"
      _StyleDefs(333) =   "Splits(0).Columns(74).HeadingStyle:id=647,.parent=80"
      _StyleDefs(334) =   "Splits(0).Columns(74).FooterStyle:id=648,.parent=81"
      _StyleDefs(335) =   "Splits(0).Columns(74).EditorStyle:id=649,.parent=83"
      _StyleDefs(336) =   "Splits(0).Columns(75).Style:id=658,.parent=79"
      _StyleDefs(337) =   "Splits(0).Columns(75).HeadingStyle:id=655,.parent=80"
      _StyleDefs(338) =   "Splits(0).Columns(75).FooterStyle:id=656,.parent=81"
      _StyleDefs(339) =   "Splits(0).Columns(75).EditorStyle:id=657,.parent=83"
      _StyleDefs(340) =   "Splits(0).Columns(76).Style:id=666,.parent=79"
      _StyleDefs(341) =   "Splits(0).Columns(76).HeadingStyle:id=663,.parent=80"
      _StyleDefs(342) =   "Splits(0).Columns(76).FooterStyle:id=664,.parent=81"
      _StyleDefs(343) =   "Splits(0).Columns(76).EditorStyle:id=665,.parent=83"
      _StyleDefs(344) =   "Splits(0).Columns(77).Style:id=674,.parent=79"
      _StyleDefs(345) =   "Splits(0).Columns(77).HeadingStyle:id=671,.parent=80"
      _StyleDefs(346) =   "Splits(0).Columns(77).FooterStyle:id=672,.parent=81"
      _StyleDefs(347) =   "Splits(0).Columns(77).EditorStyle:id=673,.parent=83"
      _StyleDefs(348) =   "Splits(0).Columns(78).Style:id=682,.parent=79"
      _StyleDefs(349) =   "Splits(0).Columns(78).HeadingStyle:id=679,.parent=80"
      _StyleDefs(350) =   "Splits(0).Columns(78).FooterStyle:id=680,.parent=81"
      _StyleDefs(351) =   "Splits(0).Columns(78).EditorStyle:id=681,.parent=83"
      _StyleDefs(352) =   "Splits(0).Columns(79).Style:id=690,.parent=79"
      _StyleDefs(353) =   "Splits(0).Columns(79).HeadingStyle:id=687,.parent=80"
      _StyleDefs(354) =   "Splits(0).Columns(79).FooterStyle:id=688,.parent=81"
      _StyleDefs(355) =   "Splits(0).Columns(79).EditorStyle:id=689,.parent=83"
      _StyleDefs(356) =   "Splits(0).Columns(80).Style:id=698,.parent=79"
      _StyleDefs(357) =   "Splits(0).Columns(80).HeadingStyle:id=695,.parent=80"
      _StyleDefs(358) =   "Splits(0).Columns(80).FooterStyle:id=696,.parent=81"
      _StyleDefs(359) =   "Splits(0).Columns(80).EditorStyle:id=697,.parent=83"
      _StyleDefs(360) =   "Splits(0).Columns(81).Style:id=706,.parent=79"
      _StyleDefs(361) =   "Splits(0).Columns(81).HeadingStyle:id=703,.parent=80"
      _StyleDefs(362) =   "Splits(0).Columns(81).FooterStyle:id=704,.parent=81"
      _StyleDefs(363) =   "Splits(0).Columns(81).EditorStyle:id=705,.parent=83"
      _StyleDefs(364) =   "Splits(0).Columns(82).Style:id=714,.parent=79"
      _StyleDefs(365) =   "Splits(0).Columns(82).HeadingStyle:id=711,.parent=80"
      _StyleDefs(366) =   "Splits(0).Columns(82).FooterStyle:id=712,.parent=81"
      _StyleDefs(367) =   "Splits(0).Columns(82).EditorStyle:id=713,.parent=83"
      _StyleDefs(368) =   "Splits(0).Columns(83).Style:id=722,.parent=79"
      _StyleDefs(369) =   "Splits(0).Columns(83).HeadingStyle:id=719,.parent=80"
      _StyleDefs(370) =   "Splits(0).Columns(83).FooterStyle:id=720,.parent=81"
      _StyleDefs(371) =   "Splits(0).Columns(83).EditorStyle:id=721,.parent=83"
      _StyleDefs(372) =   "Splits(0).Columns(84).Style:id=730,.parent=79"
      _StyleDefs(373) =   "Splits(0).Columns(84).HeadingStyle:id=727,.parent=80"
      _StyleDefs(374) =   "Splits(0).Columns(84).FooterStyle:id=728,.parent=81"
      _StyleDefs(375) =   "Splits(0).Columns(84).EditorStyle:id=729,.parent=83"
      _StyleDefs(376) =   "Splits(0).Columns(85).Style:id=738,.parent=79"
      _StyleDefs(377) =   "Splits(0).Columns(85).HeadingStyle:id=735,.parent=80"
      _StyleDefs(378) =   "Splits(0).Columns(85).FooterStyle:id=736,.parent=81"
      _StyleDefs(379) =   "Splits(0).Columns(85).EditorStyle:id=737,.parent=83"
      _StyleDefs(380) =   "Splits(0).Columns(86).Style:id=746,.parent=79"
      _StyleDefs(381) =   "Splits(0).Columns(86).HeadingStyle:id=743,.parent=80"
      _StyleDefs(382) =   "Splits(0).Columns(86).FooterStyle:id=744,.parent=81"
      _StyleDefs(383) =   "Splits(0).Columns(86).EditorStyle:id=745,.parent=83"
      _StyleDefs(384) =   "Splits(0).Columns(87).Style:id=754,.parent=79"
      _StyleDefs(385) =   "Splits(0).Columns(87).HeadingStyle:id=751,.parent=80"
      _StyleDefs(386) =   "Splits(0).Columns(87).FooterStyle:id=752,.parent=81"
      _StyleDefs(387) =   "Splits(0).Columns(87).EditorStyle:id=753,.parent=83"
      _StyleDefs(388) =   "Splits(0).Columns(88).Style:id=762,.parent=79"
      _StyleDefs(389) =   "Splits(0).Columns(88).HeadingStyle:id=759,.parent=80"
      _StyleDefs(390) =   "Splits(0).Columns(88).FooterStyle:id=760,.parent=81"
      _StyleDefs(391) =   "Splits(0).Columns(88).EditorStyle:id=761,.parent=83"
      _StyleDefs(392) =   "Splits(0).Columns(89).Style:id=770,.parent=79"
      _StyleDefs(393) =   "Splits(0).Columns(89).HeadingStyle:id=767,.parent=80"
      _StyleDefs(394) =   "Splits(0).Columns(89).FooterStyle:id=768,.parent=81"
      _StyleDefs(395) =   "Splits(0).Columns(89).EditorStyle:id=769,.parent=83"
      _StyleDefs(396) =   "Splits(0).Columns(90).Style:id=778,.parent=79"
      _StyleDefs(397) =   "Splits(0).Columns(90).HeadingStyle:id=775,.parent=80"
      _StyleDefs(398) =   "Splits(0).Columns(90).FooterStyle:id=776,.parent=81"
      _StyleDefs(399) =   "Splits(0).Columns(90).EditorStyle:id=777,.parent=83"
      _StyleDefs(400) =   "Splits(0).Columns(91).Style:id=786,.parent=79"
      _StyleDefs(401) =   "Splits(0).Columns(91).HeadingStyle:id=783,.parent=80"
      _StyleDefs(402) =   "Splits(0).Columns(91).FooterStyle:id=784,.parent=81"
      _StyleDefs(403) =   "Splits(0).Columns(91).EditorStyle:id=785,.parent=83"
      _StyleDefs(404) =   "Splits(0).Columns(92).Style:id=794,.parent=79"
      _StyleDefs(405) =   "Splits(0).Columns(92).HeadingStyle:id=791,.parent=80"
      _StyleDefs(406) =   "Splits(0).Columns(92).FooterStyle:id=792,.parent=81"
      _StyleDefs(407) =   "Splits(0).Columns(92).EditorStyle:id=793,.parent=83"
      _StyleDefs(408) =   "Splits(0).Columns(93).Style:id=802,.parent=79"
      _StyleDefs(409) =   "Splits(0).Columns(93).HeadingStyle:id=799,.parent=80"
      _StyleDefs(410) =   "Splits(0).Columns(93).FooterStyle:id=800,.parent=81"
      _StyleDefs(411) =   "Splits(0).Columns(93).EditorStyle:id=801,.parent=83"
      _StyleDefs(412) =   "Splits(0).Columns(94).Style:id=810,.parent=79"
      _StyleDefs(413) =   "Splits(0).Columns(94).HeadingStyle:id=807,.parent=80"
      _StyleDefs(414) =   "Splits(0).Columns(94).FooterStyle:id=808,.parent=81"
      _StyleDefs(415) =   "Splits(0).Columns(94).EditorStyle:id=809,.parent=83"
      _StyleDefs(416) =   "Splits(0).Columns(95).Style:id=818,.parent=79"
      _StyleDefs(417) =   "Splits(0).Columns(95).HeadingStyle:id=815,.parent=80"
      _StyleDefs(418) =   "Splits(0).Columns(95).FooterStyle:id=816,.parent=81"
      _StyleDefs(419) =   "Splits(0).Columns(95).EditorStyle:id=817,.parent=83"
      _StyleDefs(420) =   "Splits(0).Columns(96).Style:id=826,.parent=79"
      _StyleDefs(421) =   "Splits(0).Columns(96).HeadingStyle:id=823,.parent=80"
      _StyleDefs(422) =   "Splits(0).Columns(96).FooterStyle:id=824,.parent=81"
      _StyleDefs(423) =   "Splits(0).Columns(96).EditorStyle:id=825,.parent=83"
      _StyleDefs(424) =   "Splits(0).Columns(97).Style:id=834,.parent=79"
      _StyleDefs(425) =   "Splits(0).Columns(97).HeadingStyle:id=831,.parent=80"
      _StyleDefs(426) =   "Splits(0).Columns(97).FooterStyle:id=832,.parent=81"
      _StyleDefs(427) =   "Splits(0).Columns(97).EditorStyle:id=833,.parent=83"
      _StyleDefs(428) =   "Splits(0).Columns(98).Style:id=842,.parent=79"
      _StyleDefs(429) =   "Splits(0).Columns(98).HeadingStyle:id=839,.parent=80"
      _StyleDefs(430) =   "Splits(0).Columns(98).FooterStyle:id=840,.parent=81"
      _StyleDefs(431) =   "Splits(0).Columns(98).EditorStyle:id=841,.parent=83"
      _StyleDefs(432) =   "Splits(0).Columns(99).Style:id=850,.parent=79"
      _StyleDefs(433) =   "Splits(0).Columns(99).HeadingStyle:id=847,.parent=80"
      _StyleDefs(434) =   "Splits(0).Columns(99).FooterStyle:id=848,.parent=81"
      _StyleDefs(435) =   "Splits(0).Columns(99).EditorStyle:id=849,.parent=83"
      _StyleDefs(436) =   "Splits(0).Columns(100).Style:id=858,.parent=79"
      _StyleDefs(437) =   "Splits(0).Columns(100).HeadingStyle:id=855,.parent=80"
      _StyleDefs(438) =   "Splits(0).Columns(100).FooterStyle:id=856,.parent=81"
      _StyleDefs(439) =   "Splits(0).Columns(100).EditorStyle:id=857,.parent=83"
      _StyleDefs(440) =   "Splits(0).Columns(101).Style:id=866,.parent=79"
      _StyleDefs(441) =   "Splits(0).Columns(101).HeadingStyle:id=863,.parent=80"
      _StyleDefs(442) =   "Splits(0).Columns(101).FooterStyle:id=864,.parent=81"
      _StyleDefs(443) =   "Splits(0).Columns(101).EditorStyle:id=865,.parent=83"
      _StyleDefs(444) =   "Splits(0).Columns(102).Style:id=874,.parent=79"
      _StyleDefs(445) =   "Splits(0).Columns(102).HeadingStyle:id=871,.parent=80"
      _StyleDefs(446) =   "Splits(0).Columns(102).FooterStyle:id=872,.parent=81"
      _StyleDefs(447) =   "Splits(0).Columns(102).EditorStyle:id=873,.parent=83"
      _StyleDefs(448) =   "Splits(0).Columns(103).Style:id=882,.parent=79"
      _StyleDefs(449) =   "Splits(0).Columns(103).HeadingStyle:id=879,.parent=80"
      _StyleDefs(450) =   "Splits(0).Columns(103).FooterStyle:id=880,.parent=81"
      _StyleDefs(451) =   "Splits(0).Columns(103).EditorStyle:id=881,.parent=83"
      _StyleDefs(452) =   "Splits(0).Columns(104).Style:id=890,.parent=79"
      _StyleDefs(453) =   "Splits(0).Columns(104).HeadingStyle:id=887,.parent=80"
      _StyleDefs(454) =   "Splits(0).Columns(104).FooterStyle:id=888,.parent=81"
      _StyleDefs(455) =   "Splits(0).Columns(104).EditorStyle:id=889,.parent=83"
      _StyleDefs(456) =   "Splits(1).Style:id=13,.parent=1"
      _StyleDefs(457) =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(458) =   "Splits(1).HeadingStyle:id=14,.parent=2"
      _StyleDefs(459) =   "Splits(1).FooterStyle:id=15,.parent=3"
      _StyleDefs(460) =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(461) =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(462) =   "Splits(1).EditorStyle:id=17,.parent=7"
      _StyleDefs(463) =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(464) =   "Splits(1).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
      _StyleDefs(465) =   "Splits(1).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(466) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(467) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(468) =   "Splits(1).Columns(0).Style:id=526,.parent=13"
      _StyleDefs(469) =   "Splits(1).Columns(0).HeadingStyle:id=523,.parent=14"
      _StyleDefs(470) =   "Splits(1).Columns(0).FooterStyle:id=524,.parent=15"
      _StyleDefs(471) =   "Splits(1).Columns(0).EditorStyle:id=525,.parent=17"
      _StyleDefs(472) =   "Splits(1).Columns(1).Style:id=513,.parent=13"
      _StyleDefs(473) =   "Splits(1).Columns(1).HeadingStyle:id=510,.parent=14"
      _StyleDefs(474) =   "Splits(1).Columns(1).FooterStyle:id=511,.parent=15"
      _StyleDefs(475) =   "Splits(1).Columns(1).EditorStyle:id=512,.parent=17"
      _StyleDefs(476) =   "Splits(1).Columns(2).Style:id=505,.parent=13"
      _StyleDefs(477) =   "Splits(1).Columns(2).HeadingStyle:id=502,.parent=14"
      _StyleDefs(478) =   "Splits(1).Columns(2).FooterStyle:id=503,.parent=15"
      _StyleDefs(479) =   "Splits(1).Columns(2).EditorStyle:id=504,.parent=17"
      _StyleDefs(480) =   "Splits(1).Columns(3).Style:id=497,.parent=13"
      _StyleDefs(481) =   "Splits(1).Columns(3).HeadingStyle:id=494,.parent=14"
      _StyleDefs(482) =   "Splits(1).Columns(3).FooterStyle:id=495,.parent=15"
      _StyleDefs(483) =   "Splits(1).Columns(3).EditorStyle:id=496,.parent=17"
      _StyleDefs(484) =   "Splits(1).Columns(4).Style:id=489,.parent=13"
      _StyleDefs(485) =   "Splits(1).Columns(4).HeadingStyle:id=486,.parent=14"
      _StyleDefs(486) =   "Splits(1).Columns(4).FooterStyle:id=487,.parent=15"
      _StyleDefs(487) =   "Splits(1).Columns(4).EditorStyle:id=488,.parent=17"
      _StyleDefs(488) =   "Splits(1).Columns(5).Style:id=481,.parent=13"
      _StyleDefs(489) =   "Splits(1).Columns(5).HeadingStyle:id=478,.parent=14"
      _StyleDefs(490) =   "Splits(1).Columns(5).FooterStyle:id=479,.parent=15"
      _StyleDefs(491) =   "Splits(1).Columns(5).EditorStyle:id=480,.parent=17"
      _StyleDefs(492) =   "Splits(1).Columns(6).Style:id=473,.parent=13"
      _StyleDefs(493) =   "Splits(1).Columns(6).HeadingStyle:id=470,.parent=14"
      _StyleDefs(494) =   "Splits(1).Columns(6).FooterStyle:id=471,.parent=15"
      _StyleDefs(495) =   "Splits(1).Columns(6).EditorStyle:id=472,.parent=17"
      _StyleDefs(496) =   "Splits(1).Columns(7).Style:id=465,.parent=13"
      _StyleDefs(497) =   "Splits(1).Columns(7).HeadingStyle:id=462,.parent=14"
      _StyleDefs(498) =   "Splits(1).Columns(7).FooterStyle:id=463,.parent=15"
      _StyleDefs(499) =   "Splits(1).Columns(7).EditorStyle:id=464,.parent=17"
      _StyleDefs(500) =   "Splits(1).Columns(8).Style:id=457,.parent=13"
      _StyleDefs(501) =   "Splits(1).Columns(8).HeadingStyle:id=454,.parent=14"
      _StyleDefs(502) =   "Splits(1).Columns(8).FooterStyle:id=455,.parent=15"
      _StyleDefs(503) =   "Splits(1).Columns(8).EditorStyle:id=456,.parent=17"
      _StyleDefs(504) =   "Splits(1).Columns(9).Style:id=449,.parent=13"
      _StyleDefs(505) =   "Splits(1).Columns(9).HeadingStyle:id=446,.parent=14"
      _StyleDefs(506) =   "Splits(1).Columns(9).FooterStyle:id=447,.parent=15"
      _StyleDefs(507) =   "Splits(1).Columns(9).EditorStyle:id=448,.parent=17"
      _StyleDefs(508) =   "Splits(1).Columns(10).Style:id=441,.parent=13"
      _StyleDefs(509) =   "Splits(1).Columns(10).HeadingStyle:id=438,.parent=14"
      _StyleDefs(510) =   "Splits(1).Columns(10).FooterStyle:id=439,.parent=15"
      _StyleDefs(511) =   "Splits(1).Columns(10).EditorStyle:id=440,.parent=17"
      _StyleDefs(512) =   "Splits(1).Columns(11).Style:id=433,.parent=13"
      _StyleDefs(513) =   "Splits(1).Columns(11).HeadingStyle:id=430,.parent=14"
      _StyleDefs(514) =   "Splits(1).Columns(11).FooterStyle:id=431,.parent=15"
      _StyleDefs(515) =   "Splits(1).Columns(11).EditorStyle:id=432,.parent=17"
      _StyleDefs(516) =   "Splits(1).Columns(12).Style:id=425,.parent=13"
      _StyleDefs(517) =   "Splits(1).Columns(12).HeadingStyle:id=422,.parent=14"
      _StyleDefs(518) =   "Splits(1).Columns(12).FooterStyle:id=423,.parent=15"
      _StyleDefs(519) =   "Splits(1).Columns(12).EditorStyle:id=424,.parent=17"
      _StyleDefs(520) =   "Splits(1).Columns(13).Style:id=417,.parent=13"
      _StyleDefs(521) =   "Splits(1).Columns(13).HeadingStyle:id=414,.parent=14"
      _StyleDefs(522) =   "Splits(1).Columns(13).FooterStyle:id=415,.parent=15"
      _StyleDefs(523) =   "Splits(1).Columns(13).EditorStyle:id=416,.parent=17"
      _StyleDefs(524) =   "Splits(1).Columns(14).Style:id=409,.parent=13"
      _StyleDefs(525) =   "Splits(1).Columns(14).HeadingStyle:id=406,.parent=14"
      _StyleDefs(526) =   "Splits(1).Columns(14).FooterStyle:id=407,.parent=15"
      _StyleDefs(527) =   "Splits(1).Columns(14).EditorStyle:id=408,.parent=17"
      _StyleDefs(528) =   "Splits(1).Columns(15).Style:id=401,.parent=13"
      _StyleDefs(529) =   "Splits(1).Columns(15).HeadingStyle:id=398,.parent=14"
      _StyleDefs(530) =   "Splits(1).Columns(15).FooterStyle:id=399,.parent=15"
      _StyleDefs(531) =   "Splits(1).Columns(15).EditorStyle:id=400,.parent=17"
      _StyleDefs(532) =   "Splits(1).Columns(16).Style:id=393,.parent=13"
      _StyleDefs(533) =   "Splits(1).Columns(16).HeadingStyle:id=390,.parent=14"
      _StyleDefs(534) =   "Splits(1).Columns(16).FooterStyle:id=391,.parent=15"
      _StyleDefs(535) =   "Splits(1).Columns(16).EditorStyle:id=392,.parent=17"
      _StyleDefs(536) =   "Splits(1).Columns(17).Style:id=385,.parent=13"
      _StyleDefs(537) =   "Splits(1).Columns(17).HeadingStyle:id=382,.parent=14"
      _StyleDefs(538) =   "Splits(1).Columns(17).FooterStyle:id=383,.parent=15"
      _StyleDefs(539) =   "Splits(1).Columns(17).EditorStyle:id=384,.parent=17"
      _StyleDefs(540) =   "Splits(1).Columns(18).Style:id=377,.parent=13"
      _StyleDefs(541) =   "Splits(1).Columns(18).HeadingStyle:id=374,.parent=14"
      _StyleDefs(542) =   "Splits(1).Columns(18).FooterStyle:id=375,.parent=15"
      _StyleDefs(543) =   "Splits(1).Columns(18).EditorStyle:id=376,.parent=17"
      _StyleDefs(544) =   "Splits(1).Columns(19).Style:id=369,.parent=13"
      _StyleDefs(545) =   "Splits(1).Columns(19).HeadingStyle:id=366,.parent=14"
      _StyleDefs(546) =   "Splits(1).Columns(19).FooterStyle:id=367,.parent=15"
      _StyleDefs(547) =   "Splits(1).Columns(19).EditorStyle:id=368,.parent=17"
      _StyleDefs(548) =   "Splits(1).Columns(20).Style:id=361,.parent=13"
      _StyleDefs(549) =   "Splits(1).Columns(20).HeadingStyle:id=358,.parent=14"
      _StyleDefs(550) =   "Splits(1).Columns(20).FooterStyle:id=359,.parent=15"
      _StyleDefs(551) =   "Splits(1).Columns(20).EditorStyle:id=360,.parent=17"
      _StyleDefs(552) =   "Splits(1).Columns(21).Style:id=353,.parent=13"
      _StyleDefs(553) =   "Splits(1).Columns(21).HeadingStyle:id=350,.parent=14"
      _StyleDefs(554) =   "Splits(1).Columns(21).FooterStyle:id=351,.parent=15"
      _StyleDefs(555) =   "Splits(1).Columns(21).EditorStyle:id=352,.parent=17"
      _StyleDefs(556) =   "Splits(1).Columns(22).Style:id=345,.parent=13"
      _StyleDefs(557) =   "Splits(1).Columns(22).HeadingStyle:id=342,.parent=14"
      _StyleDefs(558) =   "Splits(1).Columns(22).FooterStyle:id=343,.parent=15"
      _StyleDefs(559) =   "Splits(1).Columns(22).EditorStyle:id=344,.parent=17"
      _StyleDefs(560) =   "Splits(1).Columns(23).Style:id=337,.parent=13"
      _StyleDefs(561) =   "Splits(1).Columns(23).HeadingStyle:id=334,.parent=14"
      _StyleDefs(562) =   "Splits(1).Columns(23).FooterStyle:id=335,.parent=15"
      _StyleDefs(563) =   "Splits(1).Columns(23).EditorStyle:id=336,.parent=17"
      _StyleDefs(564) =   "Splits(1).Columns(24).Style:id=329,.parent=13"
      _StyleDefs(565) =   "Splits(1).Columns(24).HeadingStyle:id=326,.parent=14"
      _StyleDefs(566) =   "Splits(1).Columns(24).FooterStyle:id=327,.parent=15"
      _StyleDefs(567) =   "Splits(1).Columns(24).EditorStyle:id=328,.parent=17"
      _StyleDefs(568) =   "Splits(1).Columns(25).Style:id=321,.parent=13"
      _StyleDefs(569) =   "Splits(1).Columns(25).HeadingStyle:id=318,.parent=14"
      _StyleDefs(570) =   "Splits(1).Columns(25).FooterStyle:id=319,.parent=15"
      _StyleDefs(571) =   "Splits(1).Columns(25).EditorStyle:id=320,.parent=17"
      _StyleDefs(572) =   "Splits(1).Columns(26).Style:id=313,.parent=13"
      _StyleDefs(573) =   "Splits(1).Columns(26).HeadingStyle:id=310,.parent=14"
      _StyleDefs(574) =   "Splits(1).Columns(26).FooterStyle:id=311,.parent=15"
      _StyleDefs(575) =   "Splits(1).Columns(26).EditorStyle:id=312,.parent=17"
      _StyleDefs(576) =   "Splits(1).Columns(27).Style:id=305,.parent=13"
      _StyleDefs(577) =   "Splits(1).Columns(27).HeadingStyle:id=302,.parent=14"
      _StyleDefs(578) =   "Splits(1).Columns(27).FooterStyle:id=303,.parent=15"
      _StyleDefs(579) =   "Splits(1).Columns(27).EditorStyle:id=304,.parent=17"
      _StyleDefs(580) =   "Splits(1).Columns(28).Style:id=297,.parent=13"
      _StyleDefs(581) =   "Splits(1).Columns(28).HeadingStyle:id=294,.parent=14"
      _StyleDefs(582) =   "Splits(1).Columns(28).FooterStyle:id=295,.parent=15"
      _StyleDefs(583) =   "Splits(1).Columns(28).EditorStyle:id=296,.parent=17"
      _StyleDefs(584) =   "Splits(1).Columns(29).Style:id=289,.parent=13"
      _StyleDefs(585) =   "Splits(1).Columns(29).HeadingStyle:id=286,.parent=14"
      _StyleDefs(586) =   "Splits(1).Columns(29).FooterStyle:id=287,.parent=15"
      _StyleDefs(587) =   "Splits(1).Columns(29).EditorStyle:id=288,.parent=17"
      _StyleDefs(588) =   "Splits(1).Columns(30).Style:id=281,.parent=13"
      _StyleDefs(589) =   "Splits(1).Columns(30).HeadingStyle:id=278,.parent=14"
      _StyleDefs(590) =   "Splits(1).Columns(30).FooterStyle:id=279,.parent=15"
      _StyleDefs(591) =   "Splits(1).Columns(30).EditorStyle:id=280,.parent=17"
      _StyleDefs(592) =   "Splits(1).Columns(31).Style:id=273,.parent=13"
      _StyleDefs(593) =   "Splits(1).Columns(31).HeadingStyle:id=270,.parent=14"
      _StyleDefs(594) =   "Splits(1).Columns(31).FooterStyle:id=271,.parent=15"
      _StyleDefs(595) =   "Splits(1).Columns(31).EditorStyle:id=272,.parent=17"
      _StyleDefs(596) =   "Splits(1).Columns(32).Style:id=265,.parent=13"
      _StyleDefs(597) =   "Splits(1).Columns(32).HeadingStyle:id=262,.parent=14"
      _StyleDefs(598) =   "Splits(1).Columns(32).FooterStyle:id=263,.parent=15"
      _StyleDefs(599) =   "Splits(1).Columns(32).EditorStyle:id=264,.parent=17"
      _StyleDefs(600) =   "Splits(1).Columns(33).Style:id=257,.parent=13"
      _StyleDefs(601) =   "Splits(1).Columns(33).HeadingStyle:id=254,.parent=14"
      _StyleDefs(602) =   "Splits(1).Columns(33).FooterStyle:id=255,.parent=15"
      _StyleDefs(603) =   "Splits(1).Columns(33).EditorStyle:id=256,.parent=17"
      _StyleDefs(604) =   "Splits(1).Columns(34).Style:id=249,.parent=13"
      _StyleDefs(605) =   "Splits(1).Columns(34).HeadingStyle:id=246,.parent=14"
      _StyleDefs(606) =   "Splits(1).Columns(34).FooterStyle:id=247,.parent=15"
      _StyleDefs(607) =   "Splits(1).Columns(34).EditorStyle:id=248,.parent=17"
      _StyleDefs(608) =   "Splits(1).Columns(35).Style:id=241,.parent=13"
      _StyleDefs(609) =   "Splits(1).Columns(35).HeadingStyle:id=238,.parent=14"
      _StyleDefs(610) =   "Splits(1).Columns(35).FooterStyle:id=239,.parent=15"
      _StyleDefs(611) =   "Splits(1).Columns(35).EditorStyle:id=240,.parent=17"
      _StyleDefs(612) =   "Splits(1).Columns(36).Style:id=233,.parent=13"
      _StyleDefs(613) =   "Splits(1).Columns(36).HeadingStyle:id=230,.parent=14"
      _StyleDefs(614) =   "Splits(1).Columns(36).FooterStyle:id=231,.parent=15"
      _StyleDefs(615) =   "Splits(1).Columns(36).EditorStyle:id=232,.parent=17"
      _StyleDefs(616) =   "Splits(1).Columns(37).Style:id=225,.parent=13"
      _StyleDefs(617) =   "Splits(1).Columns(37).HeadingStyle:id=222,.parent=14"
      _StyleDefs(618) =   "Splits(1).Columns(37).FooterStyle:id=223,.parent=15"
      _StyleDefs(619) =   "Splits(1).Columns(37).EditorStyle:id=224,.parent=17"
      _StyleDefs(620) =   "Splits(1).Columns(38).Style:id=217,.parent=13"
      _StyleDefs(621) =   "Splits(1).Columns(38).HeadingStyle:id=214,.parent=14"
      _StyleDefs(622) =   "Splits(1).Columns(38).FooterStyle:id=215,.parent=15"
      _StyleDefs(623) =   "Splits(1).Columns(38).EditorStyle:id=216,.parent=17"
      _StyleDefs(624) =   "Splits(1).Columns(39).Style:id=209,.parent=13"
      _StyleDefs(625) =   "Splits(1).Columns(39).HeadingStyle:id=206,.parent=14"
      _StyleDefs(626) =   "Splits(1).Columns(39).FooterStyle:id=207,.parent=15"
      _StyleDefs(627) =   "Splits(1).Columns(39).EditorStyle:id=208,.parent=17"
      _StyleDefs(628) =   "Splits(1).Columns(40).Style:id=201,.parent=13"
      _StyleDefs(629) =   "Splits(1).Columns(40).HeadingStyle:id=198,.parent=14"
      _StyleDefs(630) =   "Splits(1).Columns(40).FooterStyle:id=199,.parent=15"
      _StyleDefs(631) =   "Splits(1).Columns(40).EditorStyle:id=200,.parent=17"
      _StyleDefs(632) =   "Splits(1).Columns(41).Style:id=193,.parent=13"
      _StyleDefs(633) =   "Splits(1).Columns(41).HeadingStyle:id=190,.parent=14"
      _StyleDefs(634) =   "Splits(1).Columns(41).FooterStyle:id=191,.parent=15"
      _StyleDefs(635) =   "Splits(1).Columns(41).EditorStyle:id=192,.parent=17"
      _StyleDefs(636) =   "Splits(1).Columns(42).Style:id=185,.parent=13"
      _StyleDefs(637) =   "Splits(1).Columns(42).HeadingStyle:id=182,.parent=14"
      _StyleDefs(638) =   "Splits(1).Columns(42).FooterStyle:id=183,.parent=15"
      _StyleDefs(639) =   "Splits(1).Columns(42).EditorStyle:id=184,.parent=17"
      _StyleDefs(640) =   "Splits(1).Columns(43).Style:id=177,.parent=13"
      _StyleDefs(641) =   "Splits(1).Columns(43).HeadingStyle:id=174,.parent=14"
      _StyleDefs(642) =   "Splits(1).Columns(43).FooterStyle:id=175,.parent=15"
      _StyleDefs(643) =   "Splits(1).Columns(43).EditorStyle:id=176,.parent=17"
      _StyleDefs(644) =   "Splits(1).Columns(44).Style:id=169,.parent=13"
      _StyleDefs(645) =   "Splits(1).Columns(44).HeadingStyle:id=166,.parent=14"
      _StyleDefs(646) =   "Splits(1).Columns(44).FooterStyle:id=167,.parent=15"
      _StyleDefs(647) =   "Splits(1).Columns(44).EditorStyle:id=168,.parent=17"
      _StyleDefs(648) =   "Splits(1).Columns(45).Style:id=161,.parent=13"
      _StyleDefs(649) =   "Splits(1).Columns(45).HeadingStyle:id=158,.parent=14"
      _StyleDefs(650) =   "Splits(1).Columns(45).FooterStyle:id=159,.parent=15"
      _StyleDefs(651) =   "Splits(1).Columns(45).EditorStyle:id=160,.parent=17"
      _StyleDefs(652) =   "Splits(1).Columns(46).Style:id=153,.parent=13"
      _StyleDefs(653) =   "Splits(1).Columns(46).HeadingStyle:id=150,.parent=14"
      _StyleDefs(654) =   "Splits(1).Columns(46).FooterStyle:id=151,.parent=15"
      _StyleDefs(655) =   "Splits(1).Columns(46).EditorStyle:id=152,.parent=17"
      _StyleDefs(656) =   "Splits(1).Columns(47).Style:id=74,.parent=13"
      _StyleDefs(657) =   "Splits(1).Columns(47).HeadingStyle:id=71,.parent=14"
      _StyleDefs(658) =   "Splits(1).Columns(47).FooterStyle:id=72,.parent=15"
      _StyleDefs(659) =   "Splits(1).Columns(47).EditorStyle:id=73,.parent=17"
      _StyleDefs(660) =   "Splits(1).Columns(48).Style:id=70,.parent=13"
      _StyleDefs(661) =   "Splits(1).Columns(48).HeadingStyle:id=67,.parent=14"
      _StyleDefs(662) =   "Splits(1).Columns(48).FooterStyle:id=68,.parent=15"
      _StyleDefs(663) =   "Splits(1).Columns(48).EditorStyle:id=69,.parent=17"
      _StyleDefs(664) =   "Splits(1).Columns(49).Style:id=66,.parent=13"
      _StyleDefs(665) =   "Splits(1).Columns(49).HeadingStyle:id=63,.parent=14"
      _StyleDefs(666) =   "Splits(1).Columns(49).FooterStyle:id=64,.parent=15"
      _StyleDefs(667) =   "Splits(1).Columns(49).EditorStyle:id=65,.parent=17"
      _StyleDefs(668) =   "Splits(1).Columns(50).Style:id=62,.parent=13"
      _StyleDefs(669) =   "Splits(1).Columns(50).HeadingStyle:id=59,.parent=14"
      _StyleDefs(670) =   "Splits(1).Columns(50).FooterStyle:id=60,.parent=15"
      _StyleDefs(671) =   "Splits(1).Columns(50).EditorStyle:id=61,.parent=17"
      _StyleDefs(672) =   "Splits(1).Columns(51).Style:id=58,.parent=13"
      _StyleDefs(673) =   "Splits(1).Columns(51).HeadingStyle:id=55,.parent=14"
      _StyleDefs(674) =   "Splits(1).Columns(51).FooterStyle:id=56,.parent=15"
      _StyleDefs(675) =   "Splits(1).Columns(51).EditorStyle:id=57,.parent=17"
      _StyleDefs(676) =   "Splits(1).Columns(52).Style:id=54,.parent=13"
      _StyleDefs(677) =   "Splits(1).Columns(52).HeadingStyle:id=51,.parent=14"
      _StyleDefs(678) =   "Splits(1).Columns(52).FooterStyle:id=52,.parent=15"
      _StyleDefs(679) =   "Splits(1).Columns(52).EditorStyle:id=53,.parent=17"
      _StyleDefs(680) =   "Splits(1).Columns(53).Style:id=50,.parent=13"
      _StyleDefs(681) =   "Splits(1).Columns(53).HeadingStyle:id=47,.parent=14"
      _StyleDefs(682) =   "Splits(1).Columns(53).FooterStyle:id=48,.parent=15"
      _StyleDefs(683) =   "Splits(1).Columns(53).EditorStyle:id=49,.parent=17"
      _StyleDefs(684) =   "Splits(1).Columns(54).Style:id=46,.parent=13"
      _StyleDefs(685) =   "Splits(1).Columns(54).HeadingStyle:id=43,.parent=14"
      _StyleDefs(686) =   "Splits(1).Columns(54).FooterStyle:id=44,.parent=15"
      _StyleDefs(687) =   "Splits(1).Columns(54).EditorStyle:id=45,.parent=17"
      _StyleDefs(688) =   "Splits(1).Columns(55).Style:id=32,.parent=13"
      _StyleDefs(689) =   "Splits(1).Columns(55).HeadingStyle:id=29,.parent=14"
      _StyleDefs(690) =   "Splits(1).Columns(55).FooterStyle:id=30,.parent=15"
      _StyleDefs(691) =   "Splits(1).Columns(55).EditorStyle:id=31,.parent=17"
      _StyleDefs(692) =   "Splits(1).Columns(56).Style:id=28,.parent=13"
      _StyleDefs(693) =   "Splits(1).Columns(56).HeadingStyle:id=25,.parent=14"
      _StyleDefs(694) =   "Splits(1).Columns(56).FooterStyle:id=26,.parent=15"
      _StyleDefs(695) =   "Splits(1).Columns(56).EditorStyle:id=27,.parent=17"
      _StyleDefs(696) =   "Splits(1).Columns(57).Style:id=78,.parent=13"
      _StyleDefs(697) =   "Splits(1).Columns(57).HeadingStyle:id=75,.parent=14"
      _StyleDefs(698) =   "Splits(1).Columns(57).FooterStyle:id=76,.parent=15"
      _StyleDefs(699) =   "Splits(1).Columns(57).EditorStyle:id=77,.parent=17"
      _StyleDefs(700) =   "Splits(1).Columns(58).Style:id=534,.parent=13"
      _StyleDefs(701) =   "Splits(1).Columns(58).HeadingStyle:id=531,.parent=14"
      _StyleDefs(702) =   "Splits(1).Columns(58).FooterStyle:id=532,.parent=15"
      _StyleDefs(703) =   "Splits(1).Columns(58).EditorStyle:id=533,.parent=17"
      _StyleDefs(704) =   "Splits(1).Columns(59).Style:id=91,.parent=13"
      _StyleDefs(705) =   "Splits(1).Columns(59).HeadingStyle:id=88,.parent=14"
      _StyleDefs(706) =   "Splits(1).Columns(59).FooterStyle:id=89,.parent=15"
      _StyleDefs(707) =   "Splits(1).Columns(59).EditorStyle:id=90,.parent=17"
      _StyleDefs(708) =   "Splits(1).Columns(60).Style:id=542,.parent=13"
      _StyleDefs(709) =   "Splits(1).Columns(60).HeadingStyle:id=539,.parent=14"
      _StyleDefs(710) =   "Splits(1).Columns(60).FooterStyle:id=540,.parent=15"
      _StyleDefs(711) =   "Splits(1).Columns(60).EditorStyle:id=541,.parent=17"
      _StyleDefs(712) =   "Splits(1).Columns(61).Style:id=550,.parent=13"
      _StyleDefs(713) =   "Splits(1).Columns(61).HeadingStyle:id=547,.parent=14"
      _StyleDefs(714) =   "Splits(1).Columns(61).FooterStyle:id=548,.parent=15"
      _StyleDefs(715) =   "Splits(1).Columns(61).EditorStyle:id=549,.parent=17"
      _StyleDefs(716) =   "Splits(1).Columns(62).Style:id=558,.parent=13"
      _StyleDefs(717) =   "Splits(1).Columns(62).HeadingStyle:id=555,.parent=14"
      _StyleDefs(718) =   "Splits(1).Columns(62).FooterStyle:id=556,.parent=15"
      _StyleDefs(719) =   "Splits(1).Columns(62).EditorStyle:id=557,.parent=17"
      _StyleDefs(720) =   "Splits(1).Columns(63).Style:id=566,.parent=13"
      _StyleDefs(721) =   "Splits(1).Columns(63).HeadingStyle:id=563,.parent=14"
      _StyleDefs(722) =   "Splits(1).Columns(63).FooterStyle:id=564,.parent=15"
      _StyleDefs(723) =   "Splits(1).Columns(63).EditorStyle:id=565,.parent=17"
      _StyleDefs(724) =   "Splits(1).Columns(64).Style:id=574,.parent=13"
      _StyleDefs(725) =   "Splits(1).Columns(64).HeadingStyle:id=571,.parent=14"
      _StyleDefs(726) =   "Splits(1).Columns(64).FooterStyle:id=572,.parent=15"
      _StyleDefs(727) =   "Splits(1).Columns(64).EditorStyle:id=573,.parent=17"
      _StyleDefs(728) =   "Splits(1).Columns(65).Style:id=582,.parent=13"
      _StyleDefs(729) =   "Splits(1).Columns(65).HeadingStyle:id=579,.parent=14"
      _StyleDefs(730) =   "Splits(1).Columns(65).FooterStyle:id=580,.parent=15"
      _StyleDefs(731) =   "Splits(1).Columns(65).EditorStyle:id=581,.parent=17"
      _StyleDefs(732) =   "Splits(1).Columns(66).Style:id=590,.parent=13"
      _StyleDefs(733) =   "Splits(1).Columns(66).HeadingStyle:id=587,.parent=14"
      _StyleDefs(734) =   "Splits(1).Columns(66).FooterStyle:id=588,.parent=15"
      _StyleDefs(735) =   "Splits(1).Columns(66).EditorStyle:id=589,.parent=17"
      _StyleDefs(736) =   "Splits(1).Columns(67).Style:id=598,.parent=13"
      _StyleDefs(737) =   "Splits(1).Columns(67).HeadingStyle:id=595,.parent=14"
      _StyleDefs(738) =   "Splits(1).Columns(67).FooterStyle:id=596,.parent=15"
      _StyleDefs(739) =   "Splits(1).Columns(67).EditorStyle:id=597,.parent=17"
      _StyleDefs(740) =   "Splits(1).Columns(68).Style:id=606,.parent=13"
      _StyleDefs(741) =   "Splits(1).Columns(68).HeadingStyle:id=603,.parent=14"
      _StyleDefs(742) =   "Splits(1).Columns(68).FooterStyle:id=604,.parent=15"
      _StyleDefs(743) =   "Splits(1).Columns(68).EditorStyle:id=605,.parent=17"
      _StyleDefs(744) =   "Splits(1).Columns(69).Style:id=614,.parent=13"
      _StyleDefs(745) =   "Splits(1).Columns(69).HeadingStyle:id=611,.parent=14"
      _StyleDefs(746) =   "Splits(1).Columns(69).FooterStyle:id=612,.parent=15"
      _StyleDefs(747) =   "Splits(1).Columns(69).EditorStyle:id=613,.parent=17"
      _StyleDefs(748) =   "Splits(1).Columns(70).Style:id=622,.parent=13"
      _StyleDefs(749) =   "Splits(1).Columns(70).HeadingStyle:id=619,.parent=14"
      _StyleDefs(750) =   "Splits(1).Columns(70).FooterStyle:id=620,.parent=15"
      _StyleDefs(751) =   "Splits(1).Columns(70).EditorStyle:id=621,.parent=17"
      _StyleDefs(752) =   "Splits(1).Columns(71).Style:id=630,.parent=13"
      _StyleDefs(753) =   "Splits(1).Columns(71).HeadingStyle:id=627,.parent=14"
      _StyleDefs(754) =   "Splits(1).Columns(71).FooterStyle:id=628,.parent=15"
      _StyleDefs(755) =   "Splits(1).Columns(71).EditorStyle:id=629,.parent=17"
      _StyleDefs(756) =   "Splits(1).Columns(72).Style:id=638,.parent=13"
      _StyleDefs(757) =   "Splits(1).Columns(72).HeadingStyle:id=635,.parent=14"
      _StyleDefs(758) =   "Splits(1).Columns(72).FooterStyle:id=636,.parent=15"
      _StyleDefs(759) =   "Splits(1).Columns(72).EditorStyle:id=637,.parent=17"
      _StyleDefs(760) =   "Splits(1).Columns(73).Style:id=646,.parent=13"
      _StyleDefs(761) =   "Splits(1).Columns(73).HeadingStyle:id=643,.parent=14"
      _StyleDefs(762) =   "Splits(1).Columns(73).FooterStyle:id=644,.parent=15"
      _StyleDefs(763) =   "Splits(1).Columns(73).EditorStyle:id=645,.parent=17"
      _StyleDefs(764) =   "Splits(1).Columns(74).Style:id=654,.parent=13"
      _StyleDefs(765) =   "Splits(1).Columns(74).HeadingStyle:id=651,.parent=14"
      _StyleDefs(766) =   "Splits(1).Columns(74).FooterStyle:id=652,.parent=15"
      _StyleDefs(767) =   "Splits(1).Columns(74).EditorStyle:id=653,.parent=17"
      _StyleDefs(768) =   "Splits(1).Columns(75).Style:id=662,.parent=13"
      _StyleDefs(769) =   "Splits(1).Columns(75).HeadingStyle:id=659,.parent=14"
      _StyleDefs(770) =   "Splits(1).Columns(75).FooterStyle:id=660,.parent=15"
      _StyleDefs(771) =   "Splits(1).Columns(75).EditorStyle:id=661,.parent=17"
      _StyleDefs(772) =   "Splits(1).Columns(76).Style:id=670,.parent=13"
      _StyleDefs(773) =   "Splits(1).Columns(76).HeadingStyle:id=667,.parent=14"
      _StyleDefs(774) =   "Splits(1).Columns(76).FooterStyle:id=668,.parent=15"
      _StyleDefs(775) =   "Splits(1).Columns(76).EditorStyle:id=669,.parent=17"
      _StyleDefs(776) =   "Splits(1).Columns(77).Style:id=678,.parent=13"
      _StyleDefs(777) =   "Splits(1).Columns(77).HeadingStyle:id=675,.parent=14"
      _StyleDefs(778) =   "Splits(1).Columns(77).FooterStyle:id=676,.parent=15"
      _StyleDefs(779) =   "Splits(1).Columns(77).EditorStyle:id=677,.parent=17"
      _StyleDefs(780) =   "Splits(1).Columns(78).Style:id=686,.parent=13"
      _StyleDefs(781) =   "Splits(1).Columns(78).HeadingStyle:id=683,.parent=14"
      _StyleDefs(782) =   "Splits(1).Columns(78).FooterStyle:id=684,.parent=15"
      _StyleDefs(783) =   "Splits(1).Columns(78).EditorStyle:id=685,.parent=17"
      _StyleDefs(784) =   "Splits(1).Columns(79).Style:id=694,.parent=13"
      _StyleDefs(785) =   "Splits(1).Columns(79).HeadingStyle:id=691,.parent=14"
      _StyleDefs(786) =   "Splits(1).Columns(79).FooterStyle:id=692,.parent=15"
      _StyleDefs(787) =   "Splits(1).Columns(79).EditorStyle:id=693,.parent=17"
      _StyleDefs(788) =   "Splits(1).Columns(80).Style:id=702,.parent=13"
      _StyleDefs(789) =   "Splits(1).Columns(80).HeadingStyle:id=699,.parent=14"
      _StyleDefs(790) =   "Splits(1).Columns(80).FooterStyle:id=700,.parent=15"
      _StyleDefs(791) =   "Splits(1).Columns(80).EditorStyle:id=701,.parent=17"
      _StyleDefs(792) =   "Splits(1).Columns(81).Style:id=710,.parent=13"
      _StyleDefs(793) =   "Splits(1).Columns(81).HeadingStyle:id=707,.parent=14"
      _StyleDefs(794) =   "Splits(1).Columns(81).FooterStyle:id=708,.parent=15"
      _StyleDefs(795) =   "Splits(1).Columns(81).EditorStyle:id=709,.parent=17"
      _StyleDefs(796) =   "Splits(1).Columns(82).Style:id=718,.parent=13"
      _StyleDefs(797) =   "Splits(1).Columns(82).HeadingStyle:id=715,.parent=14"
      _StyleDefs(798) =   "Splits(1).Columns(82).FooterStyle:id=716,.parent=15"
      _StyleDefs(799) =   "Splits(1).Columns(82).EditorStyle:id=717,.parent=17"
      _StyleDefs(800) =   "Splits(1).Columns(83).Style:id=726,.parent=13"
      _StyleDefs(801) =   "Splits(1).Columns(83).HeadingStyle:id=723,.parent=14"
      _StyleDefs(802) =   "Splits(1).Columns(83).FooterStyle:id=724,.parent=15"
      _StyleDefs(803) =   "Splits(1).Columns(83).EditorStyle:id=725,.parent=17"
      _StyleDefs(804) =   "Splits(1).Columns(84).Style:id=734,.parent=13"
      _StyleDefs(805) =   "Splits(1).Columns(84).HeadingStyle:id=731,.parent=14"
      _StyleDefs(806) =   "Splits(1).Columns(84).FooterStyle:id=732,.parent=15"
      _StyleDefs(807) =   "Splits(1).Columns(84).EditorStyle:id=733,.parent=17"
      _StyleDefs(808) =   "Splits(1).Columns(85).Style:id=742,.parent=13"
      _StyleDefs(809) =   "Splits(1).Columns(85).HeadingStyle:id=739,.parent=14"
      _StyleDefs(810) =   "Splits(1).Columns(85).FooterStyle:id=740,.parent=15"
      _StyleDefs(811) =   "Splits(1).Columns(85).EditorStyle:id=741,.parent=17"
      _StyleDefs(812) =   "Splits(1).Columns(86).Style:id=750,.parent=13"
      _StyleDefs(813) =   "Splits(1).Columns(86).HeadingStyle:id=747,.parent=14"
      _StyleDefs(814) =   "Splits(1).Columns(86).FooterStyle:id=748,.parent=15"
      _StyleDefs(815) =   "Splits(1).Columns(86).EditorStyle:id=749,.parent=17"
      _StyleDefs(816) =   "Splits(1).Columns(87).Style:id=758,.parent=13"
      _StyleDefs(817) =   "Splits(1).Columns(87).HeadingStyle:id=755,.parent=14"
      _StyleDefs(818) =   "Splits(1).Columns(87).FooterStyle:id=756,.parent=15"
      _StyleDefs(819) =   "Splits(1).Columns(87).EditorStyle:id=757,.parent=17"
      _StyleDefs(820) =   "Splits(1).Columns(88).Style:id=766,.parent=13"
      _StyleDefs(821) =   "Splits(1).Columns(88).HeadingStyle:id=763,.parent=14"
      _StyleDefs(822) =   "Splits(1).Columns(88).FooterStyle:id=764,.parent=15"
      _StyleDefs(823) =   "Splits(1).Columns(88).EditorStyle:id=765,.parent=17"
      _StyleDefs(824) =   "Splits(1).Columns(89).Style:id=774,.parent=13"
      _StyleDefs(825) =   "Splits(1).Columns(89).HeadingStyle:id=771,.parent=14"
      _StyleDefs(826) =   "Splits(1).Columns(89).FooterStyle:id=772,.parent=15"
      _StyleDefs(827) =   "Splits(1).Columns(89).EditorStyle:id=773,.parent=17"
      _StyleDefs(828) =   "Splits(1).Columns(90).Style:id=782,.parent=13"
      _StyleDefs(829) =   "Splits(1).Columns(90).HeadingStyle:id=779,.parent=14"
      _StyleDefs(830) =   "Splits(1).Columns(90).FooterStyle:id=780,.parent=15"
      _StyleDefs(831) =   "Splits(1).Columns(90).EditorStyle:id=781,.parent=17"
      _StyleDefs(832) =   "Splits(1).Columns(91).Style:id=790,.parent=13"
      _StyleDefs(833) =   "Splits(1).Columns(91).HeadingStyle:id=787,.parent=14"
      _StyleDefs(834) =   "Splits(1).Columns(91).FooterStyle:id=788,.parent=15"
      _StyleDefs(835) =   "Splits(1).Columns(91).EditorStyle:id=789,.parent=17"
      _StyleDefs(836) =   "Splits(1).Columns(92).Style:id=798,.parent=13"
      _StyleDefs(837) =   "Splits(1).Columns(92).HeadingStyle:id=795,.parent=14"
      _StyleDefs(838) =   "Splits(1).Columns(92).FooterStyle:id=796,.parent=15"
      _StyleDefs(839) =   "Splits(1).Columns(92).EditorStyle:id=797,.parent=17"
      _StyleDefs(840) =   "Splits(1).Columns(93).Style:id=806,.parent=13"
      _StyleDefs(841) =   "Splits(1).Columns(93).HeadingStyle:id=803,.parent=14"
      _StyleDefs(842) =   "Splits(1).Columns(93).FooterStyle:id=804,.parent=15"
      _StyleDefs(843) =   "Splits(1).Columns(93).EditorStyle:id=805,.parent=17"
      _StyleDefs(844) =   "Splits(1).Columns(94).Style:id=814,.parent=13"
      _StyleDefs(845) =   "Splits(1).Columns(94).HeadingStyle:id=811,.parent=14"
      _StyleDefs(846) =   "Splits(1).Columns(94).FooterStyle:id=812,.parent=15"
      _StyleDefs(847) =   "Splits(1).Columns(94).EditorStyle:id=813,.parent=17"
      _StyleDefs(848) =   "Splits(1).Columns(95).Style:id=822,.parent=13"
      _StyleDefs(849) =   "Splits(1).Columns(95).HeadingStyle:id=819,.parent=14"
      _StyleDefs(850) =   "Splits(1).Columns(95).FooterStyle:id=820,.parent=15"
      _StyleDefs(851) =   "Splits(1).Columns(95).EditorStyle:id=821,.parent=17"
      _StyleDefs(852) =   "Splits(1).Columns(96).Style:id=830,.parent=13"
      _StyleDefs(853) =   "Splits(1).Columns(96).HeadingStyle:id=827,.parent=14"
      _StyleDefs(854) =   "Splits(1).Columns(96).FooterStyle:id=828,.parent=15"
      _StyleDefs(855) =   "Splits(1).Columns(96).EditorStyle:id=829,.parent=17"
      _StyleDefs(856) =   "Splits(1).Columns(97).Style:id=838,.parent=13"
      _StyleDefs(857) =   "Splits(1).Columns(97).HeadingStyle:id=835,.parent=14"
      _StyleDefs(858) =   "Splits(1).Columns(97).FooterStyle:id=836,.parent=15"
      _StyleDefs(859) =   "Splits(1).Columns(97).EditorStyle:id=837,.parent=17"
      _StyleDefs(860) =   "Splits(1).Columns(98).Style:id=846,.parent=13"
      _StyleDefs(861) =   "Splits(1).Columns(98).HeadingStyle:id=843,.parent=14"
      _StyleDefs(862) =   "Splits(1).Columns(98).FooterStyle:id=844,.parent=15"
      _StyleDefs(863) =   "Splits(1).Columns(98).EditorStyle:id=845,.parent=17"
      _StyleDefs(864) =   "Splits(1).Columns(99).Style:id=854,.parent=13"
      _StyleDefs(865) =   "Splits(1).Columns(99).HeadingStyle:id=851,.parent=14"
      _StyleDefs(866) =   "Splits(1).Columns(99).FooterStyle:id=852,.parent=15"
      _StyleDefs(867) =   "Splits(1).Columns(99).EditorStyle:id=853,.parent=17"
      _StyleDefs(868) =   "Splits(1).Columns(100).Style:id=862,.parent=13"
      _StyleDefs(869) =   "Splits(1).Columns(100).HeadingStyle:id=859,.parent=14"
      _StyleDefs(870) =   "Splits(1).Columns(100).FooterStyle:id=860,.parent=15"
      _StyleDefs(871) =   "Splits(1).Columns(100).EditorStyle:id=861,.parent=17"
      _StyleDefs(872) =   "Splits(1).Columns(101).Style:id=870,.parent=13"
      _StyleDefs(873) =   "Splits(1).Columns(101).HeadingStyle:id=867,.parent=14"
      _StyleDefs(874) =   "Splits(1).Columns(101).FooterStyle:id=868,.parent=15"
      _StyleDefs(875) =   "Splits(1).Columns(101).EditorStyle:id=869,.parent=17"
      _StyleDefs(876) =   "Splits(1).Columns(102).Style:id=878,.parent=13"
      _StyleDefs(877) =   "Splits(1).Columns(102).HeadingStyle:id=875,.parent=14"
      _StyleDefs(878) =   "Splits(1).Columns(102).FooterStyle:id=876,.parent=15"
      _StyleDefs(879) =   "Splits(1).Columns(102).EditorStyle:id=877,.parent=17"
      _StyleDefs(880) =   "Splits(1).Columns(103).Style:id=886,.parent=13"
      _StyleDefs(881) =   "Splits(1).Columns(103).HeadingStyle:id=883,.parent=14"
      _StyleDefs(882) =   "Splits(1).Columns(103).FooterStyle:id=884,.parent=15"
      _StyleDefs(883) =   "Splits(1).Columns(103).EditorStyle:id=885,.parent=17"
      _StyleDefs(884) =   "Splits(1).Columns(104).Style:id=894,.parent=13"
      _StyleDefs(885) =   "Splits(1).Columns(104).HeadingStyle:id=891,.parent=14"
      _StyleDefs(886) =   "Splits(1).Columns(104).FooterStyle:id=892,.parent=15"
      _StyleDefs(887) =   "Splits(1).Columns(104).EditorStyle:id=893,.parent=17"
      _StyleDefs(888) =   "Named:id=33:Normal"
      _StyleDefs(889) =   ":id=33,.parent=0"
      _StyleDefs(890) =   "Named:id=34:Heading"
      _StyleDefs(891) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(892) =   ":id=34,.wraptext=-1"
      _StyleDefs(893) =   "Named:id=35:Footing"
      _StyleDefs(894) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(895) =   "Named:id=36:Selected"
      _StyleDefs(896) =   ":id=36,.parent=33,.bgcolor=&H80FF&,.fgcolor=&H8000000E&"
      _StyleDefs(897) =   "Named:id=37:Caption"
      _StyleDefs(898) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(899) =   "Named:id=38:HighlightRow"
      _StyleDefs(900) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(901) =   "Named:id=39:EvenRow"
      _StyleDefs(902) =   ":id=39,.parent=33"
      _StyleDefs(903) =   "Named:id=40:OddRow"
      _StyleDefs(904) =   ":id=40,.parent=33"
      _StyleDefs(905) =   "Named:id=41:RecordSelector"
      _StyleDefs(906) =   ":id=41,.parent=34"
      _StyleDefs(907) =   "Named:id=42:FilterBar"
      _StyleDefs(908) =   ":id=42,.parent=33"
      _StyleDefs(909) =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(910) =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(911) =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(912) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(913) =   "Named:id=87:SudahDitagih"
      _StyleDefs(914) =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(915) =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(916) =   ":id=87,.fontname=Tahoma"
      _StyleDefs(917) =   "Named:id=514:Kuning"
      _StyleDefs(918) =   ":id=514,.parent=87,.bgcolor=&HFFFF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(919) =   ":id=514,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(920) =   ":id=514,.fontname=Tahoma"
      _StyleDefs(921) =   "Named:id=515:Biru"
      _StyleDefs(922) =   ":id=515,.parent=514,.bgcolor=&HFF8080&"
      _StyleDefs(923) =   "Named:id=516:Merah"
      _StyleDefs(924) =   ":id=516,.parent=515,.bgcolor=&H88A0F7&,.fgcolor=&H0&"
      _StyleDefs(925) =   "Named:id=517:Hijau"
      _StyleDefs(926) =   ":id=517,.parent=516,.bgcolor=&HFF00&"
      _StyleDefs(927) =   "Named:id=518:Putih"
      _StyleDefs(928) =   ":id=518,.parent=517,.bgcolor=&HFFFFFF&"
   End
   Begin TDBDate6Ctl.TDBDate Tgl1 
      Height          =   330
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInfoPelanggan.frx":364E
      Caption         =   "frmInfoPelanggan.frx":377A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInfoPelanggan.frx":37E6
      Keys            =   "frmInfoPelanggan.frx":3804
      Spin            =   "frmInfoPelanggan.frx":3862
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
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInfoPelanggan.frx":388A
      Caption         =   "frmInfoPelanggan.frx":39B6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInfoPelanggan.frx":3A22
      Keys            =   "frmInfoPelanggan.frx":3A40
      Spin            =   "frmInfoPelanggan.frx":3A9E
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
   Begin TDBDate6Ctl.TDBDate Tgl3 
      Height          =   330
      Left            =   1920
      TabIndex        =   11
      Top             =   480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInfoPelanggan.frx":3AC6
      Caption         =   "frmInfoPelanggan.frx":3BF2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInfoPelanggan.frx":3C5E
      Keys            =   "frmInfoPelanggan.frx":3C7C
      Spin            =   "frmInfoPelanggan.frx":3CDA
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
   Begin TDBDate6Ctl.TDBDate Tgl4 
      Height          =   330
      Left            =   3960
      TabIndex        =   12
      Top             =   480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInfoPelanggan.frx":3D02
      Caption         =   "frmInfoPelanggan.frx":3E2E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInfoPelanggan.frx":3E9A
      Keys            =   "frmInfoPelanggan.frx":3EB8
      Spin            =   "frmInfoPelanggan.frx":3F16
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
   Begin ucXPButton.XPButton cmdRef 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
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
      MICON           =   "frmInfoPelanggan.frx":3F3E
      PICN            =   "frmInfoPelanggan.frx":3F5A
      PICH            =   "frmInfoPelanggan.frx":44D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame FrameAktifBulan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   21
      Top             =   120
      Width           =   1335
      Begin TDBNumber6Ctl.TDBNumber blnAktif 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmInfoPelanggan.frx":4B0D
         Caption         =   "frmInfoPelanggan.frx":4B2D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelanggan.frx":4B99
         Keys            =   "frmInfoPelanggan.frx":4BB7
         Spin            =   "frmInfoPelanggan.frx":4C01
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1376261
         Value           =   7
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   23
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000C000&
      Caption         =   "GRAFIK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   8520
      Width           =   1215
   End
   Begin TDBDate6Ctl.TDBDate tglDDate1 
      Height          =   330
      Left            =   1920
      TabIndex        =   95
      Top             =   840
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInfoPelanggan.frx":4C29
      Caption         =   "frmInfoPelanggan.frx":4D55
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInfoPelanggan.frx":4DC1
      Keys            =   "frmInfoPelanggan.frx":4DDF
      Spin            =   "frmInfoPelanggan.frx":4E3D
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
   Begin TDBDate6Ctl.TDBDate tglDDate2 
      Height          =   330
      Left            =   3960
      TabIndex        =   96
      Top             =   840
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInfoPelanggan.frx":4E65
      Caption         =   "frmInfoPelanggan.frx":4F91
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInfoPelanggan.frx":4FFD
      Keys            =   "frmInfoPelanggan.frx":501B
      Spin            =   "frmInfoPelanggan.frx":5079
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
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   360
      TabIndex        =   98
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "                               s/d"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   97
      Top             =   840
      Width           =   2205
   End
   Begin VB.Label LabelPasif2 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kuning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   75
      Top             =   840
      Width           =   735
   End
   Begin VB.Label LabelPasif3 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Merah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   73
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label LabelPasif 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   51
      Top             =   480
      Width           =   735
   End
   Begin VB.Label LabelAktif 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hijau"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   41
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblJumlah 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   8520
      Width           =   4065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH DATA = "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   8520
      Width           =   2655
   End
   Begin VB.Label lblPengambilanData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROSES PENGAMBILAN DATA"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Area : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1200
      TabIndex        =   13
      Top             =   1250
      Width           =   645
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "    :                          s/d"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   480
      Width           =   2205
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last SBE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   9
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: DAFTAR PELANGGAN ::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6480
      TabIndex        =   4
      Top             =   2100
      Width           =   2760
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Service"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":                              s/d"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2205
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   14055
   End
End
Attribute VB_Name = "frmInfoPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRecord As New ADODB.Recordset
Dim rsCari As ADODB.Recordset

Dim cols As TrueOleDBGrid70.Columns
Private Urut As String
Private jmlTagih As Integer
Dim jumAktif As Integer
Dim jumPasif1 As Integer
Dim jumPasif2 As Integer
Dim jumPasif3 As Integer

Dim jumlahLastGr As Integer
Dim jumlanLastSBE As Integer
Dim jumlahService As Integer

Private Sub cmdBatal_Click()
    frmeKonfirm.Visible = False
End Sub

Private Sub cbOperator_Change()
    If cbOperator.Text = "RANGE" Then
        FrameAktifRange.Visible = True
        FrameAktifBulan.Visible = False
    Else
        FrameAktifBulan.Visible = True
       FrameAktifRange.Visible = False
    End If
End Sub

Private Sub cbOperator_Click()
    Call cbOperator_Change
End Sub

Private Sub Check1_Click()
    If Check1.value = vbChecked Then
        If Log.LogHOConn = 0 Then
            MsgBox "Koneksi ke HO terputus", vbInformation
            Check1.value = vbUnchecked
            
        End If
    End If
End Sub

Private Sub ckAktif_Click()
    If ckAktif.value = vbChecked Then
        cbOperator.Enabled = True
        cmdWarna.Enabled = True
        FrameAktifBulan.Enabled = True
        FrameAktifRange.Enabled = True
    Else
        cbOperator.Enabled = False
        cmdWarna.Enabled = False
        FrameAktifBulan.Enabled = False
        FrameAktifRange.Enabled = False
    End If
End Sub

Private Sub cmdExport_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsRecord.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Reminder Service Kosong !"
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
16        gridTagih.Columns(7).Visible = False
17         gridTagih.ExportToFile onVal, True
18        gridTagih.Columns(7).Visible = True
19                Er.HeadCap = "Export to excel !"
20                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
21                Beep
22                Er.show (vbModal)
23         Exit Sub
Err_Export:
24          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
25          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:

End Sub

Private Sub cmdExportCSV_Click()
1     On Error GoTo Err_Export
      Dim nmFile As String
      '2     CDialog.filter = "*.csv"
      '3     CDialog.FileName = LblRincian & " " & NamaCabang & " per. " & Format("01/" & Right(Dbs, 2) & "/" & Mid(Dbs, 7, 4), "mm-yyyy")
      '4     CDialog.CancelError = True
      '5     CDialog.ShowSave
      '6         nmFile = CDialog.FileName & ".csv"
      '7         GridGSO.ExportToDelimitedFile nmFile, , ";", , , True
      '8         If MsgBox("File " & nmFile & " telah di Export." + vbCrLf + "Apakah mau dibuka sekarang?", vbQuestion + vbYesNo, "Export File") = vbYes Then
      '9            Shell "Explorer " & nmFile, vbMaximizedFocus
      '10        End If
      '
          Dim Er As New ErrInfo
2         If rsRecord.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Reminder Service Kosong !"
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
15        onVal = MyPath & f.FileName & ".csv"
16        gridTagih.Columns(7).Visible = False
17        gridTagih.ExportToDelimitedFile onVal, , ";", , , True
18        gridTagih.Columns(7).Visible = True
19        If MsgBox("File " & nmFile & " telah di Export." + vbCrLf + "Apakah mau dibuka sekarang?", vbQuestion + vbYesNo, "Export File") = vbYes Then
20           Shell "Explorer " & onVal, vbMaximizedFocus
21        End If
22         Exit Sub
Err_Export:

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdKonfir_Click()
    If CInt(nTagih.Caption) < 1 Then
        MsgBox "Anda belum memilih Tagihan Cuci yang akan ditagihkan !", vbExclamation + vbOKOnly
        Exit Sub
    End If
        For Each col In gridTagih.Columns
            col.FilterText = ""
        Next col
        rsRecord.filter = ""

    frmeKonfirm.Visible = True
End Sub

Private Sub cmdRef_Click()
    lblPengambilanData.Visible = False
    'Call refreshGridBaru
    'Call refreshGrid
    Call RefreshGridMapping
    
End Sub


Private Sub cmdRetension_Click()
    Frame2.Visible = True
End Sub

 

Private Sub Command8_Click()

1        On Error GoTo cmdRetension_Click_Error
2          Screen.MousePointer = vbHourglass
3         Set Report = New rptRepoMappingRetention
          
          Dim oRs As ADODB.Recordset
          Dim rsTempAll As ADODB.Recordset
          Dim rsTemp As ADODB.Recordset
          
          Dim konter As Integer
          Dim isTambah As Boolean
          
          Dim total1 As Integer
          Dim total2 As Integer
          Dim total3 As Integer
          Dim total4 As Integer
          Dim total5 As Integer
          Dim total6 As Integer
          Dim total7 As Integer
          Dim total8 As Integer
          Dim total9 As Integer
          Dim total10 As Integer
          
          Dim Cetak As Double
          'Call createRS
          
          Dim i As Integer
4             i = 1
5             konter = 1
6             ProgressBar2.Visible = True
7             ProgressBar2.Min = 0
8             ProgressBar2.Max = CInt(cbBulanRet.Text)
9         Do While i <= CInt(cbBulanRet.Text)
10            lblPengambilanDataRet.Visible = True
11            sSQL = "select COUNT(noinv) total,fs_rc"
12            sSQL = sSQL & " From T_INVOICE"
13            sSQL = sSQL & " where NOINV <> '' AND FS_NO_RETUR = '' "
14            sSQL = sSQL & " and FS_RC <> '' and noinv not like '%-B' and FS_RC in ('R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') "
15            sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "' and dateadd(day,-1,(DATEADD(month,1,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))"
16            sSQL = sSQL & " grOUP BY FS_RC"
17            sSQL = sSQL & " order by FS_RC"
18            Set rsTempAll = objO.FetchData(sSQL)
              
      '        sSQL = "select COUNT(noinv) total,fs_rc"
      '        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
      '        sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
      '        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
      '        sSQL = sSQL & " and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') "
      '        sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "' and dateadd(day,-1,(DATEADD(month,1,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))"
      '        sSQL = sSQL & " grOUP BY FS_RC"
      '        sSQL = sSQL & " order by FS_RC"
      '        Set rsTemp = objO.FetchData(sSQL)
              
              ''--------------------------------------------
              '' 2016-02-25 || ganti sales unit
              ''--------------------------------------------
19            If Log.RPLHistori = 0 Then
20                sSQL = "select COUNT(noinv) total,fs_rc"
21                sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
22                sSQL = sSQL & " INNER JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
23                sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
24                sSQL = sSQL & " where NOINV <> '' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
25                sSQL = sSQL & " and noinv not like '%-B' and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') "
26                sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "' and dateadd(day,-1,(DATEADD(month,1,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))"
27                sSQL = sSQL & " grOUP BY FS_RC"
28                sSQL = sSQL & " order by FS_RC"
29            Else
30                sSQL = "select COUNT(noinv) total,fs_rc"
31                sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
32                sSQL = sSQL & " INNER JOIN [" & dbReplikasiHO & "]..summary s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
33                sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
34                sSQL = sSQL & " where NOINV <> '' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
35                sSQL = sSQL & " and noinv not like '%-B' and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') "
36                sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "' and dateadd(day,-1,(DATEADD(month,1,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))"
37                sSQL = sSQL & " grOUP BY FS_RC"
38                sSQL = sSQL & " order by FS_RC"

39            End If
40            Set rsTemp = objO.FetchData(sSQL)
              ''--------------------------------------------
              
41            total1 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
42            total2 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
43            total3 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
44            total4 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
45            total5 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
46            total6 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
47            total7 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
48            total8 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
49            total9 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
50            total10 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
              
51            lblPengambilanDataRet.Visible = False
52            If Not rsTemp.EOF Then
              
53                rsTemp.filter = "fs_rc = 'R010'"
54                If Not rsTemp.EOF Then
55                    If total1 > 0 Then
56                        Cetak = Round((CInt(rsTemp("TOTAL")) / total1) * 100)
57                    Else
58                        Cetak = 0
59                    End If
60                End If
61                rsTemp.filter = ""
62                If i = 1 Then Report.txtS11.SetText Cetak & " %"
63                If i = 2 Then Report.txtS21.SetText Cetak & " %"
64                If i = 3 Then Report.txtS31.SetText Cetak & " %"
65                If i = 4 Then Report.txtS41.SetText Cetak & " %"
66                If i = 5 Then Report.txtS51.SetText Cetak & " %"
67                If i = 6 Then Report.txtS61.SetText Cetak & " %"
68                If i = 7 Then Report.txtS71.SetText Cetak & " %"
69                If i = 8 Then Report.txtS81.SetText Cetak & " %"
70                If i = 9 Then Report.txtS91.SetText Cetak & " %"
71                If i = 10 Then Report.txtS101.SetText Cetak & " %"
72                If i = 11 Then Report.txtS111.SetText Cetak & " %"
73                If i = 12 Then Report.txtS121.SetText Cetak & " %"
74                DoEvents
                  
75                rsTemp.filter = "fs_rc = 'R020'"
76                If Not rsTemp.EOF Then
77                    If total2 > 0 Then
78                        Cetak = Round((CInt(rsTemp("TOTAL")) / total2) * 100)
79                    Else
80                        Cetak = 0
81                    End If
82                End If
83                rsTemp.filter = ""
84                If i = 1 Then Report.txtS12.SetText Cetak & " %"
85                If i = 2 Then Report.txtS22.SetText Cetak & " %"
86                If i = 3 Then Report.txtS32.SetText Cetak & " %"
87                If i = 4 Then Report.txtS42.SetText Cetak & " %"
88                If i = 5 Then Report.txtS52.SetText Cetak & " %"
89                If i = 6 Then Report.txtS62.SetText Cetak & " %"
90                If i = 7 Then Report.txtS72.SetText Cetak & " %"
91                If i = 8 Then Report.txtS82.SetText Cetak & " %"
92                If i = 9 Then Report.txtS92.SetText Cetak & " %"
93                If i = 10 Then Report.txtS102.SetText Cetak & " %"
94                If i = 11 Then Report.txtS112.SetText Cetak & " %"
95                If i = 12 Then Report.txtS122.SetText Cetak & " %"
96                DoEvents
                  
97                rsTemp.filter = "fs_rc = 'R030'"
98                If Not rsTemp.EOF Then
99                    If total3 > 0 Then
100                       Cetak = Round((CInt(rsTemp("TOTAL")) / total3) * 100)
101                   Else
102                       Cetak = 0
103                   End If
104               End If
105               rsTemp.filter = ""
106               If i = 1 Then Report.txtS13.SetText Cetak & " %"
107               If i = 2 Then Report.txtS23.SetText Cetak & " %"
108               If i = 3 Then Report.txtS33.SetText Cetak & " %"
109               If i = 4 Then Report.txtS43.SetText Cetak & " %"
110               If i = 5 Then Report.txtS53.SetText Cetak & " %"
111               If i = 6 Then Report.txtS63.SetText Cetak & " %"
112               If i = 7 Then Report.txtS73.SetText Cetak & " %"
113               If i = 8 Then Report.txtS83.SetText Cetak & " %"
114               If i = 9 Then Report.txtS93.SetText Cetak & " %"
115               If i = 10 Then Report.txtS103.SetText Cetak & " %"
116               If i = 11 Then Report.txtS113.SetText Cetak & " %"
117               If i = 12 Then Report.txtS123.SetText Cetak & " %"
118               DoEvents
                  
119               rsTemp.filter = "fs_rc = 'R040'"
120               If Not rsTemp.EOF Then
121                   If total4 > 0 Then
122                       Cetak = Round((CInt(rsTemp("TOTAL")) / total4) * 100)
123                   Else
124                       Cetak = 0
125                   End If
126               End If
127               rsTemp.filter = ""
128               If i = 1 Then Report.txtS14.SetText Cetak & " %"
129               If i = 2 Then Report.txtS24.SetText Cetak & " %"
130               If i = 3 Then Report.txtS34.SetText Cetak & " %"
131               If i = 4 Then Report.txtS44.SetText Cetak & " %"
132               If i = 5 Then Report.txtS54.SetText Cetak & " %"
133               If i = 6 Then Report.txtS64.SetText Cetak & " %"
134               If i = 7 Then Report.txtS74.SetText Cetak & " %"
135               If i = 8 Then Report.txtS84.SetText Cetak & " %"
136               If i = 9 Then Report.txtS94.SetText Cetak & " %"
137               If i = 10 Then Report.txtS104.SetText Cetak & " %"
138               If i = 11 Then Report.txtS114.SetText Cetak & " %"
139               If i = 12 Then Report.txtS124.SetText Cetak & " %"
140               DoEvents
                  
                  
141               rsTemp.filter = "fs_rc = 'R050'"
142               If Not rsTemp.EOF Then
143                   If total5 > 0 Then
144                       Cetak = Round((CInt(rsTemp("TOTAL")) / total5) * 100)
145                   Else
146                       Cetak = 0
147                   End If
148               End If
149               rsTemp.filter = ""
150               If i = 1 Then Report.txtS15.SetText Cetak & " %"
151               If i = 2 Then Report.txtS25.SetText Cetak & " %"
152               If i = 3 Then Report.txtS35.SetText Cetak & " %"
153               If i = 4 Then Report.txtS45.SetText Cetak & " %"
154               If i = 5 Then Report.txtS55.SetText Cetak & " %"
155               If i = 6 Then Report.txtS65.SetText Cetak & " %"
156               If i = 7 Then Report.txtS75.SetText Cetak & " %"
157               If i = 8 Then Report.txtS85.SetText Cetak & " %"
158               If i = 9 Then Report.txtS95.SetText Cetak & " %"
159               If i = 10 Then Report.txtS105.SetText Cetak & " %"
160               If i = 11 Then Report.txtS115.SetText Cetak & " %"
161               If i = 12 Then Report.txtS125.SetText Cetak & " %"
162               DoEvents
                  
                  
163               rsTemp.filter = "fs_rc = 'R060'"
164               If Not rsTemp.EOF Then
165                   If total6 > 0 Then
166                       Cetak = Round((CInt(rsTemp("TOTAL")) / total6) * 100)
167                   Else
168                       Cetak = 0
169                   End If
170               End If
171               rsTemp.filter = ""
172               If i = 1 Then Report.txtS16.SetText Cetak & " %"
173               If i = 2 Then Report.txtS26.SetText Cetak & " %"
174               If i = 3 Then Report.txtS36.SetText Cetak & " %"
175               If i = 4 Then Report.txtS46.SetText Cetak & " %"
176               If i = 5 Then Report.txtS56.SetText Cetak & " %"
177               If i = 6 Then Report.txtS66.SetText Cetak & " %"
178               If i = 7 Then Report.txtS76.SetText Cetak & " %"
179               If i = 8 Then Report.txtS86.SetText Cetak & " %"
180               If i = 9 Then Report.txtS96.SetText Cetak & " %"
181               If i = 10 Then Report.txtS106.SetText Cetak & " %"
182               If i = 11 Then Report.txtS116.SetText Cetak & " %"
183               If i = 12 Then Report.txtS126.SetText Cetak & " %"
184               DoEvents
                  
                  
185               rsTemp.filter = "fs_rc = 'R070'"
186               If Not rsTemp.EOF Then
187                   If total7 > 0 Then
188                       Cetak = Round((CInt(rsTemp("TOTAL")) / total7) * 100)
189                   Else
190                       Cetak = 0
191                   End If
192               End If
193               rsTemp.filter = ""
194               If i = 1 Then Report.txtS17.SetText Cetak & " %"
195               If i = 2 Then Report.txtS27.SetText Cetak & " %"
196               If i = 3 Then Report.txtS37.SetText Cetak & " %"
197               If i = 4 Then Report.txtS47.SetText Cetak & " %"
198               If i = 5 Then Report.txtS57.SetText Cetak & " %"
199               If i = 6 Then Report.txtS67.SetText Cetak & " %"
200               If i = 7 Then Report.txtS77.SetText Cetak & " %"
201               If i = 8 Then Report.txtS87.SetText Cetak & " %"
202               If i = 9 Then Report.txtS97.SetText Cetak & " %"
203               If i = 10 Then Report.txtS107.SetText Cetak & " %"
204               If i = 11 Then Report.txtS117.SetText Cetak & " %"
205               If i = 12 Then Report.txtS127.SetText Cetak & " %"
206               DoEvents
                  
207               rsTemp.filter = "fs_rc = 'R080'"
208               If Not rsTemp.EOF Then
209                   If total8 > 0 Then
210                       Cetak = Round((CInt(rsTemp("TOTAL")) / total8) * 100)
211                   Else
212                       Cetak = 0
213                   End If
214               End If
215               rsTemp.filter = ""
216               If i = 1 Then Report.txtS18.SetText Cetak & " %"
217               If i = 2 Then Report.txtS28.SetText Cetak & " %"
218               If i = 3 Then Report.txtS38.SetText Cetak & " %"
219               If i = 4 Then Report.txtS48.SetText Cetak & " %"
220               If i = 5 Then Report.txtS58.SetText Cetak & " %"
221               If i = 6 Then Report.txtS68.SetText Cetak & " %"
222               If i = 7 Then Report.txtS78.SetText Cetak & " %"
223               If i = 8 Then Report.txtS88.SetText Cetak & " %"
224               If i = 9 Then Report.txtS98.SetText Cetak & " %"
225               If i = 10 Then Report.txtS108.SetText Cetak & " %"
226               If i = 11 Then Report.txtS118.SetText Cetak & " %"
227               If i = 12 Then Report.txtS128.SetText Cetak & " %"
228               DoEvents
                  
                  
229               rsTemp.filter = "fs_rc = 'R090'"
230               If Not rsTemp.EOF Then
231                   If total9 > 0 Then
232                       Cetak = Round((CInt(rsTemp("TOTAL")) / total9) * 100)
233                   Else
234                       Cetak = 0
235                   End If
236               End If
237               rsTemp.filter = ""
238               If i = 1 Then Report.txtS19.SetText Cetak & " %"
239               If i = 2 Then Report.txtS29.SetText Cetak & " %"
240               If i = 3 Then Report.txtS39.SetText Cetak & " %"
241               If i = 4 Then Report.txtS49.SetText Cetak & " %"
242               If i = 5 Then Report.txtS59.SetText Cetak & " %"
243               If i = 6 Then Report.txtS69.SetText Cetak & " %"
244               If i = 7 Then Report.txtS79.SetText Cetak & " %"
245               If i = 8 Then Report.txtS89.SetText Cetak & " %"
246               If i = 9 Then Report.txtS99.SetText Cetak & " %"
247               If i = 10 Then Report.txtS109.SetText Cetak & " %"
248               If i = 11 Then Report.txtS119.SetText Cetak & " %"
249               If i = 12 Then Report.txtS129.SetText Cetak & " %"
250               DoEvents
                  
                  
251               rsTemp.filter = "fs_rc = 'R100'"
252               If Not rsTemp.EOF Then
253                   If total10 > 0 Then
254                       Cetak = Round((CInt(rsTemp("TOTAL")) / total10) * 100)
255                   Else
256                       Cetak = 0
257                   End If
258               End If
259               rsTemp.filter = ""
260               If i = 1 Then Report.txtS110.SetText Cetak & " %"
261               If i = 2 Then Report.txtS210.SetText Cetak & " %"
262               If i = 3 Then Report.txtS310.SetText Cetak & " %"
263               If i = 4 Then Report.txtS410.SetText Cetak & " %"
264               If i = 5 Then Report.txtS510.SetText Cetak & " %"
265               If i = 6 Then Report.txtS610.SetText Cetak & " %"
266               If i = 7 Then Report.txtS710.SetText Cetak & " %"
267               If i = 8 Then Report.txtS810.SetText Cetak & " %"
268               If i = 9 Then Report.txtS910.SetText Cetak & " %"
269               If i = 10 Then Report.txtS1010.SetText Cetak & " %"
270               If i = 11 Then Report.txtS1110.SetText Cetak & " %"
271               If i = 12 Then Report.txtS1210.SetText Cetak & " %"
272               DoEvents
273           End If
              
274           If Not rsTempAll.EOF Then
              
275               rsTempAll.filter = "fs_rc = 'R010'"
276               If Not rsTempAll.EOF Then
277                   If total1 > 0 Then
278                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total1) * 100)
279                   Else
280                       Cetak = 0
281                   End If
282               End If
283               rsTempAll.filter = ""
284               If i = 1 Then Report.txtA11.SetText Cetak & " %"
285               If i = 2 Then Report.txtA21.SetText Cetak & " %"
286               If i = 3 Then Report.txtA31.SetText Cetak & " %"
287               If i = 4 Then Report.txtA41.SetText Cetak & " %"
288               If i = 5 Then Report.txtA51.SetText Cetak & " %"
289               If i = 6 Then Report.txtA61.SetText Cetak & " %"
290               If i = 7 Then Report.txtA71.SetText Cetak & " %"
291               If i = 8 Then Report.txtA81.SetText Cetak & " %"
292               If i = 9 Then Report.txtA91.SetText Cetak & " %"
293               If i = 10 Then Report.txtA101.SetText Cetak & " %"
294               If i = 11 Then Report.txtA111.SetText Cetak & " %"
295               If i = 12 Then Report.txtA121.SetText Cetak & " %"
296               DoEvents
                  
297               rsTempAll.filter = "fs_rc = 'R020'"
298               If Not rsTempAll.EOF Then
299                   If total2 > 0 Then
300                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total2) * 100)
301                   Else
302                       Cetak = 0
303                   End If
304               End If
305               rsTempAll.filter = ""
306               If i = 1 Then Report.txtA12.SetText Cetak & " %"
307               If i = 2 Then Report.txtA22.SetText Cetak & " %"
308               If i = 3 Then Report.txtA32.SetText Cetak & " %"
309               If i = 4 Then Report.txtA42.SetText Cetak & " %"
310               If i = 5 Then Report.txtA52.SetText Cetak & " %"
311               If i = 6 Then Report.txtA62.SetText Cetak & " %"
312               If i = 7 Then Report.txtA72.SetText Cetak & " %"
313               If i = 8 Then Report.txtA82.SetText Cetak & " %"
314               If i = 9 Then Report.txtA92.SetText Cetak & " %"
315               If i = 10 Then Report.txtA102.SetText Cetak & " %"
316               If i = 11 Then Report.txtA112.SetText Cetak & " %"
317               If i = 12 Then Report.txtA122.SetText Cetak & " %"
318               DoEvents
                  
319               rsTempAll.filter = "fs_rc = 'R030'"
320               If Not rsTempAll.EOF Then
321                   If total3 > 0 Then
322                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total3) * 100)
323                   Else
324                       Cetak = 0
325                   End If
326               End If
327               rsTempAll.filter = ""
328               If i = 1 Then Report.txtA13.SetText Cetak & " %"
329               If i = 2 Then Report.txtA23.SetText Cetak & " %"
330               If i = 3 Then Report.txtA33.SetText Cetak & " %"
331               If i = 4 Then Report.txtA43.SetText Cetak & " %"
332               If i = 5 Then Report.txtA53.SetText Cetak & " %"
333               If i = 6 Then Report.txtA63.SetText Cetak & " %"
334               If i = 7 Then Report.txtA73.SetText Cetak & " %"
335               If i = 8 Then Report.txtA83.SetText Cetak & " %"
336               If i = 9 Then Report.txtA93.SetText Cetak & " %"
337               If i = 10 Then Report.txtA103.SetText Cetak & " %"
338               If i = 11 Then Report.txtA113.SetText Cetak & " %"
339               If i = 12 Then Report.txtA123.SetText Cetak & " %"
340               DoEvents
                  
341               rsTempAll.filter = "fs_rc = 'R040'"
342               If Not rsTempAll.EOF Then
343                   If total4 > 0 Then
344                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total4) * 100)
345                   Else
346                       Cetak = 0
347                   End If
348               End If
349               rsTempAll.filter = ""
350               If i = 1 Then Report.txtA14.SetText Cetak & " %"
351               If i = 2 Then Report.txtA24.SetText Cetak & " %"
352               If i = 3 Then Report.txtA34.SetText Cetak & " %"
353               If i = 4 Then Report.txtA44.SetText Cetak & " %"
354               If i = 5 Then Report.txtA54.SetText Cetak & " %"
355               If i = 6 Then Report.txtA64.SetText Cetak & " %"
356               If i = 7 Then Report.txtA74.SetText Cetak & " %"
357               If i = 8 Then Report.txtA84.SetText Cetak & " %"
358               If i = 9 Then Report.txtA94.SetText Cetak & " %"
359               If i = 10 Then Report.txtA104.SetText Cetak & " %"
360               If i = 11 Then Report.txtA114.SetText Cetak & " %"
361               If i = 12 Then Report.txtA124.SetText Cetak & " %"
362               DoEvents
                  
                  
363               rsTempAll.filter = "fs_rc = 'R050'"
364               If Not rsTempAll.EOF Then
365                   If total5 > 0 Then
366                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total5) * 100)
367                   Else
368                       Cetak = 0
369                   End If
370               End If
371               rsTempAll.filter = ""
372               If i = 1 Then Report.txtA15.SetText Cetak & " %"
373               If i = 2 Then Report.txtA25.SetText Cetak & " %"
374               If i = 3 Then Report.txtA35.SetText Cetak & " %"
375               If i = 4 Then Report.txtA45.SetText Cetak & " %"
376               If i = 5 Then Report.txtA55.SetText Cetak & " %"
377               If i = 6 Then Report.txtA65.SetText Cetak & " %"
378               If i = 7 Then Report.txtA75.SetText Cetak & " %"
379               If i = 8 Then Report.txtA85.SetText Cetak & " %"
380               If i = 9 Then Report.txtA95.SetText Cetak & " %"
381               If i = 10 Then Report.txtA105.SetText Cetak & " %"
382               If i = 11 Then Report.txtA115.SetText Cetak & " %"
383               If i = 12 Then Report.txtA125.SetText Cetak & " %"
384               DoEvents
                  
                  
385               rsTempAll.filter = "fs_rc = 'R060'"
386               If Not rsTempAll.EOF Then
387                   If total6 > 0 Then
388                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total6) * 100)
389                   Else
390                       Cetak = 0
391                   End If
392               End If
393               rsTempAll.filter = ""
394               If i = 1 Then Report.txtA16.SetText Cetak & " %"
395               If i = 2 Then Report.txtA26.SetText Cetak & " %"
396               If i = 3 Then Report.txtA36.SetText Cetak & " %"
397               If i = 4 Then Report.txtA46.SetText Cetak & " %"
398               If i = 5 Then Report.txtA56.SetText Cetak & " %"
399               If i = 6 Then Report.txtA66.SetText Cetak & " %"
400               If i = 7 Then Report.txtA76.SetText Cetak & " %"
401               If i = 8 Then Report.txtA86.SetText Cetak & " %"
402               If i = 9 Then Report.txtA96.SetText Cetak & " %"
403               If i = 10 Then Report.txtA106.SetText Cetak & " %"
404               If i = 11 Then Report.txtA116.SetText Cetak & " %"
405               If i = 12 Then Report.txtA126.SetText Cetak & " %"
406               DoEvents
                  
                  
407               rsTempAll.filter = "fs_rc = 'R070'"
408               If Not rsTempAll.EOF Then
409                   If total7 > 0 Then
410                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total7) * 100)
411                   Else
412                       Cetak = 0
413                   End If
414               End If
415               rsTempAll.filter = ""
416               If i = 1 Then Report.txtA17.SetText Cetak & " %"
417               If i = 2 Then Report.txtA27.SetText Cetak & " %"
418               If i = 3 Then Report.txtA37.SetText Cetak & " %"
419               If i = 4 Then Report.txtA47.SetText Cetak & " %"
420               If i = 5 Then Report.txtA57.SetText Cetak & " %"
421               If i = 6 Then Report.txtA67.SetText Cetak & " %"
422               If i = 7 Then Report.txtA77.SetText Cetak & " %"
423               If i = 8 Then Report.txtA87.SetText Cetak & " %"
424               If i = 9 Then Report.txtA97.SetText Cetak & " %"
425               If i = 10 Then Report.txtA107.SetText Cetak & " %"
426               If i = 11 Then Report.txtA117.SetText Cetak & " %"
427               If i = 12 Then Report.txtA127.SetText Cetak & " %"
428               DoEvents
                  
429               rsTempAll.filter = "fs_rc = 'R080'"
430               If Not rsTempAll.EOF Then
431                   If total8 > 0 Then
432                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total8) * 100)
433                   Else
434                       Cetak = 0
435                   End If
436               End If
437               rsTempAll.filter = ""
438               If i = 1 Then Report.txtA18.SetText Cetak & " %"
439               If i = 2 Then Report.txtA28.SetText Cetak & " %"
440               If i = 3 Then Report.txtA38.SetText Cetak & " %"
441               If i = 4 Then Report.txtA48.SetText Cetak & " %"
442               If i = 5 Then Report.txtA58.SetText Cetak & " %"
443               If i = 6 Then Report.txtA68.SetText Cetak & " %"
444               If i = 7 Then Report.txtA78.SetText Cetak & " %"
445               If i = 8 Then Report.txtA88.SetText Cetak & " %"
446               If i = 9 Then Report.txtA98.SetText Cetak & " %"
447               If i = 10 Then Report.txtA108.SetText Cetak & " %"
448               If i = 11 Then Report.txtA118.SetText Cetak & " %"
449               If i = 12 Then Report.txtA128.SetText Cetak & " %"
450               DoEvents
                  
                  
451               rsTempAll.filter = "fs_rc = 'R090'"
452               If Not rsTempAll.EOF Then
453                   If total9 > 0 Then
454                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total9) * 100)
455                   Else
456                       Cetak = 0
457                   End If
458               End If
459               rsTempAll.filter = ""
460               If i = 1 Then Report.txtA19.SetText Cetak & " %"
461               If i = 2 Then Report.txtA29.SetText Cetak & " %"
462               If i = 3 Then Report.txtA39.SetText Cetak & " %"
463               If i = 4 Then Report.txtA49.SetText Cetak & " %"
464               If i = 5 Then Report.txtA59.SetText Cetak & " %"
465               If i = 6 Then Report.txtA69.SetText Cetak & " %"
466               If i = 7 Then Report.txtA79.SetText Cetak & " %"
467               If i = 8 Then Report.txtA89.SetText Cetak & " %"
468               If i = 9 Then Report.txtA99.SetText Cetak & " %"
469               If i = 10 Then Report.txtA109.SetText Cetak & " %"
470               If i = 11 Then Report.txtA119.SetText Cetak & " %"
471               If i = 12 Then Report.txtA129.SetText Cetak & " %"
472               DoEvents
                  
                  
473               rsTempAll.filter = "fs_rc = 'R100'"
474               If Not rsTempAll.EOF Then
475                   If total10 > 0 Then
476                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total10) * 100)
477                   Else
478                       Cetak = 0
479                   End If
480               End If
481               rsTempAll.filter = ""
482               If i = 1 Then Report.txtA110.SetText Cetak & " %"
483               If i = 2 Then Report.txtA210.SetText Cetak & " %"
484               If i = 3 Then Report.txtA310.SetText Cetak & " %"
485               If i = 4 Then Report.txtA410.SetText Cetak & " %"
486               If i = 5 Then Report.txtA510.SetText Cetak & " %"
487               If i = 6 Then Report.txtA610.SetText Cetak & " %"
488               If i = 7 Then Report.txtA710.SetText Cetak & " %"
489               If i = 8 Then Report.txtA810.SetText Cetak & " %"
490               If i = 9 Then Report.txtA910.SetText Cetak & " %"
491               If i = 10 Then Report.txtA1010.SetText Cetak & " %"
492               If i = 11 Then Report.txtA1110.SetText Cetak & " %"
493               If i = 12 Then Report.txtA1210.SetText Cetak & " %"
494               DoEvents
495           End If
              
              
496           ProgressBar2.value = i
497           i = i + 1
498       Loop
499       Report.txtTahunSame.SetText TDBNumber1.value
500       Report.txtTahunAll.SetText TDBNumber1.value
          
501       With CRViewer1
502           .ReportSource = Report
503           .ViewReport
504           .Zoom 100
505           .Visible = True
506           .Left = 300
507           .Height = Me.Height - 400
508           .Width = Me.Width - .Left
509       End With
         
510               cmdTutup.Visible = True
511               cmdTutup.Left = CRViewer1.Left + 20
512               cmdTutup.Top = 435
                  
513       ProgressBar2.Visible = False
514       Frame2.Visible = False
515       Screen.MousePointer = vbDefault
          

516      On Error GoTo 0
517      Exit Sub

cmdRetension_Click_Error:

518       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdRetension_Click of Form frmInfoPelanggan")
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub cmdWarna_Click()
    frameWarna.Visible = True
    Text1.Text = "aktif"
    Option1(4).value = vbChecked
    
End Sub

Private Sub Combo1_Change()
    If Combo1.Text = "RANGE" Then
        framePasif1Range.Visible = True
        framePasifBulan.Visible = False
    Else
        framePasifBulan.Visible = True
        framePasif1Range.Visible = False
    End If
End Sub

Private Sub Combo1_Click()
    Call Combo1_Change
End Sub

Private Sub Combo2_Change()
    If Combo2.Text = "RANGE" Then
        FramePasif2Range.Visible = True
        FramePasif2Bulan.Visible = False
    Else
        FramePasif2Bulan.Visible = True
        FramePasif2Range.Visible = False
    End If
End Sub

Private Sub Combo2_Click()
    Call Combo2_Change
End Sub

Private Sub Combo3_Change()
    If Combo3.Text = "RANGE" Then
        FramePasif3Range.Visible = True
        FramePasif3Bulan.Visible = False
    Else
        FramePasif3Bulan.Visible = True
        FramePasif3Range.Visible = False
    End If
End Sub

Private Sub Combo3_Click()
    Call Combo3_Change
End Sub

Private Sub Command1_Click()
    If Text1.Text = "aktif" Then
        If Option1(0).value Then
            LabelAktif.Caption = Option1(0).Caption
            LabelAktif.BackColor = Option1(0).BackColor
        ElseIf Option1(1).value Then
            LabelAktif.Caption = Option1(1).Caption
            LabelAktif.BackColor = Option1(1).BackColor
        ElseIf Option1(2).value Then
            LabelAktif.Caption = Option1(2).Caption
            LabelAktif.BackColor = Option1(2).BackColor
        ElseIf Option1(3).value Then
            LabelAktif.Caption = Option1(3).Caption
            LabelAktif.BackColor = Option1(3).BackColor
        ElseIf Option1(4).value Then
            LabelAktif.Caption = "Putih" 'Option1(4).Caption
            LabelAktif.BackColor = Option1(4).BackColor
        End If
    ElseIf Text1.Text = "pasif1" Then
        If Option1(0).value Then
            LabelPasif.Caption = Option1(0).Caption
            LabelPasif.BackColor = Option1(0).BackColor
        ElseIf Option1(1).value Then
            LabelPasif.Caption = Option1(1).Caption
            LabelPasif.BackColor = Option1(1).BackColor
        ElseIf Option1(2).value Then
            LabelPasif.Caption = Option1(2).Caption
            LabelPasif.BackColor = Option1(2).BackColor
        ElseIf Option1(3).value Then
            LabelPasif.Caption = Option1(3).Caption
            LabelPasif.BackColor = Option1(3).BackColor
        ElseIf Option1(4).value Then
            LabelPasif.Caption = "Putih" 'Option1(4).Caption
            LabelPasif.BackColor = Option1(4).BackColor
        End If
    ElseIf Text1.Text = "pasif2" Then
        If Option1(0).value Then
            LabelPasif2.Caption = Option1(0).Caption
            LabelPasif2.BackColor = Option1(0).BackColor
        ElseIf Option1(1).value Then
            LabelPasif2.Caption = Option1(1).Caption
            LabelPasif2.BackColor = Option1(1).BackColor
        ElseIf Option1(2).value Then
            LabelPasif2.Caption = Option1(2).Caption
            LabelPasif2.BackColor = Option1(2).BackColor
        ElseIf Option1(3).value Then
            LabelPasif2.Caption = Option1(3).Caption
            LabelPasif2.BackColor = Option1(3).BackColor
        ElseIf Option1(4).value Then
            LabelPasif2.Caption = "Putih" 'Option1(4).Caption
            LabelPasif2.BackColor = Option1(4).BackColor
        End If
    ElseIf Text1.Text = "pasif3" Then
        If Option1(0).value Then
            LabelPasif3.Caption = Option1(0).Caption
            LabelPasif3.BackColor = Option1(0).BackColor
        ElseIf Option1(1).value Then
            LabelPasif3.Caption = Option1(1).Caption
            LabelPasif3.BackColor = Option1(1).BackColor
        ElseIf Option1(2).value Then
            LabelPasif3.Caption = Option1(2).Caption
            LabelPasif3.BackColor = Option1(2).BackColor
        ElseIf Option1(3).value Then
            LabelPasif3.Caption = Option1(3).Caption
            LabelPasif3.BackColor = Option1(3).BackColor
        ElseIf Option1(4).value Then
            LabelPasif3.Caption = "Putih" 'Option1(4).Caption
            LabelPasif3.BackColor = Option1(4).BackColor
        End If
    End If
    frameWarna.Visible = False
End Sub

Private Sub Command2_Click()
    frameWarna.Visible = True
    Text1.Text = "pasif1"
    Option1(4).value = vbChecked
End Sub

Private Sub Command3_Click()
    frameWarna.Visible = True
    Text1.Text = "pasif3"
    Option1(4).value = vbChecked
End Sub

Private Sub Command4_Click()
    frameWarna.Visible = True
    Text1.Text = "pasif2"
    Option1(4).value = vbChecked
End Sub

Private Sub Command5_Click()
    Frame1.Visible = True
    Dim rsChart As New ADODB.Recordset
    If rsChart.State Then rsChart.Close
    Set rsChart = Nothing
    With rsChart
        'If .State <> adStateClosed Then .Close
        .Fields.Append "JUMLAH", adInteger
        .Open
    End With
 '   If ckAktif.value = vbChecked Then
        With rsChart
            .AddNew
            !JUMLAH = jumAktif
            .Update
        End With
   ' End If
   ' If ckPasif1.value = vbChecked Then
        With rsChart
            .AddNew
            !JUMLAH = jumPasif1
            .Update
        End With
   ' End If
   ' If ckPasif2.value = vbChecked Then
        With rsChart
            .AddNew
            !JUMLAH = jumPasif2
            .Update
        End With
   ' End If
   ' If ckPasif3.value = vbChecked Then
        With rsChart
            .AddNew
            !JUMLAH = jumPasif3
            .Update
        End With
   ' End If
    rsChart.MoveFirst
    
    MSChart1.ShowWhatsThis
    MSChart1.SeriesType = VtChSeriesType2dPie
    With MSChart1
        For i = 1 To .Plot.SeriesCollection.Count
            With .Plot.SeriesCollection(i).DataPoints(-1).DataPointLabel
                .LocationType = VtChLabelLocationTypeInside
                
                .Component = VtChLabelComponent.VtChLabelComponentValue
                .PercentFormat = "0%"
                .VtFont.Size = 10
'                .LocationType = VtChLabelLocationTypeInside
'                .Component = VtChLabelComponentValue '& " " & VtChLabelComponentPercent
                
            End With
            If i = 1 Then
                .Plot.SeriesCollection(i).LegendText = "Kategori Aktif"
            ElseIf i = 2 Then
                .Plot.SeriesCollection(i).LegendText = "Kategori Pasif1"
            ElseIf i = 3 Then
                .Plot.SeriesCollection(i).LegendText = "Kategori Pasif2"
            ElseIf i = 4 Then
                .Plot.SeriesCollection(i).LegendText = "Kategori Pasif3"
            End If
            
        Next i
    End With
    Set MSChart1.DataSource = rsChart
End Sub

Private Sub Command6_Click()
    Frame1.Visible = False
End Sub

Public Function getLastGRService(kriteria As String)
    
End Function

Public Function GetService(kriteria As String, Optional isHO As Boolean = False)
    rsCari.filter = ""
    rsCari.filter = "fs_rc = '" & kriteria & "'"
    If Not rsCari.EOF Then
        If isHO Then
            GetService = Format(rsCari("invdate"), "DD/MM/YYYY") '' & vbCrLf & Trim(rsCari("Fs_nm_dealer"))
        Else
            GetService = Format(rsCari("invdate"), "DD/MM/YYYY")
        End If
        
        Exit Function
    End If
    GetService = ""
End Function

Public Function GetServiceDealer(kriteria As String, Optional isHO As Boolean = False)
    rsCari.filter = ""
    rsCari.filter = "fs_rc = '" & kriteria & "'"
    If Not rsCari.EOF Then
        If isHO Then
            GetServiceDealer = Trim(rsCari("Fs_nm_dealer"))
         
        End If
        
        Exit Function
    End If
    GetServiceDealer = ""
End Function

Public Function GetServiceLast(kriteria As String, isHO As Boolean, isRep As Boolean, isLastServ As Boolean)
    rsCari.filter = ""
    rsCari.filter = kriteria
    If isHO Then
        If Not rsCari.EOF Then
            If isRep Then
                GetServiceLast = Trim(rsCari("rep"))
            ElseIf isLastServ Then
                GetServiceLast = Trim(rsCari("fs_nm_dealer"))
            Else
                GetServiceLast = Format(rsCari("invdate"), "DD/MM/YYYY") ' & vbCrLf & Trim(rsCari("fs_nm_dealer"))
            End If
            
            Exit Function
        End If
    Else
        If Not rsCari.EOF Then
            If isRep Then
                GetServiceLast = Trim(rsCari("rep"))
            ElseIf isLastServ Then
                GetServiceLast = Trim(rsCari("fs_nm_dealer"))
            Else
                GetServiceLast = Format(rsCari("invdate"), "DD/MM/YYYY")
            End If
            
            Exit Function
        End If
    End If
    
    GetServiceLast = ""
End Function

Private Sub RefreshGridMapping()
      Dim rsTagih As ADODB.Recordset
      Dim rsDataHO As ADODB.Recordset

      Dim varTagih As String
      Dim kriteriaLastService As String
      Dim kriteria As String
      Dim sSQL As String
      Dim isSBE As Boolean
      Dim konter As Double
      Dim daftarNOChassis As String
      Dim warna As String

1        On Error GoTo RefreshGridMapping_Error

2     jumAktif = 0
3     jumPasif1 = 0
4     jumPasif2 = 0
5     jumPasif3 = 0

6     isSBE = False
7     On Error GoTo vb_error
8         Screen.MousePointer = vbHourglass
9         kriteria = ""
          'pengecekan last service berdasarkan tanggal
10        If (IIf(IsNull(Tgl1), "", Tgl1) <> "") Then
11            kriteria = kriteria & " (J.[DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "') AND "
12        End If
          'pengecekan last sbe berdasarkan tanggal
13        If (IIf(IsNull(Tgl3), "", Tgl3) <> "") Then
14            isSBE = True
15            kriteria = kriteria & " ((J.[DATE] BETWEEN '" & Format(Tgl3, "MM/DD/YYYY") & "' AND '" & Format(Tgl4, "MM/DD/YYYY") & "') AND FS_RC <> '') and "
16        End If
          '' pengecekan ddate berdasarkan tanggal
17        If (IIf(IsNull(tglDDate2), "", tglDDate2) <> "") Then

18            kriteria = kriteria & " ((CONVERT(datetime,RIGHT(c.ddate,4)+SUBSTRING(c.ddate,4,2)+LEFT(c.ddate,2)) BETWEEN '" & Format(tglDDate1, "MM/DD/YYYY") & "' AND '" & Format(tglDDate2, "MM/DD/YYYY") & "')) and "
19        End If
          
          'pengecekan alamat customer
20        If Trim(TxtAlamat.Text) <> "" Then
21            kriteria = kriteria & " C.CUST_ADRES + C.KOTA LIKE '%" & Trim(TxtAlamat.Text) & "%' AND "
22        End If
          
23        If blnAkhir > 0 And blnAwal > 0 And (blnAkhir >= blnAwal) Then
24            kriteria = kriteria & " (DATEDIFF(MONTH,J.DATE,GETDATE())) BETWEEN " & blnAwal & " AND " & blnAkhir & " AND "
25        End If
26        If blnPasif > 0 Then
27            kriteria = kriteria & " (DATEDIFF(MONTH,J.DATE,GETDATE())) " & IIf(cbOperator.Text = "", "=", cbOperator.Text) & blnPasif & " AND "
28        End If
29        If Trim(kriteria) <> "" Then
30            kriteria = Left(kriteria, Len(kriteria) - 4)
31            kriteria = " WHERE " & kriteria
32        End If
33        varTagih = " and isnull(notagih,'') = ''"
34        If nTagih Then
35            varTagih = " or isnull(notagih,'') <> ''"
36        End If
37        jmlTagih = 0
          'If Check1.value = vbChecked Then
              'MENGAMBIL NO_CHASSIS LOCAL
38            sSQL = "SELECT DISTINCT C.NO_CHASSIS,C.MDL,C.FMDL,C.NO_POLISI,C.FS_TH_MOBIL,C.CUSTOMER,C.CUST_ADRES,"
39            sSQL = sSQL & " C.KOTA,C.TELPON,C.USENAME,C.USEGSM,C.FS_NM_DEALER, C.Ddate,C.CDMA,C.GSM FROM CUSTDAT C WITH (NOLOCK) LEFT OUTER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_CHASSIS = C.NO_CHASSIS " & kriteria
              
40        lblPengambilanData.Visible = True
41        DoEvents
42        Set rsTagih = objO.FetchData(sSQL)
          
43        If rsTagih.RecordCount > 0 Then
44            Call createRS
45            lblPengambilanData.Visible = False
46            ProgressBar1.Visible = True
47            rsTagih.MoveFirst
48            ProgressBar1.Max = rsTagih.RecordCount + 1
49            ProgressBar1.Min = 0
50            Do While Not rsTagih.EOF
                  'rsDataHO.filter = ""
51                If Check1.value = vbChecked Then
                      ''@@@@ 2016-05-18 || pengambilan data dari t_invoice
                      ''Set rsCari = objO.FetchDataHOUnlimitedTime("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,fs_nm_dealer,KM,INVDATE,NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' order by INVDATE desc")
52                    If Log.RPLHistori = 0 Then
53                        Set rsCari = objO.FetchDataHOUnlimitedTime("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,fs_nm_dealer,KM,INVDATE,NO_CHASSIS FROM T_INVOICENEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B'  order by INVDATE desc")
54                    Else
55                        Set rsCari = objO.FetchData("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,fs_nm_dealer,KM,INVDATE,NO_CHASSIS FROM [" & dbReplikasiHO & "]..T_INVOICENEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B'  order by INVDATE desc")
56                    End If
57                Else
                      ''---- Set rsCari = objO.FetchData("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,KM,INVDATE,fs_nm_dealer,NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' order by INVDATE desc")
                      ''@@@@ 2016-05-18 || pengambilan data dari t_invoice
58                    Set rsCari = objO.FetchData("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,KM,INVDATE,fs_nm_dealer,NO_CHASSIS FROM T_INVOICE WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B' order by INVDATE desc")
                      
59                End If
             '     rsCari.filter
60                jumlahLastGr = 0
61                jumlanLastSBE = 0
62                jumlahService = 0
63                If Not rsCari.EOF Then
64                    jumlahService = CInt(rsCari("selisih"))
65                    rsCari.filter = "fs_rc = ''"
66                    If Not rsCari.EOF Then jumlahLastGr = CInt(rsCari("selisih"))
67                    rsCari.filter = ""
68                    rsCari.filter = "fs_rc <> ''"
69                    If Not rsCari.EOF Then jumlanLastSBE = CInt(rsCari("selisih"))
70                End If
71                rsCari.filter = ""
72                With rsRecord
73                    .AddNew
74                    !NO_CHASSIS = Trim(rsTagih("NO_CHASSIS"))
75                    !MDL = Trim(rsTagih("MDL"))
76                    !Fmdl = Trim(rsTagih("FMDL"))
77                    !NO_POLISI = Trim(rsTagih("NO_POLISI"))
78                    !Tahun = Trim(rsTagih("FS_TH_MOBIL"))
79                    !customer = Trim(rsTagih("CUSTOMER"))
80                    !cust_adres = Trim(rsTagih("CUST_ADRES"))
81                    !Kota = Trim(rsTagih("KOTA"))
82                    !telp = Trim(rsTagih("TELPON"))
83                    !UseName = Trim(rsTagih("USENAME"))
84                    !UseGSM = Trim(rsTagih("USEGSM"))
85                    !GSM = Trim(rsTagih("GSM"))
86                    !CDMA = Trim(rsTagih("CDMA"))
87                    !dDate = Trim(rsTagih("ddate"))
88                    !DealerPenjual = IIf(IsNull(Trim((rsTagih("fs_nm_dealer")))), "", Trim(rsTagih("fs_nm_dealer"))) 'Trim(rsTagih("fs_nm_dealer"))
89                    If Check1.value = vbChecked Then
90                        rsCari.filter = ""
                          'sbi
91                        !K1 = GetService("R001", True) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R001' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
92                        !dealerK1 = GetServiceDealer("R001", True)
                          'sbe
93                        !k5 = GetService("R005", True) 'IIf(IsNull(Trim(rsTagih("5K"))), "", Trim(rsTagih("5K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R005' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
94                        !dealerK5 = GetServiceDealer("R005", True)
95                        !K10 = GetService("R010", True) 'IIf(IsNull(Trim((rsTagih("10K")))), "", Trim(rsTagih("10K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R010' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
96                        !dealerK10 = GetServiceDealer("R010", True)
97                        !k15 = GetService("R015", True) 'IIf(IsNull(Trim((rsTagih("15K")))), "", Trim(rsTagih("15K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R015' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
98                        !dealerK15 = GetServiceDealer("R015", True)
99                        !K20 = GetService("R020", True) 'IIf(IsNull(Trim((rsTagih("20K")))), "", Trim(rsTagih("20K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R020' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
100                       !dealerK20 = GetServiceDealer("R020", True)
101                       !k25 = GetService("R025", True) 'IIf(IsNull(Trim((rsTagih("25K")))), "", Trim(rsTagih("25K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R025' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
102                       !dealerK25 = GetServiceDealer("R025", True)
103                       !K30 = GetService("R030", True) 'IIf(IsNull(Trim((rsTagih("30K")))), "", Trim(rsTagih("30K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R030' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
104                       !dealerK30 = GetServiceDealer("R030", True)
105                       !k35 = GetService("R035", True) 'IIf(IsNull(Trim((rsTagih("35K")))), "", Trim(rsTagih("35K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R035' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
106                       !dealerK35 = GetServiceDealer("R035", True)
107                       !K40 = GetService("R040", True) 'IIf(IsNull(Trim((rsTagih("40K")))), "", Trim(rsTagih("40K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R040' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
108                       !dealerK40 = GetServiceDealer("R040", True)
109                       !k45 = GetService("R045", True) 'IIf(IsNull(Trim((rsTagih("45K")))), "", Trim(rsTagih("45K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R045' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
110                       !dealerK45 = GetServiceDealer("R045", True)
111                       !K50 = GetService("R050", True) 'IIf(IsNull(Trim((rsTagih("50K")))), "", Trim(rsTagih("50K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R050' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
112                       !dealerK50 = GetServiceDealer("R050", True)
113                       !k55 = GetService("R055", True) 'IIf(IsNull(Trim((rsTagih("55K")))), "", Trim(rsTagih("55K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R055' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
114                       !dealerK55 = GetServiceDealer("R055", True)
115                       !K60 = GetService("R060", True) 'IIf(IsNull(Trim((rsTagih("60K")))), "", Trim(rsTagih("60K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R060' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
116                       !dealerK60 = GetServiceDealer("R060", True)
117                       !k65 = GetService("R065", True) 'IIf(IsNull(Trim((rsTagih("65K")))), "", Trim(rsTagih("65K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R065' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
118                       !dealerK65 = GetServiceDealer("R065", True)
119                       !K70 = GetService("R070", True) 'IIf(IsNull(Trim((rsTagih("70K")))), "", Trim(rsTagih("70K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R070' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
120                       !dealerK70 = GetServiceDealer("R070", True)
121                       !k75 = GetService("R075", True) 'IIf(IsNull(Trim((rsTagih("75K")))), "", Trim(rsTagih("75K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R075' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
122                       !dealerK75 = GetServiceDealer("R075", True)
123                       !K80 = GetService("R080", True) 'IIf(IsNull(Trim((rsTagih("80K")))), "", Trim(rsTagih("80K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R080' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
124                       !dealerK80 = GetServiceDealer("R080", True)
125                       !k85 = GetService("R085", True) 'IIf(IsNull(Trim((rsTagih("85K")))), "", Trim(rsTagih("85K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R085' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
126                       !dealerK85 = GetServiceDealer("R085", True)
127                       !k90 = GetService("R090", True) 'IIf(IsNull(Trim((rsTagih("90K")))), "", Trim(rsTagih("90K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R090' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
128                       !dealerK90 = GetServiceDealer("R090", True)
129                       !k95 = GetService("R095", True) 'IIf(IsNull(Trim((rsTagih("95K")))), "", Trim(rsTagih("95K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R095' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
130                       !dealerK95 = GetServiceDealer("R095", True)
131                       !k100 = GetService("R100", True) 'IIf(IsNull(Trim((rsTagih("100K")))), "", Trim(rsTagih("100K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R100' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
132                       !dealerK100 = GetServiceDealer("R100", True)
133                       !k105 = GetService("R105", True) 'IIf(IsNull(Trim((rsTagih("105K")))), "", Trim(rsTagih("105K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R105' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
134                       !dealerK105 = GetServiceDealer("R105", True)
135                       !k110 = GetService("R110", True) 'IIf(IsNull(Trim((rsTagih("110K")))), "", Trim(rsTagih("110K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R110' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
136                       !dealerK110 = GetServiceDealer("R110", True)
137                       !k115 = GetService("R115", True) 'IIf(IsNull(Trim((rsTagih("115K")))), "", Trim(rsTagih("115K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R115' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
138                       !dealerK115 = GetServiceDealer("R115", True)
139                       !k120 = GetService("R120", True) 'IIf(IsNull(Trim((rsTagih("120K")))), "", Trim(rsTagih("120K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R120' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
140                       !dealerK120 = GetServiceDealer("R120", True)
141                       !k125 = GetService("R125", True) 'IIf(IsNull(Trim((rsTagih("125K")))), "", Trim(rsTagih("125K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R125' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
142                       !dealerK125 = GetServiceDealer("R125", True)
143                       !k130 = GetService("R130", True) 'IIf(IsNull(Trim((rsTagih("130K")))), "", Trim(rsTagih("130K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R130' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
144                       !dealerK130 = GetServiceDealer("R130", True)
145                       !k135 = GetService("R135", True) 'IIf(IsNull(Trim((rsTagih("135K")))), "", Trim(rsTagih("135K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R135' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
146                       !dealerK135 = GetServiceDealer("R135", True)
147                       !k140 = GetService("R140", True) 'IIf(IsNull(Trim((rsTagih("140K")))), "", Trim(rsTagih("140K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R140' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
148                       !dealerK140 = GetServiceDealer("R140", True)
149                       !k145 = GetService("R145", True) 'IIf(IsNull(Trim((rsTagih("145K")))), "", Trim(rsTagih("145K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R145' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
150                       !dealerK145 = GetServiceDealer("R145", True)
151                       !k150 = GetService("R150", True) 'IIf(IsNull(Trim((rsTagih("150K")))), "", Trim(rsTagih("150K"))) 'ObjO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R150' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
152                       !dealerK150 = GetServiceDealer("R150", True)
153                       !k155 = GetService("R155", True) 'IIf(IsNull(Trim((rsTagih("155K")))), "", Trim(rsTagih("155K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R155' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
154                       !dealerK155 = GetServiceDealer("R155", True)
155                       !k160 = GetService("R160", True) 'IIf(IsNull(Trim((rsTagih("160K")))), "", Trim(rsTagih("160K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R160' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
156                       !dealerK160 = GetServiceDealer("R160", True)
157                       !k165 = GetService("R165", True) 'IIf(IsNull(Trim((rsTagih("165K")))), "", Trim(rsTagih("165K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R165' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
158                       !dealerK165 = GetServiceDealer("R165", True)
159                       !k170 = GetService("R170", True) 'IIf(IsNull(Trim((rsTagih("170K")))), "", Trim(rsTagih("170K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R170' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
160                       !dealerK170 = GetServiceDealer("R170", True)
161                       !k175 = GetService("R175", True) 'IIf(IsNull(Trim((rsTagih("175K")))), "", Trim(rsTagih("175K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R175' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
162                       !dealerK175 = GetServiceDealer("R175", True)
163                       !k180 = GetService("R180", True) 'IIf(IsNull(Trim((rsTagih("180K")))), "", Trim(rsTagih("180K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R180' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
164                       !dealerK180 = GetServiceDealer("R180", True)
165                       !k185 = GetService("R185", True) 'IIf(IsNull(Trim((rsTagih("185K")))), "", Trim(rsTagih("185K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R185' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
166                       !dealerK185 = GetServiceDealer("R185", True)
167                       !k190 = GetService("R190", True) 'IIf(IsNull(Trim((rsTagih("190K")))), "", Trim(rsTagih("190K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R190' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
168                       !dealerK190 = GetServiceDealer("R190", True)
169                       !k195 = GetService("R195", True) 'IIf(IsNull(Trim((rsTagih("195K")))), "", Trim(rsTagih("195K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R195' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
170                       !dealerK195 = GetServiceDealer("R195", True)
171                       !k200 = GetService("R200", True) 'IIf(IsNull(Trim((rsTagih("200K")))), "", Trim(rsTagih("200K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R200' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
172                       !dealerK200 = GetServiceDealer("R200", True)
173                       !lastSB = GetServiceLast("fs_rc <> ''", True, False, False) 'IIf(IsNull(Trim((rsTagih("lastSB")))), "", Trim(rsTagih("lastSB"))) 'Trim(rsTagih("lastSB"))
174                       !dealerLastSB = GetServiceLast("fs_rc <> ''", True, False, True)
175                       !lastGR = GetServiceLast("fs_rc = ''", True, False, False) 'IIf(IsNull(Trim((rsTagih("lastGR")))), "", Trim(rsTagih("lastGR"))) 'Trim(rsTagih("lastGR"))
176                       !dealerLastGR = GetServiceLast("fs_rc = ''", True, False, True)
177                       !lastService = GetServiceLast("", True, False, False) 'IIf(IsNull(Trim((rsTagih("lastSVC")))), "", Trim(rsTagih("lastSVC"))) 'Trim(rsTagih("lastSVC"))
178                       !lastInfoGR = GetServiceLast("", True, True, False) 'IIf(IsNull(Trim((rsTagih("jobGR")))), "", Trim(rsTagih("jobGR"))) 'Trim(rsTagih("jobGR"))
179                       !DealerLastSVC = GetServiceLast("", True, False, True)

180                   Else
181                       rsCari.filter = ""
                          'sbi
182                       !K1 = GetService("R001")  'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R001' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK1 = GetServiceDealer("R001", True)
                          'sbe
183                       !k5 = GetService("R005") 'IIf(IsNull(Trim(rsTagih("5K"))), "", Trim(rsTagih("5K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R005' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK5 = GetServiceDealer("R005", True)
184                       !K10 = GetService("R010") 'IIf(IsNull(Trim((rsTagih("10K")))), "", Trim(rsTagih("10K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R010' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK10 = GetServiceDealer("R010", True)
185                       !k15 = GetService("R015") 'IIf(IsNull(Trim((rsTagih("15K")))), "", Trim(rsTagih("15K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R015' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK15 = GetServiceDealer("R015", True)
186                       !K20 = GetService("R020") 'IIf(IsNull(Trim((rsTagih("20K")))), "", Trim(rsTagih("20K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R020' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK20 = GetServiceDealer("R020", True)
187                       !k25 = GetService("R025") 'IIf(IsNull(Trim((rsTagih("25K")))), "", Trim(rsTagih("25K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R025' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK25 = GetServiceDealer("R025", True)
188                       !K30 = GetService("R030") 'IIf(IsNull(Trim((rsTagih("30K")))), "", Trim(rsTagih("30K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R030' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK30 = GetServiceDealer("R030", True)
189                       !k35 = GetService("R035") 'IIf(IsNull(Trim((rsTagih("35K")))), "", Trim(rsTagih("35K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R035' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK35 = GetServiceDealer("R035", True)
190                       !K40 = GetService("R040") 'IIf(IsNull(Trim((rsTagih("40K")))), "", Trim(rsTagih("40K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R040' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK40 = GetServiceDealer("R040", True)
191                       !k45 = GetService("R045") 'IIf(IsNull(Trim((rsTagih("45K")))), "", Trim(rsTagih("45K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R045' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK45 = GetServiceDealer("R045", True)
192                       !K50 = GetService("R050") 'IIf(IsNull(Trim((rsTagih("50K")))), "", Trim(rsTagih("50K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R050' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK50 = GetServiceDealer("R050", True)
193                       !k55 = GetService("R055") 'IIf(IsNull(Trim((rsTagih("55K")))), "", Trim(rsTagih("55K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R055' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK55 = GetServiceDealer("R055", True)
194                       !K60 = GetService("R060") 'IIf(IsNull(Trim((rsTagih("60K")))), "", Trim(rsTagih("60K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R060' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK60 = GetServiceDealer("R060", True)
195                       !k65 = GetService("R065") 'IIf(IsNull(Trim((rsTagih("65K")))), "", Trim(rsTagih("65K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R065' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK65 = GetServiceDealer("R065", True)
196                       !K70 = GetService("R070") 'IIf(IsNull(Trim((rsTagih("70K")))), "", Trim(rsTagih("70K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R070' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                        '  !dealerK70 = GetServiceDealer("R070", True)
197                       !k75 = GetService("R075") 'IIf(IsNull(Trim((rsTagih("75K")))), "", Trim(rsTagih("75K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R075' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK75 = GetServiceDealer("R075", True)
198                       !K80 = GetService("R080") 'IIf(IsNull(Trim((rsTagih("80K")))), "", Trim(rsTagih("80K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R080' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK80 = GetServiceDealer("R080", True)
199                       !k85 = GetService("R085") 'IIf(IsNull(Trim((rsTagih("85K")))), "", Trim(rsTagih("85K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R085' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK85 = GetServiceDealer("R085", True)
200                       !k90 = GetService("R090") 'IIf(IsNull(Trim((rsTagih("90K")))), "", Trim(rsTagih("90K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R090' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK90 = GetServiceDealer("R090", True)
201                       !k95 = GetService("R095") 'IIf(IsNull(Trim((rsTagih("95K")))), "", Trim(rsTagih("95K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R095' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK95 = GetServiceDealer("R095", True)
202                       !k100 = GetService("R100") 'IIf(IsNull(Trim((rsTagih("100K")))), "", Trim(rsTagih("100K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R100' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK100 = GetServiceDealer("R100", True)
203                       !k105 = GetService("R105") 'IIf(IsNull(Trim((rsTagih("105K")))), "", Trim(rsTagih("105K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R105' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK105 = GetServiceDealer("R105", True)
204                       !k110 = GetService("R110") 'IIf(IsNull(Trim((rsTagih("110K")))), "", Trim(rsTagih("110K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R110' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK110 = GetServiceDealer("R110", True)
205                       !k115 = GetService("R115") 'IIf(IsNull(Trim((rsTagih("115K")))), "", Trim(rsTagih("115K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R115' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK115 = GetServiceDealer("R115", True)
206                       !k120 = GetService("R120") 'IIf(IsNull(Trim((rsTagih("120K")))), "", Trim(rsTagih("120K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R120' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK120 = GetServiceDealer("R120", True)
207                       !k125 = GetService("R125") 'IIf(IsNull(Trim((rsTagih("125K")))), "", Trim(rsTagih("125K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R125' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK125 = GetServiceDealer("R125", True)
208                       !k130 = GetService("R130") 'IIf(IsNull(Trim((rsTagih("130K")))), "", Trim(rsTagih("130K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R130' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK130 = GetServiceDealer("R130", True)
209                       !k135 = GetService("R135") 'IIf(IsNull(Trim((rsTagih("135K")))), "", Trim(rsTagih("135K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R135' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK135 = GetServiceDealer("R135", True)
210                       !k140 = GetService("R140") 'IIf(IsNull(Trim((rsTagih("140K")))), "", Trim(rsTagih("140K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R140' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK140 = GetServiceDealer("R140", True)
211                       !k145 = GetService("R145") 'IIf(IsNull(Trim((rsTagih("145K")))), "", Trim(rsTagih("145K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R145' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK145 = GetServiceDealer("R145", True)
212                       !k150 = GetService("R150") 'IIf(IsNull(Trim((rsTagih("150K")))), "", Trim(rsTagih("150K"))) 'ObjO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R150' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK150 = GetServiceDealer("R150", True)
213                       !k155 = GetService("R155") 'IIf(IsNull(Trim((rsTagih("155K")))), "", Trim(rsTagih("155K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R155' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK155 = GetServiceDealer("R155", True)
214                       !k160 = GetService("R160") 'IIf(IsNull(Trim((rsTagih("160K")))), "", Trim(rsTagih("160K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R160' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK160 = GetServiceDealer("R160", True)
215                       !k165 = GetService("R165") 'IIf(IsNull(Trim((rsTagih("165K")))), "", Trim(rsTagih("165K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R165' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK165 = GetServiceDealer("R165", True)
216                       !k170 = GetService("R170") 'IIf(IsNull(Trim((rsTagih("170K")))), "", Trim(rsTagih("170K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R170' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK170 = GetServiceDealer("R170", True)
217                       !k175 = GetService("R175") 'IIf(IsNull(Trim((rsTagih("175K")))), "", Trim(rsTagih("175K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R175' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK175 = GetServiceDealer("R175", True)
218                       !k180 = GetService("R180") 'IIf(IsNull(Trim((rsTagih("180K")))), "", Trim(rsTagih("180K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R180' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK180 = GetServiceDealer("R180", True)
219                       !k185 = GetService("R185") 'IIf(IsNull(Trim((rsTagih("185K")))), "", Trim(rsTagih("185K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R185' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK185 = GetServiceDealer("R185", True)
220                       !k190 = GetService("R190") 'IIf(IsNull(Trim((rsTagih("190K")))), "", Trim(rsTagih("190K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R190' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK190 = GetServiceDealer("R190", True)
221                       !k195 = GetService("R195") 'IIf(IsNull(Trim((rsTagih("195K")))), "", Trim(rsTagih("195K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R195' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                        '  !dealerK195 = GetServiceDealer("R195", True)
222                       !k200 = GetService("R200") 'IIf(IsNull(Trim((rsTagih("200K")))), "", Trim(rsTagih("200K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R200' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK200 = GetServiceDealer("R200", True)
223                       !lastSB = GetServiceLast("fs_rc <> ''", False, False, False) 'IIf(IsNull(Trim((rsTagih("lastSB")))), "", Trim(rsTagih("lastSB"))) 'Trim(rsTagih("lastSB"))
224                       !dealerLastSB = GetServiceLast("fs_rc <> ''", False, False, True)
225                       !lastGR = GetServiceLast("fs_rc = ''", False, False, False) 'IIf(IsNull(Trim((rsTagih("lastGR")))), "", Trim(rsTagih("lastGR"))) 'Trim(rsTagih("lastGR"))
226                       !dealerLastGR = GetServiceLast("fs_rc = ''", False, False, True)
227                       !lastService = GetServiceLast("", False, False, False) 'IIf(IsNull(Trim((rsTagih("lastSVC")))), "", Trim(rsTagih("lastSVC"))) 'Trim(rsTagih("lastSVC"))
228                       !lastInfoGR = GetServiceLast("", False, True, False) 'IIf(IsNull(Trim((rsTagih("jobGR")))), "", Trim(rsTagih("jobGR"))) 'Trim(rsTagih("jobGR"))
229                       !DealerLastSVC = GetServiceLast("", False, False, True) 'IIf(IsNull(Trim((rsTagih("dealerlastsvc")))), "", Trim(rsTagih("dealerlastsvc")))
230                   End If
231                   !selisih = jumlahService
232                   warna = getWarnaUnit(jumlahService)
233                   !warna = warna
234                   If warna <> "" Then
235                     If UCase(warna) = UCase(LabelAktif.Caption) Then
236                        warna = "Aktif"
237                     ElseIf UCase(warna) = UCase(LabelPasif.Caption) Then
238                        warna = "Pasif1"
239                     ElseIf UCase(warna) = UCase(LabelPasif2.Caption) Then
240                        warna = "Pasif2"
241                     ElseIf UCase(warna) = UCase(LabelPasif3.Caption) Then
242                        warna = "Pasif3"
243                     Else
244                        warna = ""
245                     End If
246                   Else
247                     warna = ""
248                   End If
249                     !aktif = warna
250                   .Update
251               End With
252               ProgressBar1.value = konter
253               DoEvents
254               LblJumlah.Caption = konter
255               konter = konter + 1
256               rsTagih.MoveNext
257           Loop
              
258       Else
              
259           Call createRS
260       End If
          
261       gridTagih.FetchRowStyle = True
262       lblPengambilanData.Visible = False
263       ProgressBar1.Visible = False
264       LblJumlah.Caption = rsRecord.RecordCount
265       Screen.MousePointer = vbDefault
266       Exit Sub
vb_error:
267       MsgBox "Terjadi kesalahan pengambilan data." & vbCrLf & "Error Number = " & err.Number & vbCrLf & err.Description & vbCrLf & "Procedure : " & Me.Name & ".RefreshGrid"
268       lblPengambilanData.Visible = False
269       ProgressBar1.Visible = False
270       LblJumlah.Caption = rsRecord.RecordCount
271       Screen.MousePointer = vbDefault

272      On Error GoTo 0
273      Exit Sub

RefreshGridMapping_Error:

274       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure RefreshGridMapping of Form frmInfoPelanggan")
          
End Sub

Private Sub Command9_Click()
    Frame2.Visible = False
End Sub

Private Sub Form_Load()
  '  FrameAktifRange.Visible = False
    ckAktif.value = vbChecked
    ckPasif1.value = vbChecked
    ckPasif2.value = vbChecked
    ckPasif3.value = vbChecked
    
    cbOperator.Text = cbOperator.LIST(1)
    Combo1.Text = Combo1.LIST(3)
    Combo2.Text = Combo2.LIST(3)
    Combo3.Text = Combo3.LIST(0)
    cbBulanRet.Text = Month(isnow)
    TDBNumber1.value = year(isnow)
End Sub


Private Sub createRS()
1         If rsRecord.State Then rsRecord.Close
2         Set rsRecord = Nothing
3         With rsRecord
              'If .State <> adStateClosed Then .Close
4             .Fields.Append "NO_CHASSIS", adVarChar, 100
5             .Fields.Append "MDL", adVarChar, 100
6             .Fields.Append "FMDL", adVarChar, 100
7             .Fields.Append "NO_POLISI", adVarChar, 40
8             .Fields.Append "TAHUN", adVarChar, 10
9             .Fields.Append "CUSTOMER", adVarChar, 200
10            .Fields.Append "CUST_ADRES", adVarChar, 300
11            .Fields.Append "KOTA", adVarChar, 100
12            .Fields.Append "TELP", adVarChar, 100
13            .Fields.Append "USENAME", adVarChar, 100
14            .Fields.Append "USEGSM", adVarChar, 100
              ''## 2016-06-28 || CDMA
15            .Fields.Append "CDMA", adVarChar, 100
16            .Fields.Append "GSM", adVarChar, 100
              ''## 2016-05-26 || DELIVERY DATE --- ##
17            .Fields.Append "DDate", adVarChar, 100
18            .Fields.Append "k1", adVarChar, 100
19            .Fields.Append "k5", adVarChar, 100
20            .Fields.Append "k10", adVarChar, 100
21            .Fields.Append "k15", adVarChar, 100
22            .Fields.Append "k20", adVarChar, 100
23            .Fields.Append "k25", adVarChar, 100
24            .Fields.Append "k30", adVarChar, 100
25            .Fields.Append "k35", adVarChar, 100
26            .Fields.Append "k40", adVarChar, 100
27            .Fields.Append "k45", adVarChar, 100
28            .Fields.Append "k50", adVarChar, 100
29            .Fields.Append "k55", adVarChar, 100
30            .Fields.Append "k60", adVarChar, 100
31            .Fields.Append "k65", adVarChar, 100
32            .Fields.Append "k70", adVarChar, 100
33            .Fields.Append "k75", adVarChar, 100
34            .Fields.Append "k80", adVarChar, 100
35            .Fields.Append "k85", adVarChar, 100
36            .Fields.Append "k90", adVarChar, 100
37            .Fields.Append "k95", adVarChar, 100
38            .Fields.Append "k100", adVarChar, 100
39            .Fields.Append "k105", adVarChar, 100
40            .Fields.Append "k110", adVarChar, 100
41            .Fields.Append "k115", adVarChar, 100
42            .Fields.Append "k120", adVarChar, 100
43            .Fields.Append "k125", adVarChar, 100
44            .Fields.Append "k130", adVarChar, 100
45            .Fields.Append "k135", adVarChar, 100
46            .Fields.Append "k140", adVarChar, 100
47            .Fields.Append "k145", adVarChar, 100
48            .Fields.Append "k150", adVarChar, 100
49            .Fields.Append "k155", adVarChar, 100
50            .Fields.Append "k160", adVarChar, 100
51            .Fields.Append "k165", adVarChar, 100
52            .Fields.Append "k170", adVarChar, 100
53            .Fields.Append "k175", adVarChar, 100
54            .Fields.Append "k180", adVarChar, 100
55            .Fields.Append "k185", adVarChar, 100
56            .Fields.Append "k190", adVarChar, 100
57            .Fields.Append "k195", adVarChar, 100
58            .Fields.Append "k200", adVarChar, 100
              
59            If Check1.value = vbChecked Then
60                .Fields.Append "dealerK1", adVarChar, 100
61                .Fields.Append "dealerK5", adVarChar, 100
62                .Fields.Append "dealerK10", adVarChar, 100
63                .Fields.Append "dealerK15", adVarChar, 100
64                .Fields.Append "dealerK20", adVarChar, 100
65                .Fields.Append "dealerK25", adVarChar, 100
66                .Fields.Append "dealerK30", adVarChar, 100
67                .Fields.Append "dealerK35", adVarChar, 100
68                .Fields.Append "dealerK40", adVarChar, 100
69                .Fields.Append "dealerK45", adVarChar, 100
70                .Fields.Append "dealerK50", adVarChar, 100
71                .Fields.Append "dealerK55", adVarChar, 100
72                .Fields.Append "dealerK60", adVarChar, 100
73                .Fields.Append "dealerK65", adVarChar, 100
74                .Fields.Append "dealerK70", adVarChar, 100
75                .Fields.Append "dealerK75", adVarChar, 100
76                .Fields.Append "dealerK80", adVarChar, 100
77                .Fields.Append "dealerK85", adVarChar, 100
78                .Fields.Append "dealerK90", adVarChar, 100
79                .Fields.Append "dealerK95", adVarChar, 100
80                .Fields.Append "dealerK100", adVarChar, 100
81                .Fields.Append "dealerK105", adVarChar, 100
82                .Fields.Append "dealerK110", adVarChar, 100
83                .Fields.Append "dealerK115", adVarChar, 100
84                .Fields.Append "dealerK120", adVarChar, 100
85                .Fields.Append "dealerK125", adVarChar, 100
86                .Fields.Append "dealerK130", adVarChar, 100
87                .Fields.Append "dealerK135", adVarChar, 100
88                .Fields.Append "dealerK140", adVarChar, 100
89                .Fields.Append "dealerK145", adVarChar, 100
90                .Fields.Append "dealerK150", adVarChar, 100
91                .Fields.Append "dealerK155", adVarChar, 100
92                .Fields.Append "dealerK160", adVarChar, 100
93                .Fields.Append "dealerK165", adVarChar, 100
94                .Fields.Append "dealerK170", adVarChar, 100
95                .Fields.Append "dealerK175", adVarChar, 100
96                .Fields.Append "dealerK180", adVarChar, 100
97                .Fields.Append "dealerK185", adVarChar, 100
98                .Fields.Append "dealerK190", adVarChar, 100
99                .Fields.Append "dealerK195", adVarChar, 100
100               .Fields.Append "dealerK200", adVarChar, 100
101           End If
102           .Fields.Append "lastSB", adVarChar, 100
103           .Fields.Append "dealerLastSB", adVarChar, 100
104           .Fields.Append "lastGR", adVarChar, 100
105           .Fields.Append "dealerLastGR", adVarChar, 100
106           .Fields.Append "lastInfoGR", adVarChar, 350
107           .Fields.Append "lastService", adVarChar, 100
108           .Fields.Append "DealerPenjual", adVarChar, 100
109           .Fields.Append "DealerLastSVC", adVarChar, 100
110           .Fields.Append "Selisih", adVarChar, 100
111           .Fields.Append "Warna", adVarChar, 50
112           .Fields.Append "aktif", adVarChar, 50
113           .Open
114       End With
115       Set gridTagih.DataSource = rsRecord
116       Call headColumns
          
End Sub

Private Sub headColumns()
          
1         With gridTagih
              Dim i As Integer
              Dim kolom As Integer
              
2              i = 0
3              Do While i <= 100
4                  .Splits(1).Columns(i).Visible = True
5                  i = i + 1
6              Loop
              
7             .Columns(0).Width = 2000
8             .Columns(0).Caption = "No. CHASSIS"
9             .Columns(0).DataField = "NO_CHASSIS"
10            .Splits(0).Columns(0).Locked = True

11            .Columns(6).Width = 1500
12            .Columns(6).Caption = "MODEL"
13            .Columns(6).DataField = "MDL"
14            .Splits(0).Columns(6).Locked = True

15            .Columns(2).Width = 2500
16            .Columns(2).Caption = "NAMA MODEL"
17            .Columns(2).DataField = "FMDL"
18            .Splits(0).Columns(2).Locked = True

19            .Columns(3).Width = 1500
20            .Columns(3).Caption = "NO POLISI"
21            .Columns(3).DataField = "NO_POLISI"
22            .Splits(0).Columns(3).Locked = True
              
23            .Columns(4).Width = 1000
24            .Columns(4).Caption = "TAHUN"
25            .Columns(4).DataField = "TAHUN"
26            .Splits(0).Columns(4).Locked = True
              
              
27            .Columns(5).Width = 2000
28            .Columns(5).Caption = "DELIVERY DATE"
29            .Columns(5).DataField = "ddate"
30            .Splits(0).Columns(5).Locked = True
              

31            .Columns(1).Width = 2000
32            .Columns(1).Caption = "NAMA CUSTOMER"
33            .Columns(1).DataField = "CUSTOMER"
34            .Splits(0).Columns(2).Locked = True
              
35            .Columns(7).Width = 4000
36            .Columns(7).Caption = "ALAMAT CUSTOMER"
37            .Columns(7).DataField = "CUST_ADRES"
38            .Splits(0).Columns(7).Locked = True
              
39            .Columns(8).Width = 2000
40            .Columns(8).Caption = "KOTA"
41            .Columns(8).DataField = "KOTA"
42            .Splits(0).Columns(8).Locked = True
              
43            .Columns(9).Width = 1500
44            .Columns(9).Caption = "TELP"
45            .Columns(9).DataField = "TELP"
46            .Splits(0).Columns(9).Locked = True
              
47            .Columns(12).Width = 2000
48            .Columns(12).Caption = "USENAME"
49            .Columns(12).DataField = "USENAME"
50            .Splits(0).Columns(12).Locked = True
              
51            .Columns(10).Width = 2000
52            .Columns(10).Caption = "GSM"
53            .Columns(10).DataField = "GSM"
54            .Splits(0).Columns(10).Locked = True
              
55            .Columns(11).Width = 2000
56            .Columns(11).Caption = "CDMA"
57            .Columns(11).DataField = "CDMA"
58            .Splits(0).Columns(11).Locked = True
              
59            .Columns(13).Width = 2000
60            .Columns(13).Caption = "USEGSM"
61            .Columns(13).DataField = "USEGSM"
62            .Splits(0).Columns(13).Locked = True
              
63            .Columns(14).Width = 2000
64            .Columns(14).Caption = "SBI 1K"
65            .Columns(14).DataField = "k1"
66            .Splits(1).Columns(14).Locked = False
              
67            If Check1.value = vbChecked Then
68                .Columns(15).Width = 2000
69                .Columns(15).Caption = "SBI 1K"
70                .Columns(15).DataField = "dealerK1"
71                .Splits(1).Columns(15).Locked = False
                  
72            End If
              'sbe 5k s/d 200k
              
              
              ''### 2016-05-26 || jika centang ho maka tambah kolom dealer
              
73            i = 5
              
74            If Check1.value = vbChecked Then
75                kolom = 16
76            Else
77                kolom = 15
78            End If
              
79            Do While i <= 200
                  
80                .Columns(kolom).Width = 2000
81                .Columns(kolom).Caption = "SBE " & i & "K"
82                .Columns(kolom).DataField = "k" & i
83                .Splits(1).Columns(kolom).Locked = False
84                If Check1.value = vbChecked Then
85                   kolom = kolom + 1
86                   .Columns(kolom).Width = 2000
87                   .Columns(kolom).Caption = "Dealer SBE " & i & "K"
88                   .Columns(kolom).DataField = "dealerK" & i
89                   .Splits(1).Columns(kolom).Locked = False
90                End If
91                kolom = kolom + 1
92                i = i + 5
93            Loop
              
94            i = 0
              
95            If Check1.value = vbChecked Then i = 43
              
96            .Columns(53 + i).Width = 2000
97            .Columns(53 + i).Caption = "Last SVC SB"
98            .Columns(53 + i).DataField = "lastSB"
99            .Splits(1).Columns(53 + i).Locked = False
              
100           .Columns(54 + i).Width = 2000
101           .Columns(54 + i).Caption = "Dealer Last SVC SB"
102           .Columns(54 + i).DataField = "dealerLastSB"
103           .Splits(1).Columns(54 + i).Locked = False
                      
104           .Columns(55 + i).Width = 2000
105           .Columns(55 + i).Caption = "Last SVC GR"
106           .Columns(55 + i).DataField = "lastGR"
107           .Splits(1).Columns(55 + i).Locked = False
              
108           .Columns(56 + i).Width = 2000
109           .Columns(56 + i).Caption = "Dealer Last SVC GR"
110           .Columns(56 + i).DataField = "dealerLastGR"
111           .Splits(1).Columns(56 + i).Locked = False
              
112           .Columns(57 + i).Width = 2000
113           .Columns(57 + i).Caption = "Info Last SVC GR"
114           .Columns(57 + i).DataField = "lastInfoGR"
115           .Splits(1).Columns(57 + i).Locked = False
              
116           .Columns(58 + i).Width = 2000
117           .Columns(58 + i).Caption = "Last SVC"
118           .Columns(58 + i).DataField = "lastService"
119           .Splits(1).Columns(58 + i).Locked = False
              
120           .Columns(59 + i).Width = 2000
121           .Columns(59 + i).Caption = "Dealer Penjual"
122           .Columns(59 + i).DataField = "DealerPenjual"
123           .Splits(1).Columns(59 + i).Locked = False
              
124           .Columns(60 + i).Width = 2000
125           .Columns(60 + i).Caption = "Dealer Terakhir Service"
126           .Columns(60 + i).DataField = "DealerLastSVC"
127           .Splits(1).Columns(60 + i).Locked = False
              
128           .Columns(61 + i).Width = 2000
129           .Columns(61 + i).Caption = "Kategori"
130           .Columns(61 + i).DataField = "aktif"
131           .Splits(1).Columns(61 + i).Locked = False
              
132           If Check1.value = vbUnchecked Then
133               i = 62
134               Do While i <= 103
135                   .Splits(1).Columns(i).Visible = False
136                   i = i + 1
137               Loop
138           End If
139           i = 2
140           .Splits(1).Columns(0).Visible = False
141           .Splits(1).Columns(1).Visible = False
142           Do While i <= 104
143               .Splits(0).Columns(i).Visible = False
144               i = i + 1
145           Loop
146       End With
End Sub

Private Sub refreshGrid()
      Dim rsTagih As ADODB.Recordset
      Dim rsDataHO As ADODB.Recordset

      Dim varTagih As String
      Dim kriteriaLastService As String
      Dim kriteria As String
      Dim sSQL As String
      Dim isSBE As Boolean
      Dim konter As Double
      Dim daftarNOChassis As String

1     isSBE = False
2     On Error GoTo vb_error
3         Screen.MousePointer = vbHourglass
4         kriteria = ""
          'pengecekan last service berdasarkan tanggal
5         If (IIf(IsNull(Tgl1), "", Tgl1) <> "") Then
6             kriteria = kriteria & " (J.[DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "') AND "
7         End If
          'pengecekan last sbe berdasarkan tanggal
8         If (IIf(IsNull(Tgl3), "", Tgl3) <> "") Then
9             isSBE = True
10            kriteria = kriteria & " ((J.[DATE] BETWEEN '" & Format(Tgl3, "MM/DD/YYYY") & "' AND '" & Format(Tgl4, "MM/DD/YYYY") & "') AND FS_RC <> '') and "
11        End If
          'pengecekan alamat customer
12        If Trim(TxtAlamat.Text) <> "" Then
13            kriteria = kriteria & " C.CUST_ADRES + C.KOTA LIKE '%" & Trim(TxtAlamat.Text) & "%' AND "
14        End If
          
15        If blnAkhir > 0 And blnAwal > 0 And (blnAkhir >= blnAwal) Then
16            kriteria = kriteria & " (DATEDIFF(MONTH,J.DATE,GETDATE())) BETWEEN " & blnAwal & " AND " & blnAkhir & " AND "
17        End If
18        If blnPasif > 0 Then
19            kriteria = kriteria & " (DATEDIFF(MONTH,J.DATE,GETDATE())) " & IIf(cbOperator.Text = "", "=", cbOperator.Text) & blnPasif & " AND "
20        End If
21        If Trim(kriteria) <> "" Then
22            kriteria = Left(kriteria, Len(kriteria) - 4)
23            kriteria = " WHERE " & kriteria
24        End If
          
          
25        varTagih = " and isnull(notagih,'') = ''"
26        If nTagih Then
27            varTagih = " or isnull(notagih,'') <> ''"
28        End If
29        jmlTagih = 0
30        If Check1.value = vbChecked Then
              'MENGAMBIL NO_CHASSIS LOCAL
31            sSQL = "SELECT DISTINCT C.NO_CHASSIS,C.MDL,C.FMDL,C.NO_POLISI,C.FS_TH_MOBIL,C.CUSTOMER,C.CUST_ADRES,"
32            sSQL = sSQL & " C.KOTA,C.TELPON,C.USENAME,C.USEGSM,C.FS_NM_DEALER FROM CUSTDAT C WITH (NOLOCK) INNER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_CHASSIS = C.NO_CHASSIS " & kriteria
              
33        Else
34            sSQL = "SELECT DISTINCT C.NO_CHASSIS,C.MDL,C.FMDL,C.NO_POLISI,C.FS_TH_MOBIL,C.CUSTOMER,C.CUST_ADRES,"
35            sSQL = sSQL & " C.KOTA,C.TELPON,C.USENAME,C.USEGSM,"
36            sSQL = sSQL & " J1.DATE AS '1K', J2.DATE AS '5K', J3.DATE AS '10K', J4.DATE AS '15K', J5.DATE AS '20K'"
37            sSQL = sSQL & " , J6.DATE AS '25K', J7.DATE AS '30K', J8.DATE AS '35K', J9.DATE AS '40K'"
38            sSQL = sSQL & " , J10.DATE AS '45K', J11.DATE AS '50K', J12.DATE AS '55K', J13.DATE AS '60K'"
39            sSQL = sSQL & " , J14.DATE AS '65K', J15.DATE AS '70K', J16.DATE AS '75K', J17.DATE AS '80K'"
40            sSQL = sSQL & " , J18.DATE AS '85K', J19.DATE AS '90K', J20.DATE AS '95K', J21.DATE AS '100K'"
41            sSQL = sSQL & " , J22.DATE AS '105K', J23.DATE AS '110K', J24.DATE AS '115K', J25.DATE AS '120K'"
42            sSQL = sSQL & " , J26.DATE AS '125K', J27.DATE AS '130K', J28.DATE AS '135K', J29.DATE AS '140K'"
43            sSQL = sSQL & " , J30.DATE AS '145K', J31.DATE AS '150K', J32.DATE AS '155K', J33.DATE AS '160K'"
44            sSQL = sSQL & " , J34.DATE AS '165K', J35.DATE AS '170K', J36.DATE AS '175K', J37.DATE AS '180K'"
45            sSQL = sSQL & " , J38.DATE AS '185K', J39.DATE AS '190K', J40.DATE AS '195K', J41.DATE AS '200K',"
46            sSQL = sSQL & " (SELECT TOP 1 REPLACE(j.REP1 + ',' + j.REP2 + ',' + j.REP3 + ',' + j.REP4 + ',' + j.REP5 + ',' + J.REP6 + ',' + J.REP7 + ',' + J.REP8,',,','') REP FROM JOBDTA J WHERE J.NO_CHASSIS = C.NO_CHASSIS and FS_RC = '' AND FS_LOG_DEL = '' ORDER BY DATE DESC) as jobGR,"
47            sSQL = sSQL & " (SELECT TOP 1 DATE FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = C.NO_CHASSIS AND FS_RC = '' AND FS_LOG_DEL = '' ORDER BY DATE DESC) AS LASTGR,"
48            sSQL = sSQL & " (SELECT TOP 1 DATE FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = C.NO_CHASSIS AND FS_RC <> '' AND FS_LOG_DEL = '' ORDER BY DATE DESC) AS LASTSB,"
49            sSQL = sSQL & " (SELECT TOP 1 DATE FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = C.NO_CHASSIS AND FS_LOG_DEL = '' ORDER BY DATE DESC) AS LASTSVC,"
50            sSQL = sSQL & " C.FS_NM_DEALER,(SELECT TOP 1 fs_nm_dealer FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = C.NO_CHASSIS AND FS_LOG_DEL = '' ORDER BY DATE DESC) AS DEALERLASTSVC"
51            sSQL = sSQL & " FROM CUSTDAT c WITH (NOLOCK) INNER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_CHASSIS = C.NO_CHASSIS "
52            sSQL = sSQL & " LEFT JOIN JOBDTA J1 WITH (NOLOCK) ON C.NO_CHASSIS = J1.NO_CHASSIS AND J1.FS_RC = 'R001' AND J1.FS_LOG_DEL = ''"
53            sSQL = sSQL & " LEFT JOIN JOBDTA J2 WITH (NOLOCK) ON C.NO_CHASSIS = J2.NO_CHASSIS AND J2.FS_RC = 'R005' AND J2.FS_LOG_DEL = ''"
54            sSQL = sSQL & " LEFT JOIN JOBDTA J3 WITH (NOLOCK) ON C.NO_CHASSIS = J3.NO_CHASSIS AND J3.FS_RC = 'R010' AND J3.FS_LOG_DEL = ''"
55            sSQL = sSQL & " LEFT JOIN JOBDTA J4 WITH (NOLOCK) ON C.NO_CHASSIS = J4.NO_CHASSIS AND J4.FS_RC = 'R015' AND J4.FS_LOG_DEL = ''"
56            sSQL = sSQL & " LEFT JOIN JOBDTA J5 WITH (NOLOCK) ON C.NO_CHASSIS = J5.NO_CHASSIS AND J5.FS_RC = 'R020' AND J5.FS_LOG_DEL = ''"
57            sSQL = sSQL & " LEFT JOIN JOBDTA J6 WITH (NOLOCK) ON C.NO_CHASSIS = J6.NO_CHASSIS AND J6.FS_RC = 'R025' AND J6.FS_LOG_DEL = ''"
58            sSQL = sSQL & " LEFT JOIN JOBDTA J7 WITH (NOLOCK) ON C.NO_CHASSIS = J7.NO_CHASSIS AND J7.FS_RC = 'R030' AND J7.FS_LOG_DEL = ''"
59            sSQL = sSQL & " LEFT JOIN JOBDTA J8 WITH (NOLOCK) ON C.NO_CHASSIS = J8.NO_CHASSIS AND J8.FS_RC = 'R035' AND J8.FS_LOG_DEL = ''"
60            sSQL = sSQL & " LEFT JOIN JOBDTA J9 WITH (NOLOCK) ON C.NO_CHASSIS = J9.NO_CHASSIS AND J9.FS_RC = 'R040' AND J9.FS_LOG_DEL = ''"
61            sSQL = sSQL & " LEFT JOIN JOBDTA J10 WITH (NOLOCK) ON C.NO_CHASSIS = J10.NO_CHASSIS AND J10.FS_RC = 'R045' AND J10.FS_LOG_DEL = ''"
62            sSQL = sSQL & " LEFT JOIN JOBDTA J11 WITH (NOLOCK) ON C.NO_CHASSIS = J11.NO_CHASSIS AND J11.FS_RC = 'R050' AND J11.FS_LOG_DEL = ''"
63            sSQL = sSQL & " LEFT JOIN JOBDTA J12 WITH (NOLOCK) ON C.NO_CHASSIS = J12.NO_CHASSIS AND J12.FS_RC = 'R055' AND J12.FS_LOG_DEL = ''"
64            sSQL = sSQL & " LEFT JOIN JOBDTA J13 WITH (NOLOCK) ON C.NO_CHASSIS = J13.NO_CHASSIS AND J13.FS_RC = 'R060' AND J13.FS_LOG_DEL = ''"
65            sSQL = sSQL & " LEFT JOIN JOBDTA J14 WITH (NOLOCK) ON C.NO_CHASSIS = J14.NO_CHASSIS AND J14.FS_RC = 'R065' AND J14.FS_LOG_DEL = ''"
66            sSQL = sSQL & " LEFT JOIN JOBDTA J15 WITH (NOLOCK) ON C.NO_CHASSIS = J15.NO_CHASSIS AND J15.FS_RC = 'R070' AND J15.FS_LOG_DEL = ''"
67            sSQL = sSQL & " LEFT JOIN JOBDTA J16 WITH (NOLOCK) ON C.NO_CHASSIS = J16.NO_CHASSIS AND J16.FS_RC = 'R075' AND J16.FS_LOG_DEL = ''"
68            sSQL = sSQL & " LEFT JOIN JOBDTA J17 WITH (NOLOCK) ON C.NO_CHASSIS = J17.NO_CHASSIS AND J17.FS_RC = 'R080' AND J17.FS_LOG_DEL = ''"
69            sSQL = sSQL & " LEFT JOIN JOBDTA J18 WITH (NOLOCK) ON C.NO_CHASSIS = J18.NO_CHASSIS AND J18.FS_RC = 'R085' AND J18.FS_LOG_DEL = ''"
70            sSQL = sSQL & " LEFT JOIN JOBDTA J19 WITH (NOLOCK) ON C.NO_CHASSIS = J19.NO_CHASSIS AND J19.FS_RC = 'R090' AND J19.FS_LOG_DEL = ''"
71            sSQL = sSQL & " LEFT JOIN JOBDTA J20 WITH (NOLOCK) ON C.NO_CHASSIS = J20.NO_CHASSIS AND J20.FS_RC = 'R095' AND J20.FS_LOG_DEL = ''"
72            sSQL = sSQL & " LEFT JOIN JOBDTA J21 WITH (NOLOCK) ON C.NO_CHASSIS = J21.NO_CHASSIS AND J21.FS_RC = 'R100' AND J21.FS_LOG_DEL = ''"
73            sSQL = sSQL & " LEFT JOIN JOBDTA J22 WITH (NOLOCK) ON C.NO_CHASSIS = J22.NO_CHASSIS AND J22.FS_RC = 'R105' AND J22.FS_LOG_DEL = ''"
74            sSQL = sSQL & " LEFT JOIN JOBDTA J23 WITH (NOLOCK) ON C.NO_CHASSIS = J23.NO_CHASSIS AND J23.FS_RC = 'R110' AND J23.FS_LOG_DEL = ''"
75            sSQL = sSQL & " LEFT JOIN JOBDTA J24 WITH (NOLOCK) ON C.NO_CHASSIS = J24.NO_CHASSIS AND J24.FS_RC = 'R115' AND J24.FS_LOG_DEL = ''"
76            sSQL = sSQL & " LEFT JOIN JOBDTA J25 WITH (NOLOCK) ON C.NO_CHASSIS = J25.NO_CHASSIS AND J25.FS_RC = 'R120' AND J25.FS_LOG_DEL = ''"
77            sSQL = sSQL & " LEFT JOIN JOBDTA J26 WITH (NOLOCK) ON C.NO_CHASSIS = J26.NO_CHASSIS AND J26.FS_RC = 'R125' AND J26.FS_LOG_DEL = ''"
78            sSQL = sSQL & " LEFT JOIN JOBDTA J27 WITH (NOLOCK) ON C.NO_CHASSIS = J27.NO_CHASSIS AND J27.FS_RC = 'R130' AND J27.FS_LOG_DEL = ''"
79            sSQL = sSQL & " LEFT JOIN JOBDTA J28 WITH (NOLOCK) ON C.NO_CHASSIS = J28.NO_CHASSIS AND J28.FS_RC = 'R135' AND J28.FS_LOG_DEL = ''"
80            sSQL = sSQL & " LEFT JOIN JOBDTA J29 WITH (NOLOCK) ON C.NO_CHASSIS = J29.NO_CHASSIS AND J29.FS_RC = 'R140' AND J29.FS_LOG_DEL = ''"
81            sSQL = sSQL & " LEFT JOIN JOBDTA J30 WITH (NOLOCK) ON C.NO_CHASSIS = J30.NO_CHASSIS AND J30.FS_RC = 'R145' AND J30.FS_LOG_DEL = ''"
82            sSQL = sSQL & " LEFT JOIN JOBDTA J31 WITH (NOLOCK) ON C.NO_CHASSIS = J31.NO_CHASSIS AND J31.FS_RC = 'R150' AND J31.FS_LOG_DEL = ''"
83            sSQL = sSQL & " LEFT JOIN JOBDTA J32 WITH (NOLOCK) ON C.NO_CHASSIS = J32.NO_CHASSIS AND J32.FS_RC = 'R155' AND J32.FS_LOG_DEL = ''"
84            sSQL = sSQL & " LEFT JOIN JOBDTA J33 WITH (NOLOCK) ON C.NO_CHASSIS = J33.NO_CHASSIS AND J33.FS_RC = 'R160' AND J33.FS_LOG_DEL = ''"
85            sSQL = sSQL & " LEFT JOIN JOBDTA J34 WITH (NOLOCK) ON C.NO_CHASSIS = J34.NO_CHASSIS AND J34.FS_RC = 'R165' AND J34.FS_LOG_DEL = ''"
86            sSQL = sSQL & " LEFT JOIN JOBDTA J35 WITH (NOLOCK) ON C.NO_CHASSIS = J35.NO_CHASSIS AND J35.FS_RC = 'R170' AND J35.FS_LOG_DEL = ''"
87            sSQL = sSQL & " LEFT JOIN JOBDTA J36 WITH (NOLOCK) ON C.NO_CHASSIS = J36.NO_CHASSIS AND J36.FS_RC = 'R175' AND J36.FS_LOG_DEL = ''"
88            sSQL = sSQL & " LEFT JOIN JOBDTA J37 WITH (NOLOCK) ON C.NO_CHASSIS = J37.NO_CHASSIS AND J37.FS_RC = 'R180' AND J37.FS_LOG_DEL = ''"
89            sSQL = sSQL & " LEFT JOIN JOBDTA J38 WITH (NOLOCK) ON C.NO_CHASSIS = J38.NO_CHASSIS AND J38.FS_RC = 'R185' AND J38.FS_LOG_DEL = ''"
90            sSQL = sSQL & " LEFT JOIN JOBDTA J39 WITH (NOLOCK) ON C.NO_CHASSIS = J39.NO_CHASSIS AND J39.FS_RC = 'R190' AND J39.FS_LOG_DEL = ''"
91            sSQL = sSQL & " LEFT JOIN JOBDTA J40 WITH (NOLOCK) ON C.NO_CHASSIS = J40.NO_CHASSIS AND J40.FS_RC = 'R195' AND J40.FS_LOG_DEL = ''"
92            sSQL = sSQL & " LEFT JOIN JOBDTA J41 WITH (NOLOCK) ON C.NO_CHASSIS = J41.NO_CHASSIS AND J41.FS_RC = 'R200' AND J41.FS_LOG_DEL = ''"
93            sSQL = sSQL & " " & kriteria
94        End If
          
          
          
95        lblPengambilanData.Visible = True
96        DoEvents
         ' If Check1.value = vbChecked Then
         '     Set rsTagih = objO.FetchDataHOUnlimitedTime(sSQL)
         ' Else
97            Set rsTagih = objO.FetchData(sSQL)
         ' End If
98        daftarNOChassis = ""
99        If Check1.value = vbChecked Then
100           If Not rsTagih.EOF Then
101               rsTagih.MoveFirst
102               daftarNOChassis = "("
103               Do While Not rsTagih.EOF
104                   daftarNOChassis = daftarNOChassis & "'" & Trim(rsTagih("")) & "',"
105                   rsTagih.MoveNext
106               Loop
107               daftarNOChassis = Left(daftarNOChassis, Len(daftarNOChassis) - 1)
108               daftarNOChassis = daftarNOChassis & ")"
109           End If
110           If daftarNOChassis <> "" Then
111               sSQL = "SELECT DISTINCT j.no_chassis, CONVERT(VARCHAR(20),isnull(J1.DATE,'')) + J1.FS_NM_DEALER AS '1K', CONVERT(VARCHAR(20),isnull(J2.DATE,'')) + J2.FS_NM_DEALER AS '5K', CONVERT(VARCHAR(20),isnull(J3.DATE,'')) + J3.FS_NM_DEALER AS '10K', CONVERT(VARCHAR(20),isnull(J4.DATE,'')) + J4.FS_NM_DEALER AS '15K',"
112               sSQL = sSQL & " CONVERT(VARCHAR(20),isnull(J5.DATE,'')) + J5.FS_NM_DEALER AS '20K', CONVERT(VARCHAR(20),isnull(J6.DATE,'')) + J6.FS_NM_DEALER AS '25K', CONVERT(VARCHAR(20),isnull(J7.DATE,'')) + J7.FS_NM_DEALER AS '30K', CONVERT(VARCHAR(20),isnull(J8.DATE,'')) + J8.FS_NM_DEALER AS '35K', CONVERT(VARCHAR(20),isnull(J9.DATE,'')) + J9.FS_NM_DEALER AS '40K'"
113               sSQL = sSQL & " , CONVERT(VARCHAR(20),isnull(J10.DATE,'')) + J10.FS_NM_DEALER AS '45K', CONVERT(VARCHAR(20),isnull(J11.DATE,'')) + J11.FS_NM_DEALER AS '50K', CONVERT(VARCHAR(20),isnull(J12.DATE,'')) + J12.FS_NM_DEALER AS '55K', CONVERT(VARCHAR(20),isnull(J13.DATE,'')) + J13.FS_NM_DEALER AS '60K'"
114               sSQL = sSQL & " , CONVERT(VARCHAR(20),isnull(J14.DATE,'')) + J14.FS_NM_DEALER AS '65K', CONVERT(VARCHAR(20),isnull(J15.DATE,'')) + J15.FS_NM_DEALER AS '70K', CONVERT(VARCHAR(20),isnull(J16.DATE,'')) + J16.FS_NM_DEALER AS '75K', CONVERT(VARCHAR(20),isnull(J17.DATE,'')) + J17.FS_NM_DEALER AS '80K'"
115               sSQL = sSQL & " , CONVERT(VARCHAR(20),isnull(J18.DATE,'')) + J18.FS_NM_DEALER AS '85K', CONVERT(VARCHAR(20),isnull(J19.DATE,'')) + J19.FS_NM_DEALER AS '90K', CONVERT(VARCHAR(20),isnull(J20.DATE,'')) + J20.FS_NM_DEALER AS '95K', CONVERT(VARCHAR(20),isnull(J21.DATE,'')) + J21.FS_NM_DEALER AS '100K'"
116               sSQL = sSQL & " , CONVERT(VARCHAR(20),isnull(J22.DATE,'')) + J22.FS_NM_DEALER AS '105K', CONVERT(VARCHAR(20),isnull(J23.DATE,'')) + J23.FS_NM_DEALER AS '110K', CONVERT(VARCHAR(20),isnull(J24.DATE,'')) + J24.FS_NM_DEALER AS '115K', CONVERT(VARCHAR(20),isnull(J25.DATE,'')) + J25.FS_NM_DEALER AS '120K'"
117               sSQL = sSQL & " , CONVERT(VARCHAR(20),isnull(J26.DATE,'')) + J26.FS_NM_DEALER AS '125K', CONVERT(VARCHAR(20),isnull(J27.DATE,'')) + J27.FS_NM_DEALER AS '130K', CONVERT(VARCHAR(20),isnull(J28.DATE,'')) + J28.FS_NM_DEALER AS '135K', CONVERT(VARCHAR(20),isnull(J29.DATE,'')) + J29.FS_NM_DEALER AS '140K'"
118               sSQL = sSQL & " , CONVERT(VARCHAR(20),isnull(J30.DATE,'')) + J30.FS_NM_DEALER AS '145K', CONVERT(VARCHAR(20),isnull(J31.DATE,'')) + J31.FS_NM_DEALER AS '150K', CONVERT(VARCHAR(20),isnull(J32.DATE,'')) + J32.FS_NM_DEALER AS '155K', CONVERT(VARCHAR(20),isnull(J33.DATE,'')) + J33.FS_NM_DEALER AS '160K'"
119               sSQL = sSQL & " , CONVERT(VARCHAR(20),isnull(J34.DATE,'')) + J34.FS_NM_DEALER AS '165K', CONVERT(VARCHAR(20),isnull(J35.DATE,'')) + J35.FS_NM_DEALER AS '170K', CONVERT(VARCHAR(20),isnull(J36.DATE,'')) + J36.FS_NM_DEALER AS '175K', CONVERT(VARCHAR(20),isnull(J37.DATE,'')) + J37.FS_NM_DEALER AS '180K'"
120               sSQL = sSQL & " , CONVERT(VARCHAR(20),isnull(J38.DATE,'')) + J38.FS_NM_DEALER AS '185K', CONVERT(VARCHAR(20),isnull(J39.DATE,'')) + J39.FS_NM_DEALER AS '190K', CONVERT(VARCHAR(20),isnull(J40.DATE,'')) + J40.FS_NM_DEALER AS '195K', CONVERT(VARCHAR(20),isnull(J41.DATE,'')) + J41.FS_NM_DEALER AS '200K',"
121               sSQL = sSQL & " (SELECT TOP 1 REPLACE(j.REP1 + ',' + j.REP2 + ',' + j.REP3 + ',' + j.REP4 + ',' + j.REP5 + ',' + J.REP6 + ',' + J.REP7 + ',' + J.REP8,',,','') REP FROM JOBDTA J WHERE J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' and FS_RC = '' AND FS_LOG_DEL = '' ORDER BY DATE DESC) as jobGR,"
122               sSQL = sSQL & " (SELECT TOP 1 DATE FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' AND FS_RC = '' AND FS_LOG_DEL = '' ORDER BY DATE DESC) AS LASTGR,"
123               sSQL = sSQL & " (SELECT TOP 1 DATE FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' AND FS_RC <> '' AND FS_LOG_DEL = '' ORDER BY DATE DESC) AS LASTSB,"
124               sSQL = sSQL & " (SELECT TOP 1 DATE FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' AND FS_LOG_DEL = '' ORDER BY DATE DESC) AS LASTSVC,"
125               sSQL = sSQL & " (SELECT TOP 1 fs_nm_dealer FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' AND FS_LOG_DEL = '' ORDER BY DATE DESC) AS DEALERLASTSVC"
126               sSQL = sSQL & " FROM JOBDTA J WITH (NOLOCK)"
127               sSQL = sSQL & " LEFT JOIN JOBDTA J1 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J1.NO_CHASSIS AND J1.FS_RC = 'R001' AND J1.FS_LOG_DEL = ''"
128               sSQL = sSQL & " LEFT JOIN JOBDTA J2 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J2.NO_CHASSIS AND J2.FS_RC = 'R005' AND J2.FS_LOG_DEL = ''"
129               sSQL = sSQL & " LEFT JOIN JOBDTA J3 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J3.NO_CHASSIS AND J3.FS_RC = 'R010' AND J3.FS_LOG_DEL = ''"
130               sSQL = sSQL & " LEFT JOIN JOBDTA J4 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J4.NO_CHASSIS AND J4.FS_RC = 'R015' AND J4.FS_LOG_DEL = ''"
131               sSQL = sSQL & " LEFT JOIN JOBDTA J5 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J5.NO_CHASSIS AND J5.FS_RC = 'R020' AND J5.FS_LOG_DEL = ''"
132               sSQL = sSQL & " LEFT JOIN JOBDTA J6 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J6.NO_CHASSIS AND J6.FS_RC = 'R025' AND J6.FS_LOG_DEL = ''"
133               sSQL = sSQL & " LEFT JOIN JOBDTA J7 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J7.NO_CHASSIS AND J7.FS_RC = 'R030' AND J7.FS_LOG_DEL = ''"
134               sSQL = sSQL & " LEFT JOIN JOBDTA J8 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J8.NO_CHASSIS AND J8.FS_RC = 'R035' AND J8.FS_LOG_DEL = ''"
135               sSQL = sSQL & " LEFT JOIN JOBDTA J9 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J9.NO_CHASSIS AND J9.FS_RC = 'R040' AND J9.FS_LOG_DEL = ''"
136               sSQL = sSQL & " LEFT JOIN JOBDTA J10 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J10.NO_CHASSIS AND J10.FS_RC = 'R045' AND J10.FS_LOG_DEL = ''"
137               sSQL = sSQL & " LEFT JOIN JOBDTA J11 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J11.NO_CHASSIS AND J11.FS_RC = 'R050' AND J11.FS_LOG_DEL = ''"
138               sSQL = sSQL & " LEFT JOIN JOBDTA J12 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J12.NO_CHASSIS AND J12.FS_RC = 'R055' AND J12.FS_LOG_DEL = ''"
139               sSQL = sSQL & " LEFT JOIN JOBDTA J13 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J13.NO_CHASSIS AND J13.FS_RC = 'R060' AND J13.FS_LOG_DEL = ''"
140               sSQL = sSQL & " LEFT JOIN JOBDTA J14 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J14.NO_CHASSIS AND J14.FS_RC = 'R065' AND J14.FS_LOG_DEL = ''"
141               sSQL = sSQL & " LEFT JOIN JOBDTA J15 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J15.NO_CHASSIS AND J15.FS_RC = 'R070' AND J15.FS_LOG_DEL = ''"
142               sSQL = sSQL & " LEFT JOIN JOBDTA J16 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J16.NO_CHASSIS AND J16.FS_RC = 'R075' AND J16.FS_LOG_DEL = ''"
143               sSQL = sSQL & " LEFT JOIN JOBDTA J17 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J17.NO_CHASSIS AND J17.FS_RC = 'R080' AND J17.FS_LOG_DEL = ''"
144               sSQL = sSQL & " LEFT JOIN JOBDTA J18 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J18.NO_CHASSIS AND J18.FS_RC = 'R085' AND J18.FS_LOG_DEL = ''"
145               sSQL = sSQL & " LEFT JOIN JOBDTA J19 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J19.NO_CHASSIS AND J19.FS_RC = 'R090' AND J19.FS_LOG_DEL = ''"
146               sSQL = sSQL & " LEFT JOIN JOBDTA J20 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J20.NO_CHASSIS AND J20.FS_RC = 'R095' AND J20.FS_LOG_DEL = ''"
147               sSQL = sSQL & " LEFT JOIN JOBDTA J21 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J21.NO_CHASSIS AND J21.FS_RC = 'R100' AND J21.FS_LOG_DEL = ''"
148               sSQL = sSQL & " LEFT JOIN JOBDTA J22 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J22.NO_CHASSIS AND J22.FS_RC = 'R105' AND J22.FS_LOG_DEL = ''"
149               sSQL = sSQL & " LEFT JOIN JOBDTA J23 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J23.NO_CHASSIS AND J23.FS_RC = 'R110' AND J23.FS_LOG_DEL = ''"
150               sSQL = sSQL & " LEFT JOIN JOBDTA J24 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J24.NO_CHASSIS AND J24.FS_RC = 'R115' AND J24.FS_LOG_DEL = ''"
151               sSQL = sSQL & " LEFT JOIN JOBDTA J25 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J25.NO_CHASSIS AND J25.FS_RC = 'R120' AND J25.FS_LOG_DEL = ''"
152               sSQL = sSQL & " LEFT JOIN JOBDTA J26 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J26.NO_CHASSIS AND J26.FS_RC = 'R125' AND J26.FS_LOG_DEL = ''"
153               sSQL = sSQL & " LEFT JOIN JOBDTA J27 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J27.NO_CHASSIS AND J27.FS_RC = 'R130' AND J27.FS_LOG_DEL = ''"
154               sSQL = sSQL & " LEFT JOIN JOBDTA J28 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J28.NO_CHASSIS AND J28.FS_RC = 'R135' AND J28.FS_LOG_DEL = ''"
155               sSQL = sSQL & " LEFT JOIN JOBDTA J29 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J29.NO_CHASSIS AND J29.FS_RC = 'R140' AND J29.FS_LOG_DEL = ''"
156               sSQL = sSQL & " LEFT JOIN JOBDTA J30 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J30.NO_CHASSIS AND J30.FS_RC = 'R145' AND J30.FS_LOG_DEL = ''"
157               sSQL = sSQL & " LEFT JOIN JOBDTA J31 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J31.NO_CHASSIS AND J31.FS_RC = 'R150' AND J31.FS_LOG_DEL = ''"
158               sSQL = sSQL & " LEFT JOIN JOBDTA J32 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J32.NO_CHASSIS AND J32.FS_RC = 'R155' AND J32.FS_LOG_DEL = ''"
159               sSQL = sSQL & " LEFT JOIN JOBDTA J33 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J33.NO_CHASSIS AND J33.FS_RC = 'R160' AND J33.FS_LOG_DEL = ''"
160               sSQL = sSQL & " LEFT JOIN JOBDTA J34 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J34.NO_CHASSIS AND J34.FS_RC = 'R165' AND J34.FS_LOG_DEL = ''"
161               sSQL = sSQL & " LEFT JOIN JOBDTA J35 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J35.NO_CHASSIS AND J35.FS_RC = 'R170' AND J35.FS_LOG_DEL = ''"
162               sSQL = sSQL & " LEFT JOIN JOBDTA J36 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J36.NO_CHASSIS AND J36.FS_RC = 'R175' AND J36.FS_LOG_DEL = ''"
163               sSQL = sSQL & " LEFT JOIN JOBDTA J37 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J37.NO_CHASSIS AND J37.FS_RC = 'R180' AND J37.FS_LOG_DEL = ''"
164               sSQL = sSQL & " LEFT JOIN JOBDTA J38 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J38.NO_CHASSIS AND J38.FS_RC = 'R185' AND J38.FS_LOG_DEL = ''"
165               sSQL = sSQL & " LEFT JOIN JOBDTA J39 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J39.NO_CHASSIS AND J39.FS_RC = 'R190' AND J39.FS_LOG_DEL = ''"
166               sSQL = sSQL & " LEFT JOIN JOBDTA J40 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J40.NO_CHASSIS AND J40.FS_RC = 'R195' AND J40.FS_LOG_DEL = ''"
167               sSQL = sSQL & " LEFT JOIN JOBDTA J41 WITH (NOLOCK) ON '" & Trim(rsTagih("NO_CHASSIS")) & "' = J41.NO_CHASSIS AND J41.FS_RC = 'R200' AND J41.FS_LOG_DEL = ''"
168               sSQL = sSQL & " where j.NO_CHASSIS in " & daftarNOChassis
                  
169               Set rsDataHO = objO.FetchDataHOUnlimitedTime(sSQL)
170           End If
171       End If
172       If rsTagih.RecordCount > 0 Then
173           Call createRS
174           lblPengambilanData.Visible = False
175           ProgressBar1.Visible = True
176           rsTagih.MoveFirst
177           ProgressBar1.Max = rsTagih.RecordCount + 1
178           ProgressBar1.Min = 0
179           Do While Not rsTagih.EOF
180               rsDataHO.filter = ""
181               With rsRecord
182                   .AddNew
                   
183                   !NO_CHASSIS = Trim(rsTagih("NO_CHASSIS"))
184                   !MDL = Trim(rsTagih("MDL"))
185                   !Fmdl = Trim(rsTagih("FMDL"))
186                   !NO_POLISI = Trim(rsTagih("NO_POLISI"))
187                   !Tahun = Trim(rsTagih("FS_TH_MOBIL"))
188                   !customer = Trim(rsTagih("CUSTOMER"))
189                   !cust_adres = Trim(rsTagih("CUST_ADRES"))
190                   !Kota = Trim(rsTagih("KOTA"))
191                   !telp = Trim(rsTagih("TELPON"))
192                   !UseName = Trim(rsTagih("USENAME"))
193                   !UseGSM = Trim(rsTagih("USEGSM"))
194                   !DealerPenjual = IIf(IsNull(Trim((rsTagih("fs_nm_dealer")))), "", Trim(rsTagih("fs_nm_dealer"))) 'Trim(rsTagih("fs_nm_dealer"))
195                   If Check1.value = vbChecked Then
                          
196                       rsDataHO.filter = "no_chassis = '" & Trim(rsTagih("no_CHassis")) & "'"
197                       If Not rsDataHO.EOF Then
198                           rsDataHO.MoveFirst
199                           !K1 = IIf(IsNull(Trim(rsDataHO("1K"))), "", Trim(rsDataHO("1K")))
                              'sbe
200                           !k5 = IIf(IsNull(Trim(rsDataHO("5K"))), "", Trim(rsDataHO("5K")))
201                           !K10 = IIf(IsNull(Trim((rsDataHO("10K")))), "", Trim(rsDataHO("10K")))
202                           !k15 = IIf(IsNull(Trim((rsDataHO("15K")))), "", Trim(rsDataHO("15K")))
203                           !K20 = IIf(IsNull(Trim((rsDataHO("20K")))), "", Trim(rsDataHO("20K")))
204                           !k25 = IIf(IsNull(Trim((rsDataHO("25K")))), "", Trim(rsDataHO("25K")))
205                           !K30 = IIf(IsNull(Trim((rsDataHO("30K")))), "", Trim(rsDataHO("30K")))
206                           !k35 = IIf(IsNull(Trim((rsDataHO("35K")))), "", Trim(rsDataHO("35K")))
207                           !K40 = IIf(IsNull(Trim((rsDataHO("40K")))), "", Trim(rsDataHO("40K")))
208                           !k45 = IIf(IsNull(Trim((rsDataHO("45K")))), "", Trim(rsDataHO("45K")))
209                           !K50 = IIf(IsNull(Trim((rsDataHO("50K")))), "", Trim(rsDataHO("50K")))
210                           !k55 = IIf(IsNull(Trim((rsDataHO("55K")))), "", Trim(rsDataHO("55K")))
211                           !K60 = IIf(IsNull(Trim((rsDataHO("60K")))), "", Trim(rsDataHO("60K")))
212                           !k65 = IIf(IsNull(Trim((rsDataHO("65K")))), "", Trim(rsDataHO("65K")))
213                           !K70 = IIf(IsNull(Trim((rsDataHO("70K")))), "", Trim(rsDataHO("70K")))
214                           !k75 = IIf(IsNull(Trim((rsDataHO("75K")))), "", Trim(rsDataHO("75K")))
215                           !K80 = IIf(IsNull(Trim((rsDataHO("80K")))), "", Trim(rsDataHO("80K")))
216                           !k85 = IIf(IsNull(Trim((rsDataHO("85K")))), "", Trim(rsDataHO("85K")))
217                           !k90 = IIf(IsNull(Trim((rsDataHO("90K")))), "", Trim(rsDataHO("90K")))
218                           !k95 = IIf(IsNull(Trim((rsDataHO("95K")))), "", Trim(rsDataHO("95K")))
219                           !k100 = IIf(IsNull(Trim((rsDataHO("100K")))), "", Trim(rsDataHO("100K")))
220                           !k105 = IIf(IsNull(Trim((rsDataHO("105K")))), "", Trim(rsDataHO("105K")))
221                           !k110 = IIf(IsNull(Trim((rsDataHO("110K")))), "", Trim(rsDataHO("110K")))
222                           !k115 = IIf(IsNull(Trim((rsDataHO("115K")))), "", Trim(rsDataHO("115K")))
223                           !k120 = IIf(IsNull(Trim((rsDataHO("120K")))), "", Trim(rsDataHO("120K")))
224                           !k125 = IIf(IsNull(Trim((rsDataHO("125K")))), "", Trim(rsDataHO("125K")))
225                           !k130 = IIf(IsNull(Trim((rsDataHO("130K")))), "", Trim(rsDataHO("130K")))
226                           !k135 = IIf(IsNull(Trim((rsDataHO("135K")))), "", Trim(rsDataHO("135K")))
227                           !k140 = IIf(IsNull(Trim((rsDataHO("140K")))), "", Trim(rsDataHO("140K")))
228                           !k145 = IIf(IsNull(Trim((rsDataHO("145K")))), "", Trim(rsDataHO("145K")))
229                           !k150 = IIf(IsNull(Trim((rsDataHO("150K")))), "", Trim(rsDataHO("150K")))
230                           !k155 = IIf(IsNull(Trim((rsDataHO("155K")))), "", Trim(rsDataHO("155K")))
231                           !k160 = IIf(IsNull(Trim((rsDataHO("160K")))), "", Trim(rsDataHO("160K")))
232                           !k165 = IIf(IsNull(Trim((rsDataHO("165K")))), "", Trim(rsDataHO("165K")))
233                           !k170 = IIf(IsNull(Trim((rsDataHO("170K")))), "", Trim(rsDataHO("170K")))
234                           !k175 = IIf(IsNull(Trim((rsDataHO("175K")))), "", Trim(rsDataHO("175K")))
235                           !k180 = IIf(IsNull(Trim((rsDataHO("180K")))), "", Trim(rsDataHO("180K")))
236                           !k185 = IIf(IsNull(Trim((rsDataHO("185K")))), "", Trim(rsDataHO("185K")))
237                           !k190 = IIf(IsNull(Trim((rsDataHO("190K")))), "", Trim(rsDataHO("190K")))
238                           !k195 = IIf(IsNull(Trim((rsDataHO("195K")))), "", Trim(rsDataHO("195K")))
239                           !k200 = IIf(IsNull(Trim((rsDataHO("200K")))), "", Trim(rsDataHO("200K")))
240                           !lastSB = IIf(IsNull(Trim((rsDataHO("lastSB")))), "", Trim(rsDataHO("lastSB")))
241                           !lastGR = IIf(IsNull(Trim((rsDataHO("lastGR")))), "", Trim(rsDataHO("lastGR")))
242                           !lastService = IIf(IsNull(Trim((rsDataHO("lastSVC")))), "", Trim(rsDataHO("lastSVC")))
243                           !lastInfoGR = IIf(IsNull(Trim((rsDataHO("jobGR")))), "", Trim(rsDataHO("jobGR")))
244                           !DealerLastSVC = IIf(IsNull(Trim((rsDataHO("dealerlastsvc")))), "", Trim(rsDataHO("dealerlastsvc")))
245                       Else
246                           !K1 = ""
                              'sbe
247                           !k5 = ""
248                           !K10 = ""
249                           !k15 = ""
250                           !K20 = ""
251                           !k25 = ""
252                           !K30 = ""
253                           !k35 = ""
254                           !K40 = ""
255                           !k45 = ""
256                           !K50 = ""
257                           !k55 = ""
258                           !K60 = ""
259                           !k65 = ""
260                           !K70 = ""
261                           !k75 = ""
262                           !K80 = ""
263                           !k85 = ""
264                           !k90 = ""
265                           !k95 = ""
266                           !k100 = ""
267                           !k105 = ""
268                           !k110 = ""
269                           !k115 = ""
270                           !k120 = ""
271                           !k125 = ""
272                           !k130 = ""
273                           !k135 = ""
274                           !k140 = ""
275                           !k145 = ""
276                           !k150 = ""
277                           !k155 = ""
278                           !k160 = ""
279                           !k165 = ""
280                           !k170 = ""
281                           !k175 = ""
282                           !k180 = ""
283                           !k185 = ""
284                           !k190 = ""
285                           !k195 = ""
286                           !k200 = ""
287                           !lastSB = ""
288                           !lastGR = ""
289                           !lastService = ""
290                           !lastInfoGR = ""
291                           !DealerLastSVC = ""
292                       End If
293                   Else
                          'sbi
294                       !K1 = ""  'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R001' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          'sbe
295                       !k5 = IIf(IsNull(Trim(rsTagih("5K"))), "", Trim(rsTagih("5K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R005' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
296                       !K10 = IIf(IsNull(Trim((rsTagih("10K")))), "", Trim(rsTagih("10K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R010' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
297                       !k15 = IIf(IsNull(Trim((rsTagih("15K")))), "", Trim(rsTagih("15K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R015' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
298                       !K20 = IIf(IsNull(Trim((rsTagih("20K")))), "", Trim(rsTagih("20K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R020' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
299                       !k25 = IIf(IsNull(Trim((rsTagih("25K")))), "", Trim(rsTagih("25K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R025' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
300                       !K30 = IIf(IsNull(Trim((rsTagih("30K")))), "", Trim(rsTagih("30K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R030' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
301                       !k35 = IIf(IsNull(Trim((rsTagih("35K")))), "", Trim(rsTagih("35K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R035' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
302                       !K40 = IIf(IsNull(Trim((rsTagih("40K")))), "", Trim(rsTagih("40K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R040' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
303                       !k45 = IIf(IsNull(Trim((rsTagih("45K")))), "", Trim(rsTagih("45K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R045' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
304                       !K50 = IIf(IsNull(Trim((rsTagih("50K")))), "", Trim(rsTagih("50K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R050' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
305                       !k55 = IIf(IsNull(Trim((rsTagih("55K")))), "", Trim(rsTagih("55K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R055' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
306                       !K60 = IIf(IsNull(Trim((rsTagih("60K")))), "", Trim(rsTagih("60K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R060' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
307                       !k65 = IIf(IsNull(Trim((rsTagih("65K")))), "", Trim(rsTagih("65K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R065' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
308                       !K70 = IIf(IsNull(Trim((rsTagih("70K")))), "", Trim(rsTagih("70K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R070' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
309                       !k75 = IIf(IsNull(Trim((rsTagih("75K")))), "", Trim(rsTagih("75K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R075' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
310                       !K80 = IIf(IsNull(Trim((rsTagih("80K")))), "", Trim(rsTagih("80K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R080' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
311                       !k85 = IIf(IsNull(Trim((rsTagih("85K")))), "", Trim(rsTagih("85K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R085' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
312                       !k90 = IIf(IsNull(Trim((rsTagih("90K")))), "", Trim(rsTagih("90K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R090' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
313                       !k95 = IIf(IsNull(Trim((rsTagih("95K")))), "", Trim(rsTagih("95K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R095' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
314                       !k100 = IIf(IsNull(Trim((rsTagih("100K")))), "", Trim(rsTagih("100K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R100' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
315                       !k105 = IIf(IsNull(Trim((rsTagih("105K")))), "", Trim(rsTagih("105K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R105' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
316                       !k110 = IIf(IsNull(Trim((rsTagih("110K")))), "", Trim(rsTagih("110K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R110' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
317                       !k115 = IIf(IsNull(Trim((rsTagih("115K")))), "", Trim(rsTagih("115K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R115' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
318                       !k120 = IIf(IsNull(Trim((rsTagih("120K")))), "", Trim(rsTagih("120K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R120' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
319                       !k125 = IIf(IsNull(Trim((rsTagih("125K")))), "", Trim(rsTagih("125K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R125' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
320                       !k130 = IIf(IsNull(Trim((rsTagih("130K")))), "", Trim(rsTagih("130K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R130' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
321                       !k135 = IIf(IsNull(Trim((rsTagih("135K")))), "", Trim(rsTagih("135K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R135' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
322                       !k140 = IIf(IsNull(Trim((rsTagih("140K")))), "", Trim(rsTagih("140K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R140' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
323                       !k145 = IIf(IsNull(Trim((rsTagih("145K")))), "", Trim(rsTagih("145K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R145' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
324                       !k150 = IIf(IsNull(Trim((rsTagih("150K")))), "", Trim(rsTagih("150K"))) 'ObjO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R150' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
325                       !k155 = IIf(IsNull(Trim((rsTagih("155K")))), "", Trim(rsTagih("155K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R155' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
326                       !k160 = IIf(IsNull(Trim((rsTagih("160K")))), "", Trim(rsTagih("160K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R160' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
327                       !k165 = IIf(IsNull(Trim((rsTagih("165K")))), "", Trim(rsTagih("165K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R165' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
328                       !k170 = IIf(IsNull(Trim((rsTagih("170K")))), "", Trim(rsTagih("170K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R170' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
329                       !k175 = IIf(IsNull(Trim((rsTagih("175K")))), "", Trim(rsTagih("175K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R175' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
330                       !k180 = IIf(IsNull(Trim((rsTagih("180K")))), "", Trim(rsTagih("180K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R180' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
331                       !k185 = IIf(IsNull(Trim((rsTagih("185K")))), "", Trim(rsTagih("185K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R185' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
332                       !k190 = IIf(IsNull(Trim((rsTagih("190K")))), "", Trim(rsTagih("190K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R190' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
333                       !k195 = IIf(IsNull(Trim((rsTagih("195K")))), "", Trim(rsTagih("195K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R195' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
334                       !k200 = IIf(IsNull(Trim((rsTagih("200K")))), "", Trim(rsTagih("200K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R200' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
335                       !lastSB = IIf(IsNull(Trim((rsTagih("lastSB")))), "", Trim(rsTagih("lastSB"))) 'Trim(rsTagih("lastSB"))
336                       !lastGR = IIf(IsNull(Trim((rsTagih("lastGR")))), "", Trim(rsTagih("lastGR"))) 'Trim(rsTagih("lastGR"))
337                       !lastService = IIf(IsNull(Trim((rsTagih("lastSVC")))), "", Trim(rsTagih("lastSVC"))) 'Trim(rsTagih("lastSVC"))
338                       !lastInfoGR = IIf(IsNull(Trim((rsTagih("jobGR")))), "", Trim(rsTagih("jobGR"))) 'Trim(rsTagih("jobGR"))
339                       !DealerLastSVC = IIf(IsNull(Trim((rsTagih("dealerlastsvc")))), "", Trim(rsTagih("dealerlastsvc")))
340                   End If
341                   .Update
342               End With
      '            If Trim(rsTagih("notagih")) <> "" Then
      '                jmlTagih = jmlTagih + 1
      '            End If
343               ProgressBar1.value = konter
344               DoEvents
345               LblJumlah.Caption = konter
346               konter = konter + 1
347               rsTagih.MoveNext
348           Loop
              
349       Else
              
350           Call createRS
351       End If
352       lblPengambilanData.Visible = False
353       ProgressBar1.Visible = False
354       LblJumlah.Caption = rsRecord.RecordCount
355       Screen.MousePointer = vbDefault
356       Exit Sub
vb_error:
357       MsgBox "Terjadi kesalahan pengambilan data." & vbCrLf & "Error Number = " & err.Number & vbCrLf & err.Description & vbCrLf & "Procedure : " & Me.Name & ".RefreshGrid"
358       lblPengambilanData.Visible = False
359       ProgressBar1.Visible = False
360       LblJumlah.Caption = rsRecord.RecordCount
361       Screen.MousePointer = vbDefault
          
End Sub

Public Sub refreshGridBaru()
Dim rsTagih As ADODB.Recordset
Dim rsDataHO As ADODB.Recordset

Dim varTagih As String
Dim kriteriaLastService As String
Dim kriteria As String
Dim sSQL As String
Dim isSBE As Boolean
Dim konter As Double
Dim daftarNOChassis As String

jumAktif = 0
jumPasif1 = 0
jumPasif2 = 0
jumPasif3 = 0

isSBE = False
On Error GoTo vb_error
    Screen.MousePointer = vbHourglass
    kriteria = ""
    'pengecekan last service berdasarkan tanggal
    If (IIf(IsNull(Tgl1), "", Tgl1) <> "") Then
        kriteria = kriteria & " (lastservice BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "') AND "
    End If
    'pengecekan last sbe berdasarkan tanggal
    If (IIf(IsNull(Tgl3), "", Tgl3) <> "") Then
        isSBE = True
        kriteria = kriteria & " ((lastsb BETWEEN '" & Format(Tgl3, "MM/DD/YYYY") & "' AND '" & Format(Tgl4, "MM/DD/YYYY") & "')) and "
    End If
    'pengecekan alamat customer
    If Trim(TxtAlamat.Text) <> "" Then
        kriteria = kriteria & " CUST_ADRES + KOTA LIKE '%" & Trim(TxtAlamat.Text) & "%' AND "
    End If
    
    If Trim(kriteria) <> "" Then
        kriteria = Left(kriteria, Len(kriteria) - 4)
        kriteria = " WHERE " & kriteria
    End If
    
    
    varTagih = " and isnull(notagih,'') = ''"
    If nTagih Then
        varTagih = " or isnull(notagih,'') <> ''"
    End If
    jmlTagih = 0
'    If Check1.value = vbChecked Then
'        'MENGAMBIL NO_CHASSIS LOCAL
'        sSQL = "SELECT DISTINCT C.NO_CHASSIS,C.MDL,C.FMDL,C.NO_POLISI,C.FS_TH_MOBIL,C.CUSTOMER,C.CUST_ADRES,"
'        sSQL = sSQL & " C.KOTA,C.TELPON,C.USENAME,C.USEGSM,C.FS_NM_DEALER FROM CUSTDAT C WITH (NOLOCK) INNER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_CHASSIS = C.NO_CHASSIS " & kriteria
'
'    Else
        sSQL = "SELECT *,DATEDIFF(MONTH,LASTSERVICE,GETDATE()) SELISIH FROM MAPPINGCUSTOMER " & kriteria
'    End If
    
    lblPengambilanData.Visible = True
    DoEvents
   ' If Check1.value = vbChecked Then
   '     Set rsTagih = objO.FetchDataHOUnlimitedTime(sSQL)
   ' Else
        Set rsTagih = objO.FetchData(sSQL)
   ' End If
    daftarNOChassis = ""
    If Check1.value = vbChecked Then
        If Not rsTagih.EOF Then
            rsTagih.MoveFirst
            daftarNOChassis = "("
            Do While Not rsTagih.EOF
                daftarNOChassis = daftarNOChassis & "'" & Trim(rsTagih("no_chassis")) & "',"
                rsTagih.MoveNext
            Loop
            daftarNOChassis = Left(daftarNOChassis, Len(daftarNOChassis) - 1)
            daftarNOChassis = daftarNOChassis & ")"
        End If
        If daftarNOChassis <> "" Then
            
            sSQL = "SELECT * FROM MAPPINGCUSTOMER with (nolock) where NO_CHASSIS in " & daftarNOChassis
            
            Set rsDataHO = objO.FetchDataHOUnlimitedTime(sSQL)
        End If
    End If
    If rsTagih.RecordCount > 0 Then
        Call createRS
        lblPengambilanData.Visible = False
        ProgressBar1.Visible = True
        rsTagih.MoveFirst
        ProgressBar1.Max = rsTagih.RecordCount + 1
        ProgressBar1.Min = 0
        Do While Not rsTagih.EOF
            If Check1.value = vbChecked Then rsDataHO.filter = ""
            With rsRecord
                .AddNew
             
                !NO_CHASSIS = Trim(rsTagih("NO_CHASSIS"))
                !MDL = Trim(rsTagih("MDL"))
                !Fmdl = Trim(rsTagih("FMDL"))
                !NO_POLISI = Trim(rsTagih("NO_POLISI"))
                !Tahun = Trim(rsTagih("TAHUN"))
                !customer = Trim(rsTagih("CUSTOMER"))
                !cust_adres = Trim(rsTagih("CUST_ADRES"))
                !Kota = Trim(rsTagih("KOTA"))
                !telp = Trim(rsTagih("TELP"))
                !UseName = Trim(rsTagih("USENAME"))
                !UseGSM = Trim(rsTagih("USEGSM"))
                !DealerPenjual = IIf(IsNull(Trim((rsTagih("DEALERPENJUAL")))), "", Trim(rsTagih("DEALERPENJUAL"))) 'Trim(rsTagih("fs_nm_dealer"))
                If Check1.value = vbChecked Then
                    
                    rsDataHO.filter = "no_chassis = '" & Trim(rsTagih("no_CHassis")) & "'"
                    If Not rsDataHO.EOF Then
                        rsDataHO.MoveFirst
                        !K1 = IIf(Trim(rsDataHO("K1")) = "01/01/1900", "", Format(Trim(rsDataHO("K1")), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K1Dealer")))
                        'sbe
                        !k5 = IIf(Trim(rsDataHO("K5")) = "01/01/1900", "", Format(Trim(rsDataHO("K5")), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K5Dealer")))
                        !K10 = IIf(Trim(rsDataHO("K10")) = "01/01/1900", "", Format(Trim((rsDataHO("K10"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K10Dealer")))
                        !k15 = IIf(Trim(rsDataHO("K15")) = "01/01/1900", "", Format(Trim((rsDataHO("K15"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K15Dealer")))
                        !K20 = IIf(Trim(rsDataHO("K20")) = "01/01/1900", "", Format(Trim((rsDataHO("K20"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K20Dealer")))
                        !k25 = IIf(Trim(rsDataHO("K25")) = "01/01/1900", "", Format(Trim((rsDataHO("K25"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K25Dealer")))
                        !K30 = IIf(Trim(rsDataHO("K30")) = "01/01/1900", "", Format(Trim((rsDataHO("K30"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K30Dealer")))
                        !k35 = IIf(Trim(rsDataHO("K35")) = "01/01/1900", "", Format(Trim((rsDataHO("K35"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K35Dealer")))
                        !K40 = IIf(Trim(rsDataHO("K40")) = "01/01/1900", "", Format(Trim((rsDataHO("K40"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K40Dealer")))
                        !k45 = IIf(Trim(rsDataHO("K45")) = "01/01/1900", "", Format(Trim((rsDataHO("K45"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K45Dealer")))
                        !K50 = IIf(Trim(rsDataHO("K50")) = "01/01/1900", "", Format(Trim((rsDataHO("K50"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K50Dealer")))
                        !k55 = IIf(Trim(rsDataHO("K55")) = "01/01/1900", "", Format(Trim((rsDataHO("K55"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K55Dealer")))
                        !K60 = IIf(Trim(rsDataHO("K60")) = "01/01/1900", "", Format(Trim((rsDataHO("K60"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K60Dealer")))
                        !k65 = IIf(Trim(rsDataHO("K65")) = "01/01/1900", "", Format(Trim((rsDataHO("K65"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K65Dealer")))
                        !K70 = IIf(Trim(rsDataHO("K70")) = "01/01/1900", "", Format(Trim((rsDataHO("K70"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K70Dealer")))
                        !k75 = IIf(Trim(rsDataHO("K75")) = "01/01/1900", "", Format(Trim((rsDataHO("K75"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K75Dealer")))
                        !K80 = IIf(Trim(rsDataHO("K80")) = "01/01/1900", "", Format(Trim((rsDataHO("K80"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K80Dealer")))
                        !k85 = IIf(Trim(rsDataHO("K85")) = "01/01/1900", "", Format(Trim((rsDataHO("K85"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K85Dealer")))
                        !k90 = IIf(Trim(rsDataHO("K90")) = "01/01/1900", "", Format(Trim((rsDataHO("K90"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K90Dealer")))
                        !k95 = IIf(Trim(rsDataHO("K95")) = "01/01/1900", "", Format(Trim((rsDataHO("K95"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K95Dealer")))
                        !k100 = IIf(Trim(rsDataHO("K100")) = "01/01/1900", "", Format(Trim((rsDataHO("K100"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K100Dealer")))
                        !k105 = IIf(Trim(rsDataHO("K105")) = "01/01/1900", "", Format(Trim((rsDataHO("K105"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K105Dealer")))
                        !k110 = IIf(Trim(rsDataHO("K110")) = "01/01/1900", "", Format(Trim((rsDataHO("K110"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K110Dealer")))
                        !k115 = IIf(Trim(rsDataHO("K115")) = "01/01/1900", "", Format(Trim((rsDataHO("K115"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K115Dealer")))
                        !k120 = IIf(Trim(rsDataHO("K120")) = "01/01/1900", "", Format(Trim((rsDataHO("K120"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K120Dealer")))
                        !k125 = IIf(Trim(rsDataHO("K125")) = "01/01/1900", "", Format(Trim((rsDataHO("K125"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K125Dealer")))
                        !k130 = IIf(Trim(rsDataHO("K130")) = "01/01/1900", "", Format(Trim((rsDataHO("K130"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K130Dealer")))
                        !k135 = IIf(Trim(rsDataHO("K135")) = "01/01/1900", "", Format(Trim((rsDataHO("K135"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K135Dealer")))
                        !k140 = IIf(Trim(rsDataHO("K140")) = "01/01/1900", "", Format(Trim((rsDataHO("K140"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K140Dealer")))
                        !k145 = IIf(Trim(rsDataHO("K145")) = "01/01/1900", "", Format(Trim((rsDataHO("K145"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K145Dealer")))
                        !k150 = IIf(Trim(rsDataHO("K150")) = "01/01/1900", "", Format(Trim((rsDataHO("K150"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K150Dealer")))
                        !k155 = IIf(Trim(rsDataHO("K155")) = "01/01/1900", "", Format(Trim((rsDataHO("K155"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K155Dealer")))
                        !k160 = IIf(Trim(rsDataHO("K160")) = "01/01/1900", "", Format(Trim((rsDataHO("K160"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K160Dealer")))
                        !k165 = IIf(Trim(rsDataHO("K165")) = "01/01/1900", "", Format(Trim((rsDataHO("K165"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K165Dealer")))
                        !k170 = IIf(Trim(rsDataHO("K170")) = "01/01/1900", "", Format(Trim((rsDataHO("K170"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K170Dealer")))
                        !k175 = IIf(Trim(rsDataHO("K175")) = "01/01/1900", "", Format(Trim((rsDataHO("K175"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K175Dealer")))
                        !k180 = IIf(Trim(rsDataHO("K180")) = "01/01/1900", "", Format(Trim((rsDataHO("K180"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K180Dealer")))
                        !k185 = IIf(Trim(rsDataHO("K185")) = "01/01/1900", "", Format(Trim((rsDataHO("K185"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K185Dealer")))
                        !k190 = IIf(Trim(rsDataHO("K190")) = "01/01/1900", "", Format(Trim((rsDataHO("K190"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K190Dealer")))
                        !k195 = IIf(Trim(rsDataHO("K195")) = "01/01/1900", "", Format(Trim((rsDataHO("K195"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K195Dealer")))
                        !k200 = IIf(Trim(rsDataHO("K200")) = "01/01/1900", "", Format(Trim((rsDataHO("K200"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("K200Dealer")))
                        !lastSB = Format(Trim((rsDataHO("lastSB"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("lastSBDealer"))
                        !lastGR = Format(Trim((rsDataHO("lastGR"))), "DD-MM-YYYY") & " - " & Trim(rsDataHO("lastGRDealer"))
                        !lastService = IIf(IsNull(Trim((rsDataHO("lastservice")))), "", Trim(rsDataHO("lastservice")))
                        !lastInfoGR = IIf(IsNull(Trim((rsDataHO("lastInfoGR")))), "", Trim(rsDataHO("lastInfoGR")))
                        !DealerLastSVC = IIf(IsNull(Trim((rsDataHO("dealerlastsvc")))), "", Trim(rsDataHO("dealerlastsvc")))
                    Else
                        !K1 = ""
                        'sbe
                        !k5 = ""
                        !K10 = ""
                        !k15 = ""
                        !K20 = ""
                        !k25 = ""
                        !K30 = ""
                        !k35 = ""
                        !K40 = ""
                        !k45 = ""
                        !K50 = ""
                        !k55 = ""
                        !K60 = ""
                        !k65 = ""
                        !K70 = ""
                        !k75 = ""
                        !K80 = ""
                        !k85 = ""
                        !k90 = ""
                        !k95 = ""
                        !k100 = ""
                        !k105 = ""
                        !k110 = ""
                        !k115 = ""
                        !k120 = ""
                        !k125 = ""
                        !k130 = ""
                        !k135 = ""
                        !k140 = ""
                        !k145 = ""
                        !k150 = ""
                        !k155 = ""
                        !k160 = ""
                        !k165 = ""
                        !k170 = ""
                        !k175 = ""
                        !k180 = ""
                        !k185 = ""
                        !k190 = ""
                        !k195 = ""
                        !k200 = ""
                        !lastSB = ""
                        !lastGR = ""
                        !lastService = ""
                        !lastInfoGR = ""
                        !DealerLastSVC = ""
                    End If
                Else
                    'sbi
                    !K1 = Format(Trim(rsTagih("K1")), "DD-MM-YYYY") 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R001' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    'sbe
                    !k5 = Format(Trim(rsTagih("K5")), "DD-MM-YYYY") 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R005' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !K10 = Format(Trim((rsTagih("K10"))), "DD-MM-YYYY") '10K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R010' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k15 = Format(Trim((rsTagih("K15"))), "DD-MM-YYYY") '15K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R015' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !K20 = Format(Trim((rsTagih("K20"))), "DD-MM-YYYY") '20K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R020' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k25 = Format(Trim((rsTagih("K25"))), "DD-MM-YYYY") '25K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R025' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !K30 = Format(Trim((rsTagih("K30"))), "DD-MM-YYYY") '30K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R030' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k35 = Format(Trim((rsTagih("K35"))), "DD-MM-YYYY") '35K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R035' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !K40 = Format(Trim((rsTagih("K40"))), "DD-MM-YYYY") '40K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R040' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k45 = Format(Trim((rsTagih("K45"))), "DD-MM-YYYY") '45K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R045' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !K50 = Format(Trim((rsTagih("K50"))), "DD-MM-YYYY") '50K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R050' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k55 = Format(Trim((rsTagih("K55"))), "DD-MM-YYYY") '55K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R055' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !K60 = Format(Trim((rsTagih("K60"))), "DD-MM-YYYY") '60K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R060' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k65 = Format(Trim((rsTagih("K65"))), "DD-MM-YYYY") '65K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R065' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !K70 = Format(Trim((rsTagih("K70"))), "DD-MM-YYYY") '70K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R070' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k75 = Format(Trim((rsTagih("K75"))), "DD-MM-YYYY") '75K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R075' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !K80 = Format(Trim((rsTagih("K80"))), "DD-MM-YYYY") '80K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R080' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k85 = Format(Trim((rsTagih("K85"))), "DD-MM-YYYY") '85K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R085' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k90 = Format(Trim((rsTagih("K90"))), "DD-MM-YYYY") '90K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R090' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k95 = Format(Trim((rsTagih("K95"))), "DD-MM-YYYY") '95K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R095' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k100 = Format(Trim((rsTagih("K100"))), "DD-MM-YYYY") '100K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R100' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k105 = Format(Trim((rsTagih("K105"))), "DD-MM-YYYY") '105K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R105' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k110 = Format(Trim((rsTagih("K110"))), "DD-MM-YYYY") '110K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R110' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k115 = Format(Trim((rsTagih("K115"))), "DD-MM-YYYY") '115K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R115' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k120 = Format(Trim((rsTagih("K120"))), "DD-MM-YYYY") '120K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R120' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k125 = Format(Trim((rsTagih("K125"))), "DD-MM-YYYY") '125K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R125' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k130 = Format(Trim((rsTagih("K130"))), "DD-MM-YYYY") '130K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R130' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k135 = Format(Trim((rsTagih("K135"))), "DD-MM-YYYY") '135K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R135' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k140 = Format(Trim((rsTagih("K140"))), "DD-MM-YYYY") '140K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R140' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k145 = Format(Trim((rsTagih("K145"))), "DD-MM-YYYY") '145K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R145' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k150 = Format(Trim((rsTagih("K150"))), "DD-MM-YYYY") '150K"))) 'ObjO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R150' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k155 = Format(Trim((rsTagih("K155"))), "DD-MM-YYYY") '155K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R155' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k160 = Format(Trim((rsTagih("K160"))), "DD-MM-YYYY") '160K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R160' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k165 = Format(Trim((rsTagih("K165"))), "DD-MM-YYYY") '165K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R165' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k170 = Format(Trim((rsTagih("K170"))), "DD-MM-YYYY") '170K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R170' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k175 = Format(Trim((rsTagih("K175"))), "DD-MM-YYYY") '175K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R175' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k180 = Format(Trim((rsTagih("K180"))), "DD-MM-YYYY") '180K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R180' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k185 = Format(Trim((rsTagih("K185"))), "DD-MM-YYYY") '185K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R185' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k190 = Format(Trim((rsTagih("K190"))), "DD-MM-YYYY") '190K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R190' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k195 = Format(Trim((rsTagih("K195"))), "DD-MM-YYYY") '195K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R195' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !k200 = Format(Trim((rsTagih("K200"))), "DD-MM-YYYY") '200K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R200' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                    !lastSB = Format(Trim((rsTagih("lastSB"))), "DD-MM-YYYY") 'lastSB"))) 'Trim(rsTagih("lastSB"))
                    !lastGR = Format(Trim((rsTagih("lastGR"))), "DD-MM-YYYY") 'lastGR"))) 'Trim(rsTagih("lastGR"))
                    !lastService = Format(Trim((rsTagih("LASTSERVICE"))), "DD-MM-YYYY") 'lastSVC"))) 'Trim(rsTagih("lastSVC"))
                    !lastInfoGR = IIf(IsNull(Trim((rsTagih("lastInfoGR")))), "", Trim(rsTagih("lastInfoGR"))) 'Trim(rsTagih("jobGR"))
                    !DealerLastSVC = IIf(IsNull(Trim((rsTagih("dealerlastsvc")))), "", Trim(rsTagih("dealerlastsvc")))
                    !selisih = CInt(rsTagih("selisih"))
                    !warna = getWarnaUnit(CInt(rsTagih("selisih")))
                    
                End If
                .Update
            End With
'            If Trim(rsTagih("notagih")) <> "" Then
'                jmlTagih = jmlTagih + 1
'            End If
            ProgressBar1.value = konter
            DoEvents
            LblJumlah.Caption = konter
            konter = konter + 1
            rsTagih.MoveNext
        Loop
        
    Else
        
        Call createRS
    End If
    gridTagih.FetchRowStyle = True
    lblPengambilanData.Visible = False
    ProgressBar1.Visible = False
    LblJumlah.Caption = rsRecord.RecordCount
    Screen.MousePointer = vbDefault
    Exit Sub
vb_error:
    MsgBox "Terjadi kesalahan pengambilan data." & vbCrLf & "Error Number = " & err.Number & vbCrLf & err.Description & vbCrLf & "Procedure : " & Me.Name & ".RefreshGrid"
    lblPengambilanData.Visible = False
    ProgressBar1.Visible = False
    LblJumlah.Caption = 0 ' rsRecord.RecordCount
    Screen.MousePointer = vbDefault

End Sub


Private Sub gridSublet_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsRecord.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    If Nul(rs.Fields(14).value) = True Then
    RowStyle = "SudahDitagih"
    End If
End Sub

Private Sub gridTagih_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsRecord.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    RowStyle = "Putih"
    If rs("warna") <> "" Then
        If UCase(rs("Warna")) = "MERAH" Then
            RowStyle = "MERAH"
        ElseIf UCase(rs("WARNA")) = "KUNING" Then
            RowStyle = "KUNING"
        ElseIf UCase(rs("WARNA")) = "BIRU" Then
            RowStyle = "BIRU"
        ElseIf UCase(rs("WARNA")) = "HIJAU" Then
            RowStyle = "HIJAU"
        Else
            RowStyle = "PUTIH"
        End If
    End If
    
End Sub

Private Sub gridTagih_FilterChange()
    On Error GoTo errHandler
    Set cols = gridTagih.Columns
    Dim C As Integer
    
    C = gridTagih.col
    gridTagih.HoldFields
    rsRecord.filter = getFilter()
    If rsRecord.EOF Then
        For Each col In gridTagih.Columns
            col.FilterText = ""
        Next col
        rsRecord.filter = ""
    End If
    gridTagih.col = C
    gridTagih.EditActive = True
    LblJumlah.Caption = rsRecord.RecordCount
    
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

Private Sub gridTagih_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 12 Then rsRecord.Sort = gridTagih.Columns(ColIndex).DataField & " " & Urut
4        gridTagih.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
End Sub

Private Sub gridTagih_AfterColUpdate(ByVal ColIndex As Integer)
    With gridTagih
        Select Case ColIndex
            Case 6
                If .Columns(ColIndex) Then
                    jmlTagih = jmlTagih + 1
                Else
                    jmlTagih = jmlTagih - 1
                End If
                If chAll.value = True Then
                    MsgBox "Cekbox tampilkan semua tidak boleh dicetang !", vbInformation + vbOKOnly, "Sublet"
                End If
        End Select
        
        nTagih.Caption = jmlTagih
    End With
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Public Function getWarnaUnit(total As Integer) As String
    Dim hasil As String
    hasil = "Putih"
    If ckAktif.value = vbChecked Then
        If cbOperator.Text = "<" Then
           
                If CInt(Trim(total)) < blnAktif Then
                    hasil = LabelAktif.Caption
                    jumAktif = jumAktif + 1
                End If
            
        ElseIf cbOperator.Text = ">" Then
            If (total) > blnAktif Then
                hasil = LabelAktif.Caption
                jumAktif = jumAktif + 1
            End If
        ElseIf cbOperator.Text = "=" Then
            If (total) = blnAktif Then
                hasil = LabelAktif.Caption
                jumAktif = jumAktif + 1
            End If
        ElseIf cbOperator.Text = "RANGE" Then
            If (total) >= blnRangeAktifAwal And (total) <= blnRangeAktifAkhir Then
                hasil = LabelAktif.Caption
                jumAktif = jumAktif + 1
            End If
        End If
    End If
    '------------ PASIF 1 ---------------
    If ckPasif1.value = vbChecked Then
        If Combo1.Text = "<" Then
         
                If CInt((total)) < blnPasif1 Then
                    hasil = LabelPasif.Caption
                    jumPasif1 = jumPasif1 + 1
                End If
         
        ElseIf Combo1.Text = ">" Then
            If (total) > blnPasif1 Then
                hasil = LabelPasif.Caption
                jumPasif1 = jumPasif1 + 1
            End If
        ElseIf Combo1.Text = "=" Then
            If (total) = blnPasif1 Then
                hasil = LabelPasif.Caption
                jumPasif1 = jumPasif1 + 1
            End If
        ElseIf Combo1.Text = "RANGE" Then
            If total >= blnRangePasif1Awal And total <= blnRangePasif1Akhir Then
                hasil = LabelPasif.Caption
                jumPasif1 = jumPasif1 + 1
            End If
        End If
    End If
    '------------ PASIF 2 ---------------
    If ckPasif2.value = vbChecked Then
        If Combo2.Text = "<" Then
            
                If CInt((total)) < blnPasif2 Then
                    hasil = LabelPasif2.Caption
                    jumPasif2 = jumPasif2 + 1
                End If
          
        ElseIf Combo2.Text = ">" Then
            If (total) > blnPasif2 Then
                hasil = LabelPasif2.Caption
                jumPasif2 = jumPasif2 + 1
            End If
        ElseIf Combo2.Text = "=" Then
            If (total) = blnPasif2 Then
                hasil = LabelPasif2.Caption
                jumPasif2 = jumPasif2 + 1
            End If
        ElseIf Combo2.Text = "RANGE" Then
            If (total) >= blnPasif2Awal And (total) <= blnPasif2Akhir Then
                hasil = LabelPasif2.Caption
                jumPasif2 = jumPasif2 + 1
            End If
        End If
    End If
    '------------ PASIF 3 ---------------
    If ckPasif3.value = vbChecked Then
        If Combo3.Text = "<" Then
            
                If CInt((total)) < blnPasif3Bulan Then
                    hasil = LabelPasif3.Caption
                    jumPasif3 = jumPasif3 + 1
                End If
           
        ElseIf Combo3.Text = ">" Then
            If (total) > blnPasif3Bulan Then
                hasil = LabelPasif3.Caption
                 jumPasif3 = jumPasif3 + 1
            End If
        ElseIf Combo3.Text = "=" Then
            If (total) = blnPasif3Bulan Then
                hasil = LabelPasif3.Caption
                 jumPasif3 = jumPasif3 + 1
            End If
        ElseIf Combo3.Text = "RANGE" Then
            If (total) >= blnPasif3Awal And (total) <= blnPasif3Akhir Then
                hasil = LabelPasif3.Caption
                 jumPasif3 = jumPasif3 + 1
            End If
        End If
    End If
    getWarnaUnit = hasil
End Function



Private Sub Option2_Click()

End Sub

