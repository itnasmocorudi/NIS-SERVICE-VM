VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmRegistrasi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Registrasi & Close Order"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   915
      Visible         =   0   'False
      Width           =   1395
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
      Height          =   300
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   915
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   12720
      TabIndex        =   64
      Top             =   480
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
   Begin VB.Frame frmLaporan 
      BackColor       =   &H00000040&
      Height          =   3135
      Left            =   2520
      TabIndex        =   76
      Top             =   2880
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CheckBox chInvB 
         BackColor       =   &H00000040&
         Caption         =   "Tampilkan Invoice -B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   85
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton cmdCloseLap 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   82
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdCetak 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   2520
         Width           =   1770
      End
      Begin TDBDate6Ctl.TDBDate DtAwal 
         Height          =   315
         Left            =   840
         TabIndex        =   77
         Top             =   600
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "frmRegistrasi.frx":0000
         Caption         =   "frmRegistrasi.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRegistrasi.frx":0198
         Keys            =   "frmRegistrasi.frx":01B6
         Spin            =   "frmRegistrasi.frx":0214
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
      Begin TDBDate6Ctl.TDBDate DtAkhir 
         Height          =   315
         Left            =   840
         TabIndex        =   79
         Top             =   1440
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "frmRegistrasi.frx":023C
         Caption         =   "frmRegistrasi.frx":0368
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRegistrasi.frx":03D4
         Keys            =   "frmRegistrasi.frx":03F2
         Spin            =   "frmRegistrasi.frx":0450
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   1320
         TabIndex        =   80
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Transaksi"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   720
         TabIndex        =   78
         Top             =   120
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdLap 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Laporan Kasir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1530
   End
   Begin VB.CommandButton TblSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9180
      Width           =   1650
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9180
      Width           =   1770
   End
   Begin VB.Frame alasanPass 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2280
      TabIndex        =   67
      Top             =   3000
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmdClose 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         TabIndex        =   72
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton tblCetakPassSementara 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5445
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2400
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ditulis lengkap penanggung, waktu membayar, pembayaran via"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   70
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alasan Pemberian Gatepass Sementara"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   4170
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   SERVICE    "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   10335
      Begin TrueOleDBGrid70.TDBGrid JobGrid 
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3201
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
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "JOB ORDER"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(58)  =   "Named:id=33:Normal"
         _StyleDefs(59)  =   ":id=33,.parent=0"
         _StyleDefs(60)  =   "Named:id=34:Heading"
         _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   ":id=34,.wraptext=-1"
         _StyleDefs(63)  =   "Named:id=35:Footing"
         _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=36:Selected"
         _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(67)  =   "Named:id=37:Caption"
         _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(69)  =   "Named:id=38:HighlightRow"
         _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=39:EvenRow"
         _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(73)  =   "Named:id=40:OddRow"
         _StyleDefs(74)  =   ":id=40,.parent=33"
         _StyleDefs(75)  =   "Named:id=41:RecordSelector"
         _StyleDefs(76)  =   ":id=41,.parent=34"
         _StyleDefs(77)  =   "Named:id=42:FilterBar"
         _StyleDefs(78)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid partGrid 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
         _LayoutType     =   0
         _RowHeight      =   -2147483647
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
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "PART ORDER"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid subletGrid 
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   4680
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "SUBLET ORDER"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Named:id=33:Normal"
         _StyleDefs(47)  =   ":id=33,.parent=0"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   ":id=34,.wraptext=-1"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=37:Caption"
         _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(57)  =   "Named:id=38:HighlightRow"
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   28
         Left            =   720
         TabIndex        =   8
         Top             =   6495
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   27
         Left            =   5640
         TabIndex        =   47
         Top             =   525
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   22
         Left            =   2160
         TabIndex        =   50
         Top             =   525
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   21
         Left            =   2160
         TabIndex        =   51
         Top             =   240
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   23
         Left            =   3960
         TabIndex        =   52
         Top             =   240
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   24
         Left            =   3960
         TabIndex        =   56
         Top             =   525
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   25
         Left            =   5640
         TabIndex        =   58
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   26
         Left            =   6960
         TabIndex        =   59
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   20
         Left            =   120
         TabIndex        =   61
         Top             =   525
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin VB.Label lblPDS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   11160
         TabIndex        =   86
         Top             =   240
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         TabIndex        =   62
         Top             =   285
         Width           =   570
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe WO"
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
         Left            =   4800
         TabIndex        =   60
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T SA."
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
         Left            =   3480
         TabIndex        =   57
         Top             =   570
         Width           =   390
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. No."
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
         Left            =   1485
         TabIndex        =   55
         Top             =   570
         Width           =   600
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM"
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
         Left            =   1845
         TabIndex        =   54
         Top             =   285
         Width           =   210
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mech"
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
         Left            =   3480
         TabIndex        =   53
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suplay Slip"
         Height          =   195
         Left            =   4695
         TabIndex        =   48
         Top             =   570
         Width           =   780
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   6540
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdGateS 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Gatepass Sementara"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1200
      Width           =   1530
   End
   Begin VB.CommandButton cmdGate 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Gatepass"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   1800
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   $"frmRegistrasi.frx":0478
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   8685
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   930
         TabIndex        =   11
         Top             =   630
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   930
         TabIndex        =   12
         Top             =   915
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   930
         TabIndex        =   13
         Top             =   1230
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   930
         TabIndex        =   14
         Top             =   1530
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   5190
         TabIndex        =   15
         Top             =   330
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   5190
         TabIndex        =   16
         Top             =   630
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   5190
         TabIndex        =   17
         Top             =   1230
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   5190
         TabIndex        =   18
         Top             =   930
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   13
         Left            =   7770
         TabIndex        =   19
         Top             =   630
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   20
         Top             =   330
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   930
         TabIndex        =   21
         Top             =   330
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   5190
         TabIndex        =   22
         Top             =   1530
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   5190
         TabIndex        =   23
         Top             =   1815
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   7350
         TabIndex        =   24
         Top             =   315
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   930
         TabIndex        =   25
         Top             =   1815
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   930
         TabIndex        =   26
         Top             =   2115
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   29
         Left            =   5190
         TabIndex        =   27
         Top             =   2115
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   19
         Left            =   3645
         TabIndex        =   49
         Top             =   330
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   18
         Left            =   5190
         TabIndex        =   83
         Top             =   2520
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine."
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
         Left            =   4440
         TabIndex        =   84
         Top             =   2565
         Width           =   540
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   6720
         X2              =   8400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Customer && Kendaraan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2400
         TabIndex        =   74
         Top             =   0
         Width           =   3675
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOU"
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
         Left            =   4635
         TabIndex        =   41
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA"
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
         Left            =   345
         TabIndex        =   40
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM"
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
         Left            =   465
         TabIndex        =   39
         Top             =   1860
         Width           =   315
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO TAM"
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
         Left            =   6600
         TabIndex        =   38
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame."
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
         Left            =   4530
         TabIndex        =   37
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
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
         Left            =   4395
         TabIndex        =   36
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Lgn"
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
         TabIndex        =   35
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
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
         Left            =   4560
         TabIndex        =   34
         Top             =   1245
         Width           =   480
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   4620
         TabIndex        =   33
         Top             =   645
         Width           =   420
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Th. Mobil"
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
         Left            =   4395
         TabIndex        =   32
         Top             =   375
         Width           =   645
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN"
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
         TabIndex        =   31
         Top             =   1575
         Width           =   375
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
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
         Left            =   450
         TabIndex        =   30
         Top             =   1275
         Width           =   330
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Left            =   285
         TabIndex        =   29
         Top             =   975
         Width           =   495
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   375
         TabIndex        =   28
         Top             =   675
         Width           =   405
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   8535
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2600
      TabIndex        =   42
      Top             =   47
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   40
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   7080
      TabIndex        =   43
      Top             =   83
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
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
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9855
      Left            =   0
      Picture         =   "frmRegistrasi.frx":0518
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLOSING"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   9120
      TabIndex        =   73
      Top             =   120
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. WO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   46
      Top             =   55
      Width           =   945
   End
   Begin VB.Label LblReturn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Job"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3120
      TabIndex        =   45
      Top             =   97
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   6375
      TabIndex        =   44
      Top             =   135
      Width           =   570
   End
End
Attribute VB_Name = "frmRegistrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsPart As New ADODB.Recordset
Dim rsSublet As New ADODB.Recordset
Dim Report As Object
Dim objBilling As New BClsBilling
Dim simpan As Boolean
Dim nilaiRupiah As Double

Private Sub hapusForm()
1        On Error GoTo hapusForm_Error

2         For i = 0 To 29
3             Select Case i
                  Case 10, 21
4                     Text(i) = 0
5                 Case 1, 11, 20
6                     Text(i) = "99-99-9999"
7                 Case 0
8                 Case Else
9                     Text(i) = ""
10            End Select
11        Next
          
12        If rsJob.State Then rsJob.Close
13        Set rsJob = Nothing
14        If rsPart.State Then rsPart.Close
15        Set rsPart = Nothing
16        If rsSublet.State Then rsSublet.Close
17        Set rsSublet = Nothing
18        CRViewer1.Visible = False
19        cmdTutup.Visible = False
20        lblClose.Visible = False
21        Text(1) = TglDDMMYYYY(Format(isnow, "dd/MM/yyyy"))
22        TblSimpan.Visible = False
23        nilaiRupiah = 0

24       On Error GoTo 0
25       Exit Sub

hapusForm_Error:

26        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmRegistrasi")
End Sub

Private Sub doCreatePart()
          
1        On Error GoTo doCreatePart_Error
2         If rsPart.State Then rsPart.Close
3         'Set rsPart = Nothing
4         With rsPart
              'If .State Then .Close
5             .Fields.Append "KODEITEM", adVarChar, 50
6             .Fields.Append "NAMAITEM", adVarChar, 200
7             .Fields.Append "QTY", adInteger
8             .Open
9         End With

10       On Error GoTo 0
11       Exit Sub

doCreatePart_Error:

12        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreatePart of Form frmRegistrasi")
End Sub

Private Sub doHeadPart()
          
1        On Error GoTo doHeadPart_Error

2         With partGrid
3             .Columns(0).Width = 1800
4             .Columns(0).Caption = "KODE PART"
5             .Columns(0).DataField = "KODEITEM"
6             .Columns(1).Width = 2500
7             .Columns(1).Caption = "NAMA PART"
8             .Columns(1).DataField = "NAMAITEM"
9             .Columns(2).Width = 800
10            .Columns(2).Caption = "QTY"
11            .Columns(2).DataField = "QTY"
12        End With

13       On Error GoTo 0
14       Exit Sub

doHeadPart_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadPart of Form frmRegistrasi")
End Sub

Private Sub doCreateJob()
          
1        On Error GoTo doCreateJob_Error
2         If rsJob.State Then rsJob.Close
3         'Set rsJob = Nothing
4         With rsJob
              'If .State Then .Close
5             .Fields.Append "JOB", adVarChar, 50
6             .Fields.Append "RATE", adDouble
7             .Fields.Append "GRP", adVarChar, 6
8             .Fields.Append "HRG", adDouble
9             .Fields.Append "fOther", adBoolean
10            .Open
11        End With

12       On Error GoTo 0
13       Exit Sub

doCreateJob_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreateJob of Form frmRegistrasi")
End Sub

Private Sub doHeadJob()
          
1        On Error GoTo doHeadJob_Error

2         With JobGrid
3             .Columns(0).Width = 4500
4             .Columns(0).Caption = "JOB ORDER"
5             .Columns(0).DataField = "JOB"
6             .Columns(1).Width = 800
7             .Columns(1).Caption = "RATE"
8             .Columns(1).DataField = "RATE"
9             .Columns(2).Caption = "GROUP"
10            .Columns(2).DataField = "GRP"
11            .Columns(2).Width = 0
12            .Columns(2).Visible = False

13            .Columns(3).Caption = "HARGA"
14            .Columns(3).DataField = "HRG"
15            .Columns(3).Width = 0
16            .Columns(3).Visible = False

17            .Columns(4).Caption = "FOTHER"
18            .Columns(4).DataField = "fOther"
19            .Columns(4).Width = 0
20            .Columns(4).Visible = False
21        End With

22       On Error GoTo 0
23       Exit Sub

doHeadJob_Error:

24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadJob of Form frmRegistrasi")
End Sub

Private Sub doHeadSublet()

1        On Error GoTo doHeadSublet_Error

2         With subletGrid
3             .Columns(0).Width = 1000
4             .Columns(0).Caption = "NO NOTA"
5             .Columns(0).DataField = "NONOTA"
6             .Columns(1).Width = 2000
7             .Columns(1).Caption = "KODE ITEM"
8             .Columns(1).DataField = "KODEITEM"
9             .Columns(1).Visible = False
10            .Columns(2).Width = 2500
11            .Columns(2).Caption = "NAMA ITEM"
12            .Columns(2).DataField = "NAMABARANG"
13            .Columns(3).Width = 1000
14            .Columns(3).Caption = "SATUAN"
15            .Columns(3).DataField = "KODESATUAN"
16            .Columns(4).Width = 1000
17            .Columns(4).Caption = "QTY"
18            .Columns(4).DataField = "QTY"
19            .Columns(4).NumberFormat = "#,#0"
20            .Columns(5).Width = 2000
21            .Columns(5).Caption = "HARGA"
22            .Columns(5).DataField = "HARGA"
23            .Columns(5).NumberFormat = "#,#0"
24        End With

25       On Error GoTo 0
26       Exit Sub

doHeadSublet_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadSublet of Form frmRegistrasi")
End Sub

Private Sub panggilWO(nWO As String)
      Dim rsHlp As ADODB.Recordset
          
1        On Error GoTo panggilWO_Error

2         If okCallFinishService(nWO) Then
3             Set rsTemp = objO.FetchData("SELECT DISTINCT A.*, B.GSM, B.CDMA, C.SLIP, C.NOINV, C.TOTAL FROM JOBDTA A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS " _
                                       & "LEFT JOIN T_INVOICE C WITH (NOLOCK) ON A.NO_WO = C.NO_WO WHERE A.NO_wo = '" & nWO & "' AND A.DELIV <> 'BTL' AND C.DELIV <> 'BTL' AND C.FS_LOG_DEL = '' AND A.NOINV <> '' AND (A.RPR_COND = 'S' OR A.RPR_COND = 'K') ") ' AND A.NOINV <> ''
4             If Not rsTemp.EOF Then
5                 TblSimpan.Visible = True
6                 cmdGate.Visible = True
7                 cmdGateS.Visible = False
                  
8                 If rsTemp("RPR_COND") = "S" Then
9                     cmdGate.Visible = False
10                    cmdGateS.Visible = True
11                    lblClose.Visible = False
12                Else
13                    TblSimpan.Visible = False
14                    lblClose.Visible = True
15                End If
                  
                  ' TAMPIL DATA CUSTOMER
16                Text(2) = rsTemp("FS_KD_PLG")
17                Text(3) = rsTemp("FS_KD_WIL")
18                Text(4) = rsTemp("CUSTOMER")
19                Text(5) = Left(Replace(rsTemp("CUST_ADRES"), vbNewLine, ""), 60)
20                Text(6) = rsTemp("KOTA")
21                Text(7) = rsTemp("TELPON")
22                Text(8) = rsTemp("GSM")
23                Text(9) = rsTemp("CDMA")
24                Text(10) = rsTemp("FS_TH_MOBIL")
25                Text(11).Mask = ""
26                Text(11) = rsTemp("DOTAM")
27                Text(12) = rsTemp("MDL")
28                Text(13) = rsTemp("MODEL")
29                Text(14) = rsTemp("FMDL")
30                Text(15) = rsTemp("COLOR")
31                Text(16) = rsTemp("NO_POLISI")
32                Text(17) = rsTemp("NO_CHASSIS")
33                Text(18) = rsTemp("NO_ENGINE")
                  
                  ' TAMPIL DATA SERVICE
                  'text(20).Mask = ""
34                Text(20) = TglDDMMYYYY(rsTemp("DATE"))
35                Text(21) = rsTemp("KM")
36                Text(22) = rsTemp("NOINV")
37                Text(23) = rsTemp("MECH")
38                Text(24) = rsTemp("FS_LOG_ADD")
39                Text(25) = IIf(rsTemp("TPSVC") <> "", rsTemp("KODETIPE"), "REGULAR")
40                Text(29) = cekMOU(rsTemp("MDL"), Nul(rsTemp("NO_CHASSIS")))
                  
41                If Trim(rsTemp("TPSVC")) = "W" Then
                      
42                    If Log.RPLHistori = 0 Then
43                        Set rsHlp = objO.FetchData("SELECT * FROM HO_WARRANTY WITH (NOLOCK) WHERE ID_WARRANTY = '" & rsTemp("KODETIPE") & "' and FS_LOG_DEL = ''")
44                    Else
45                        Set rsHlp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH (NOLOCK) WHERE ID_WARRANTY = '" & rsTemp("KODETIPE") & "' and FS_LOG_DEL = ''")
46                    End If
47                    If Not rsHlp.EOF Then
48                        Text(26) = rsHlp("WARRANTY")
49                    End If
50                ElseIf Trim(rsTemp("TPSVC")) = "P" Then
                      
51                    If Log.RPLHistori = 0 Then
52                        Set rsHlp = objO.FetchData("SELECT * FROM HO_PROGRAM WITH (NOLOCK) WHERE ID_PROGRAM = '" & rsTemp("KODETIPE") & "' and FS_LOG_DEL = ''")
53                    Else
54                        Set rsHlp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_PROGRAM WITH (NOLOCK) WHERE ID_PROGRAM = '" & rsTemp("KODETIPE") & "' and FS_LOG_DEL = ''")
55                    End If
56                    If Not rsHlp.EOF Then
57                        Text(26) = rsHlp("PROGRAM")
58                    End If
59                End If
                  
60                myChassis = Trim(Text(17))
                  Dim nDisN As Double
61                If CexChassis_Discon(Trim(Text(13))) = True Then
62                    nDisN = (100 - nDiscNRM) / 100
63                Else
64                    nDisN = (100 - nDiscNON) / 100
65                End If
                  ' TAMPIL JOB ORDER
66                If rsJob.State Then rsJob.Close
67                For i = 1 To 7
68                    hlp = "REP" & i
69                    hlp2 = "RTE" & i
70                    hlp3 = "KLPKERJA" & i

71                    If rsTemp(hlp) <> "" Then
72                        Set rsHlp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "' AND MDL = '" & Trim(rsTemp("MDL")) & "'")
73                        If Not rsHlp.EOF Then
74                            fOth = False
75                        Else
76                            fOth = True
77                        End If

                          Dim HrgCar As Double
                           
78                        If Log.RPLHistori = 0 Then
79                            Set rsHlp = objO.FetchData("SELECT price FROM carclass WITH (NOLOCK) WHERE CLASSCODE = '" & rsTemp("MODEL") & "'")
80                        Else
81                            Set rsHlp = objO.FetchData("SELECT price FROM [" & dbReplikasiHO & "]..carclass WITH (NOLOCK) WHERE CLASSCODE = '" & rsTemp("MODEL") & "'")
82                        End If
83                        If Not rsHlp.EOF Then
84                            HrgCar = CDbl(rsHlp("PRICE")) * nDisN
85                        Else
86                            HrgCar = 0
87                        End If

88                        With rsJob
89                        If .State = adStateClosed Then doCreateJob
90                           .AddNew
91                           !job = rsTemp(hlp)
92                           !Rate = rsTemp(hlp2)
93                           !grp = rsTemp(hlp3)
94                           !hrg = HrgCar * CDbl(rsTemp(hlp2))
95                           !fOther = fOth
96                           .Update
97                        End With
98                        rsHlp.Close
99                    End If
100               Next

                  ' add JOB
                  Dim rsHlp2 As ADODB.Recordset
101               Set rsHlp2 = objO.FetchData("SELECT * FROM HIST_JOBADD WITH (NOLOCK) WHERE NO_WO = '" & nWO & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
102               If Not rsHlp2.EOF Then
103                   Do While Not rsHlp2.EOF
104                       With rsJob
105                       If .State = adStateClosed Then doCreateJob
106                           .AddNew
107                           !job = Trim(rsHlp2("REP"))
108                           !Rate = CDbl(rsHlp2("RTE"))
109                           !grp = Trim(rsHlp2("KLPKERJA"))
110                           !hrg = CDbl(HrgCar) * CDbl(rsHlp2("RTE"))
111                           .Update
112                       End With
113                       rsHlp2.MoveNext
114                   Loop
115               End If
116               Set JobGrid.DataSource = rsJob
117               Call doHeadJob
                  
                  ' TAMPIL PART ORDER
118               Set rsHlp = objO.FetchData("SELECT KODE_ITEM, NAMA_ITEM, PARTBAWA, QTY, isnull(fs_kd_oli,'') as OLI, BAWA FROM jobpart a WITH (NOLOCK) LEFT JOIN T_oli b WITH (NOLOCK) ON LEFT(a.kode_item,1) = b.fs_kd_Oli AND B.TIPE = 'O' WHERE no_WO = '" & nWO & "' AND no_Chassis = '" & rsTemp("NO_CHASSIS") & "' AND GANTI = 'Y'")
119               If Not rsHlp.EOF Then
120                   Do While Not rsHlp.EOF
121                       With rsPart
122                       If .State = adStateClosed Then doCreatePart
123                       .AddNew
124                       !KodeItem = IIf(rsHlp("BAWA") = "Y", "BAWA SENDIRI", rsHlp("KODE_ITEM"))
125                       !NAMAITEM = IIf(rsHlp("BAWA") = "Y", rsHlp("PARTBAWA"), rsHlp("NAMA_ITEM"))
126                       !qty = rsHlp("QTY")
127                       .Update
128                       End With
129                       rsHlp.MoveNext
130                   Loop
131               End If
132               rsHlp.Close
                  
133               Set rsHlp = objO.FetchData("SELECT KODEITEM, KETERANGAN, QTY, HARGA, OLI FROM T_PART WITH (NOLOCK) WHERE no_WO = '" & nWO & "'")
134               If Not rsHlp.EOF Then
135                   Do While Not rsHlp.EOF
136                       With rsPart
137                           If .State = adStateClosed Then doCreatePart
138                           .AddNew
139                           !KodeItem = Trim(rsHlp("KODEITEM"))
140                           !NAMAITEM = Trim(rsHlp("KETERANGAN"))
141                           !qty = CDbl(rsHlp("QTY"))
142                           .Update
143                       End With
144                       rsHlp.MoveNext
145                   Loop
146               End If
147               rsHlp.Close
148               Set partGrid.DataSource = rsPart
149               Call doHeadPart
                  
                  ' TAMPIL SUBLET ORDER
150               Set rsSublet = objO.FetchData("SELECT NONOTA, KODEITEM, NAMABARANG, KODESATUAN, QTY, PRICE as HARGA FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & nWO & "' AND no_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
151               If Not rsSublet.EOF Then
152                   Set subletGrid.DataSource = rsSublet
153                   Call doHeadSublet
154               End If
                  
155               If rsTemp("NOINV") = "" Then
156                   TblSimpan.Visible = False
157                   cmdGate.Visible = False
158                   cmdGateS.Visible = True
159               End If
                  
160               nilaiRupiah = 0
161               Do While Not rsTemp.EOF
162                   If Right(rsTemp("NOINV"), 2) <> "-B" Then
163                       nilaiRupiah = nilaiRupiah + CDbl(rsTemp("TOTAL"))
164                   End If
165                   rsTemp.MoveNext
166               Loop
167           Else
168               TblSimpan.Visible = False
169           End If
170       Else
171           MsgBox "Customer belum dipanggil melalui Menu Panggilan Selesai Service. " & vbCr & "Hubungi SA bersangkutan ( " & cariNamaSA(Trim(Text(0))) & " )", vbCritical, "Panggil Selesai Service"
172       End If

173      On Error GoTo 0
174      Exit Sub

panggilWO_Error:

175       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilWO of Form frmRegistrasi")
End Sub

Private Function cariNamaSA(noWO As String) As String
          
1        On Error GoTo cariNamaSA_Error

2         cariNamaSA = ""
3         Set rsTemp = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) LEFT JOIN T_PETUGAS B WITH (NOLOCK) ON A.FRONT = B.FS_KD_PETUGAS WHERE NO_WO = '" & Trim(noWO) & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = ''")
4         If Not rsTemp.EOF Then
5             cariNamaSA = Trim(rsTemp("FRONT")) & " | " & Trim(rsTemp("FS_NM_PETUGAS"))
6         End If
7         If rsTemp.State <> adStateClosed Then rsTemp.Close

8        On Error GoTo 0
9        Exit Function

cariNamaSA_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariNamaSA of Form frmRegistrasi")
          
End Function

Private Sub BuatTabelRepoKasir(vNamaTabel As String)
1        On Error GoTo BuatTabelRepoKasir_Error

2         sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]') ) " _
               & " drop table [tempdb]..[" & vNamaTabel & "]"
3         Call objO.InsertData(sSQL)
          
4         sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
               & "    [SA][char](5) NOT NULL DEFAULT '', " _
               & "    [NOINV][varchar](15) NOT NULL DEFAULT '', " _
               & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
               & "    [NO_POLISI][varchar](12) NOT NULL DEFAULT '', " _
               & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
               & "    [JOR][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [JRTE][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [RTJ][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [THS][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [SBI][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [SBE][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [GRP][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [TWC][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [LUB][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [PDI][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [WAS][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [BOR][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [MATERIAL][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [PART][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [OLI][decimal](18,2) NOT NULL DEFAULT 0, "
               
5         sSQL = sSQL + "[SUBLET][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [SUBTOT][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [METERAI][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [TOTAL][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [TIPEINV][varchar](20) NOT NULL DEFAULT '', " _
               & "    [CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
               & "    [LABOR][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [NOFAKTUR][varchar](30) NOT NULL DEFAULT ''," _
               & "    [HPP][decimal](18,2) NOT NULL DEFAULT 0," _
               & "    [INVDATE][datetime] NOT NULL DEFAULT '01/01/3000') ON [PRIMARY]"
6         Call objO.UpdateData(sSQL)

7        On Error GoTo 0
8        Exit Sub

BuatTabelRepoKasir_Error:

9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure BuatTabelRepoKasir of Form frmRegistrasi")
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
      Dim rsHlp As ADODB.Recordset

1        On Error GoTo createTemp_Error

2         createTemp = False
3         Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]') ")
4         If Not rsHlp.EOF Then
5             createTemp = True
6         Else
7             Call BuatTabelRepoKasir(vNamaTabel)
8             createTemp = True
9         End If

10       On Error GoTo 0
11       Exit Function

createTemp_Error:

12        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createTemp of Form frmRegistrasi")
End Function

Private Function statusDealer(codeStatus As Integer)
          
1        On Error GoTo statusDealer_Error

2         Select Case codeStatus
              Case 1
3                 statusDealer = "AFFNONPPN"
4             Case 2
5                 statusDealer = "AFFPPN"
6             Case 3
7                 statusDealer = "EXT"
8         End Select

9        On Error GoTo 0
10       Exit Function

statusDealer_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure statusDealer of Form frmRegistrasi")
End Function

Private Function cekDuo(hrg As Double, nINV As String, Optional stat As Integer)
          
1        On Error GoTo cekDuo_Error

2         If Right(Trim(nINV), 2) = "-B" Then
3             cekDuo = 0
4             If stat = 1 Then
5                 cekDuo = CDbl(hrg)
6             End If
7         Else
8             cekDuo = CDbl(hrg)
9             If stat = 2 Then
10                cekDuo = 0
11            End If
12        End If

13       On Error GoTo 0
14       Exit Function

cekDuo_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekDuo of Form frmRegistrasi")
End Function

Private Sub IsiRepo(vNamaTabel As String)
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim hlp6 As String
      Dim hlp7 As String
      Dim rsHlp As ADODB.Recordset
      Dim rsAdd As ADODB.Recordset
      Dim nNilai As Double
      Dim nDuoInv As Boolean
      Dim subletOnly As Boolean
      Dim noInv As String
      Dim nmCust As String
      Dim nStatus As String
      'Dim vNamaTabel As  String
          
1        On Error GoTo IsiRepo_Error

2         Call BuatTabelRepoKasir(vNamaTabel)
3         If Log.RPLHistori = 0 Then
4             Set rsTemp = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH (NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C WITH (NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE DELIV <> 'BTL' AND INVDATE BETWEEN '" & Format(DtAwal, "MM/DD/YYYY") & "' AND '" & Format(DtAkhir, "MM/DD/YYYY") & "' ORDER BY NOINV")
5         Else
6             Set rsTemp = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH (NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C WITH (NOLOCK) ON C.FS_LOG_DEL = '' AND A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE DELIV <> 'BTL' AND INVDATE BETWEEN '" & Format(DtAwal, "MM/DD/YYYY") & "' AND '" & Format(DtAkhir, "MM/DD/YYYY") & "' ORDER BY NOINV")
7         End If
8         If Not rsTemp.EOF Then
9             Do While Not rsTemp.EOF
10                subletOnly = True
11                For i = 1 To 8
12                    hlp = "REP" & i
13                    If rsTemp(hlp) <> "" Then
14                        subletOnly = False
15                        hlp2 = "KLPKERJA" & i
16                        hlp3 = "RTE" & i    '"ACTRATE" & i
17                        hlp4 = "MEKANIK" & i
18                        hlp5 = "OFF" & i
19                        hlp6 = "DISC" & i
20                        hlp7 = "FREEC" & i
21                        If createTemp(vNamaTabel) Then
22                            nDuoInv = False
23                            If rsTemp("NDUOINV") = 1 Then
24                                nDuoInv = True
25                            End If
                              '#########################################################'
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                              '#########################################################'
26                            nStatus = rsTemp("STATUS")
27                            nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
28                            noInv = Trim(rsTemp("NOINV"))
29                            nmCust = Trim(rsTemp("CUSTOMER"))

30                            If nStatus = "AFFNONPPN" Then
31                                nNilai = nNilai * (100 / 110)
32                            End If
                              
33                            Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
34                            If Not rsHlp.EOF Then
35                                Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                               & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "BOR = '" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', MATERIAL = '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', PART = '" & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', " _
                                               & "OLI = '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', SUBLET = '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', SUBTOT = '" & RepRp(cekDuo(CDbl(rsTemp("SUBTOT")), noInv)) & "', " _
                                               & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', HPP = '" & RepRp(cekDuo(CDbl(rsTemp("HPP")), noInv)) & "' " _
                                               & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
36                            Else
                                  
37                                Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                               & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "'," & RepRp(IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0)) & ", " & RepRp(IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0)) & ", " _
                                               & "'" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & " " & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & ",  " & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", " & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & ", " & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & ", " & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", " & RepRp(cekDuo(CDbl(rsTemp("SUBTOT")), noInv)) & ", " _
                                               & " " & RepRp(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & ", " & RepRp(CDbl(rsTemp("TOTALL"))) & ", '" & Trim(nStatus) & "', '" & Trim(nmCust) & "'," & Replace(CDbl(rsTemp("LABOR")), ",", ".") & ", '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', '" & cekDuo(CDbl(rsTemp("HPP")), noInv) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
38                            End If
                               
                              
39                        End If
40                    End If
41                Next
42                If subletOnly Then
43                    nDuoInv = False
44                    If rsTemp("NDUOINV") = 1 Then
45                        nDuoInv = True
46                    End If
                      '#########################################################'
                      '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                      '#########################################################'
47                    noInv = Trim(rsTemp("NOINV"))
48                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
49                    If Not rsHlp.EOF Then
50                        Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                       & "JOR = '0', TIPEINV = '" & Trim(rsTemp("STATUS")) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', BOR = '" & RepRp(CDbl(rsTemp("BODY"))) & "', MATERIAL = '" & RepRp(CDbl(rsTemp("MTRL"))) & "', " _
                                       & "PART = '" & RepRp(CDbl(rsTemp("PART_G") + rsTemp("PART_NG"))) & "', OLI = '" & RepRp(CDbl(rsTemp("OIL_G") + rsTemp("OIL_NG"))) & "', SUBLET = '" & RepRp(CDbl(rsTemp("SBLT"))) & "', SUBTOT = '" & RepRp(CDbl(rsTemp("SUBTOT"))) & "', " _
                                       & "LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', HPP = '" & RepRp(CDbl(rsTemp("HPP"))) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
51                    Else
52                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                       & "'0','0', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                       & "'" & RepRp(CDbl(rsTemp("BODY"))) & "', '" & RepRp(CDbl(rsTemp("MTRL"))) & "', '" & RepRp(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))) & "', '" & RepRp(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))) & "', '" & RepRp(CDbl(rsTemp("SBLT"))) & "', '" & RepRp(CDbl(rsTemp("SUBTOT"))) & "', " _
                                       & " '" & RepRp(CDbl(rsTemp("MET"))) & "', '" & RepRp(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(rsTemp("STATUS")) & "', '" & Trim(rsTemp("CUSTOMER")) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & Trim(rsTemp("NOFAKTUR")) & "','" & RepRp(CDbl(rsTemp("HPP"))) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
53                    End If
54                End If
                  
                  ' DARI ADD JOB
55                Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND FS_LOG_DEL = '" & Trim(rsTemp("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
56                If Not rsAdd.EOF Then
57                    Do While Not rsAdd.EOF
58                        hlp2 = "KLPKERJA"
59                        hlp3 = "RTE"
60                        hlp4 = "MEKANIK"
61                        hlp5 = "OFFN"
62                        hlp6 = "FREEC"
                          
63                        nDuoInv = False
64                        If rsTemp("NDUOINV") = 1 Then
65                            nDuoInv = True
66                        End If
                          '#########################################################'
                          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                          '#########################################################'
67                        nStatus = rsTemp("STATUS")
68                        nNilai = CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsTemp("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
69                        noInv = Trim(rsTemp("NOINV"))
70                        nmCust = Trim(rsTemp("CUSTOMER"))
                          
71                        If Trim(nStatus) = "AFFNONPPN" Then
72                            nNilai = nNilai * (100 / 110)
73                        End If
                          
74                        Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
75                        If Not rsHlp.EOF Then
76                            Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                           & "JOR = 0, JRTE = JRTE + " & Replace(CDbl(rsAdd(hlp3)), ",", ".") & ", SBI = SBI + " & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & ", " _
                                           & "SBE = SBE + " & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & ", GRP = GRP + " & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & ", " _
                                           & "TWC = TWC + " & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & ", LUB = LUB + " & IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & ", " _
                                           & "PDI = PDI + " & IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & ", WAS = WAS + " & IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & ", " _
                                           & "BOR = " & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & ", MATERIAL = " & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", PART = " & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & ", " _
                                           & "OLI = " & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & ", SUBLET = " & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", SUBTOT = " & RepRp(CDbl(rsTemp("SUBTOT"))) & ", HPP = " & RepRp(cekDuo(CDbl(rsTemp("HPP")), noInv)) & ", " _
                                           & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = " & Replace(CDbl(rsTemp("LABOR")), ",", ".") & ", NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "' " _
                                           & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
77                        Else
78                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                           & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "SBE", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "LUB", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS", CDbl(nNilai), 0) & "', " _
                                           & "'" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', '" & RepRp(CDbl(rsTemp("SUBTOT"))) & "', " _
                                           & " '" & RepRp(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & "', '" & RepRp(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "','" & RepRp(cekDuo(CDbl(rsTemp("HPP")), noInv)) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
79                        End If
                           
80                        rsAdd.MoveNext
81                    Loop
82                End If
83                If rsAdd.State <> adStateClosed Then rsAdd.Close
84                rsTemp.MoveNext
85            Loop
86        End If

87       On Error GoTo 0
88       Exit Sub

IsiRepo_Error:

89        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure IsiRepo of Form frmRegistrasi")
End Sub

Private Sub cmdCetak_Click()
1        On Error GoTo cmdCetak_Click_Error

2         Set Report = New rptLaporanKasir
          Dim rsRepo As ADODB.Recordset
          Dim vNamaTabel As String
          Dim hlp As String
3         vNamaTabel = "KASIR"
          
4         Call IsiRepo(vNamaTabel)
          '#########################################################'
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
          '#########################################################'
5         hlp = ""
6         If chInvB.value = False Then
7             hlp = " WHERE RIGHT(A.NOINV,2) <> '-B' "
8         End If
9         Set rsRepo = objO.FetchData("SELECT A.NOINV, (CASE (CASE RIGHT(A.NOINV,2) WHEN '-B' THEN 1 ELSE (CASE WHEN CASHADD = '' THEN (CASE WHEN CONVERT(VarChar(15), B.TGLKASIR, 1) <> B.INVDATE THEN 1 ELSE FS_KREDIT END) ELSE FS_KREDIT END) END) WHEN 1 THEN 0 ELSE SBI+A.SBE+GRP+TWC+LUB+PDI+WAS+BOR+MATERIAL+A.PART+OLI+SUBLET END) AS CASH, " & _
                                     "(CASE (CASE RIGHT(A.NOINV,2) WHEN '-B' THEN 1 ELSE (CASE WHEN CASHADD = '' THEN (CASE WHEN CONVERT(VarChar(15), B.TGLKASIR, 1) <> B.INVDATE THEN 1 ELSE FS_KREDIT END) ELSE FS_KREDIT END) END) WHEN 1 THEN SBI+A.SBE+GRP+TWC+LUB+PDI+WAS+BOR+MATERIAL+A.PART+OLI+SUBLET ELSE 0 END) AS KREDIT, A.NO_POLISI, ALASAN_GATE, B.INVDATE " & _
                                     "FROM [tempdb]..[" & vNamaTabel & "] A INNER JOIN T_INVOICE B ON A.NOINV = B.NOINV INNER JOIN JOBDTA C ON LEFT(B.NOINV,8) = C.NOINV " & hlp & " ORDER BY A.NOINV")
                                     
10        If Not rsRepo.EOF Then
11            Report.txtTanggal.SetText Format(DtAwal, "DD/MM/YYYY") & " - " & Format(DtAkhir, "DD/MM/YYYY")
12            Report.database.SetDataSource rsRepo
13            With CRViewer1
14                .ReportSource = Report
15                .ViewReport
16                .Zoom 100
                  
17                .Top = 0
18                .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
19                .Width = Frame6.Width + 100
20                .Height = Frame6.Top + Frame6.Height
                  
21                cmdTutup.Top = .Top + 430
22                cmdTutup.Left = .Left
23                cmdTutup.Visible = True
                  
24                cmdSettingPrinter.Top = cmdTutup.Top
25                cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
26                cmdSettingPrinter.Visible = True
              
27                .Visible = True
28            End With
29            Screen.MousePointer = vbDefault
30            frmLaporan.Visible = False
31        Else
32            MsgBox "Belum ada data !!"
33        End If

34       On Error GoTo 0
35       Exit Sub

cmdCetak_Click_Error:

36        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetak_Click of Form frmRegistrasi")
End Sub

Private Sub cmdClose_Click()
    alasanPass.Visible = False
End Sub

Private Sub cmdCloseLap_Click()
    frmLaporan.Visible = False
End Sub

Private Sub cmdGate_Click()
1        On Error GoTo cmdGate_Click_Error

2     Set Report = New RptGatePass
      Dim rsRepo As ADODB.Recordset
          
      '    sSQL = "SELECT NO_WO, NOINV, NO_POLISI, NO_CHASSIS, MDL, FMDL, CUSTOMER, CUST_ADRES FROM T_INVOICE WHERE NOINV = '" & Text(22) & "' AND NO_WO = '" & Text(0) & "'"
          '#########################################################'
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
          '#########################################################'
3         sSQL = "SELECT A.NO_WO, A.NOINV, A.NO_POLISI, A.NO_CHASSIS, A.MDL, A.FMDL, B.CUSTOMER, B.CUST_ADRES FROM T_INVOICE a WITH (NOLOCK) INNER JOIN JOBDTA b WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND LEFT(A.NOINV,8) = B.NOINV WHERE A.NOINV LIKE '" & Text(22) & "%' AND A.NO_WO = '" & Text(0) & "'"
4         Set rsRepo = objO.FetchData(sSQL)
5         Report.database.SetDataSource rsRepo
6         Report.txtPT.SetText Trim(gPerush)
7         Report.txtAlamat.SetText Trim(gAlamat) & "  " & Trim(gKotaCbg)
8         Report.txtKabeng.SetText gSecmanager

9         With CRViewer1
10            .ReportSource = Report
11            .ViewReport
12            .Zoom 100
              
13            .Top = 0
14            .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
15            .Width = Frame6.Width + 100
16            .Height = Me.Height
              
17            cmdTutup.Top = .Top + 430
18            cmdTutup.Left = .Left
19            cmdTutup.Visible = True
              
20            cmdSettingPrinter.Top = cmdTutup.Top
21            cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
22            cmdSettingPrinter.Visible = True
23            .Visible = True
24        End With
25        Screen.MousePointer = vbDefault
26        cmdTutup.SetFocus

27       On Error GoTo 0
28       Exit Sub

cmdGate_Click_Error:

29        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdGate_Click of Form frmRegistrasi")
End Sub

Private Sub cmdGateS_Click()
1        On Error GoTo cmdGateS_Click_Error

2         alasanPass.Visible = True

3        On Error GoTo 0
4        Exit Sub

cmdGateS_Click_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdGateS_Click of Form frmRegistrasi")
End Sub

Private Sub cmdLap_Click()
1        On Error GoTo cmdLap_Click_Error

2         frmLaporan.Visible = True
3         chInvB.value = False
4         DtAwal = Format(isnow, "DD/MM/YYYY")
5         DtAkhir = Format(isnow, "DD/MM/YYYY")

6        On Error GoTo 0
7        Exit Sub

cmdLap_Click_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdLap_Click of Form frmRegistrasi")
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    Call hapusForm
End Sub

Private Sub Form_Load()
    Call hapusForm
    simpan = True
    'text(1) = TglDDMMYYYY(Format(isnow, "dd/MM/yyyy"))
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub tblCetakPassSementara_Click()
1        On Error GoTo tblCetakPassSementara_Click_Error

2         Set Report = New RptGatePassSementara
          Dim rsRepo As ADODB.Recordset
          
3         If Trim(Text1) <> "" And Len(Trim(Text1)) > 5 Then
4             objO.UpdateData ("UPDATE JOBDTA SET ALASAN_GATE = '" & Trim(Text1) & "' WHERE NOINV = '" & Text(22) & "' AND NO_WO = '" & Text(0) & "'")
5             objO.UpdateData ("UPDATE T_INVOICE SET FS_KREDIT = '1', TGLKASIR = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "' WHERE NOINV = '" & Text(22) & "' AND NO_WO = '" & Text(0) & "' AND DELIV <> 'BTL'")
              
6             Set rsRepo = objO.FetchData("SELECT NO_WO, NOINV, NO_POLISI, NO_CHASSIS, MDL, FMDL, CUSTOMER, CUST_ADRES, ALASAN_GATE FROM JOBDTA WITH (NOLOCK) WHERE NOINV = '" & Text(22) & "' AND NO_WO = '" & Text(0) & "'")
7             If Not rsRepo.EOF Then
8                 Report.database.SetDataSource rsRepo
9                 Report.txtPT.SetText Trim(gPerush)
10                Report.txtAlamat.SetText Trim(gAlamat) & "  " & Trim(gKotaCbg)
11                Report.txtKabeng.SetText gSecmanager
              
12                With CRViewer1
13                    .ReportSource = Report
14                    .ViewReport
15                    .Zoom 100
                      
16                    .Top = 0
17                    .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
18                    .Width = Frame6.Width + 100
19                    .Height = Me.Height
                      
20                    cmdTutup.Top = .Top + 430
21                    cmdTutup.Left = .Left
22                    cmdTutup.Visible = True
                      
23                    cmdSettingPrinter.Top = cmdTutup.Top
24                    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
25                    cmdSettingPrinter.Visible = True
                  
26                    .Visible = True
27                End With
28                Screen.MousePointer = vbDefault
29                alasanPass.Visible = False
30                Text1 = ""
31            End If
32        Else
33            MsgBox "Alasan harus ditulis Lengkap !!"
34        End If

35       On Error GoTo 0
36       Exit Sub

tblCetakPassSementara_Click_Error:

37        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblCetakPassSementara_Click of Form frmRegistrasi")
End Sub

Private Sub TblKeluar_Click()
    Unload Me
    If oRs.State <> adStateClosed Then oRs.Close
End Sub

Private Sub tblSimpan_Click()
    
    On Error GoTo vbError
    If Trim(Text(20)) <> "" And Trim(Text(0)) <> "" Then
        If MsgBox("Pembayaran Customer Tunai ?", vbDefaultButton1 + vbYesNo, "Pembayaran") = vbYes Then
            Call saveLogMenu(gSA, ComptName, gVersi, "frmRegistrasi", "tblSimpan", 0)
            
            erTranAll = False
            onTR = True
            Call objHT.OpenTransactionBengkel
                'TANGGAL 21-01-2015 RUDI
                Call objHT.UpdateData_BengkelOnly("UPDATE CHASSINGFORM SET TANGGALREGISTRASI = GETDATE() WHERE NOWO = '" & Text(0) & "'")
                Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET RPR_COND = 'K' WHERE NO_WO = '" & Text(0) & "' AND NOINV = '" & Text(22) & "' AND DELIV <> 'BTL'")
                Call objHT.UpdateData_BengkelOnly("UPDATE T_INVOICE SET FS_KREDIT = '0', RPR_COND = 'K', TGLKASIR = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "', KASIRADD = '" & Trim(gSA) & "' WHERE NO_WO = '" & Text(0) & "' AND NOINV LIKE '" & Text(22) & "%' AND DELIV <> 'BTL'")
                Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = 'K' WHERE KETERANGAN LIKE '%|" & Text(0) & "|%' ") 'AND TANGGAL = '" & Format(Text(20), "MM/DD/YYYY") & "'
                
                If Log.DiDxP1 = 1 Then
                    objHT.InsertData_Bengkel ("UPDATE [" & dbReplikasiHO & "]..MonitorSvcProggres SET service_delivery_handover_datetime = GETDATE() WHERE wo_number = '" & Right(gDb, 2) & "." & Trim(Text(0)) & "'")
                    
                    sSQL = "INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,status,proses,event_no,tglentry) VALUES ('MONITORINGGR','" & Right(gDb, 2) & "." & Trim(Text(0)) & "','" & Right(gDb, 2) & "." & Trim(Text(0)) & "','','" & Right(gDb, 2) & "','BILLING','0','',GETDATE())"
                    Call objHT.InsertData_Bengkel(sSQL)
                End If
                If Trim(Text(29)) <> "" Then
                    
                    If Log.RPLHistori = 0 Then
                        Call objHT.InsertData_BengkelOnly("INSERT INTO MOUTRN VALUES ('" & Format(isnow, "MM/DD/YYYY") & "', '" & Text(29) & "', '" & Text(0) & "', '" & Text(22) & "', '" & Text(17) & "', '" & Text(16) & "', '" & gKdCbg & "', " & CDbl(nilaiRupiah) & ", 'B','','01/01/3000')")
                        Call objHT.UpdateData_BengkelOnly("UPDATE MOUHDR SET BAYAR = BAYAR + " & CDbl(nilaiRupiah) & " WHERE NO_MOU = '" & Text(29) & "'")
                    Else
                        Call objHT.InsertData_BengkelOnly("INSERT INTO [" & dbReplikasiHO & "]..MOUTRN (TANGGAL,NO_MOU,NO_WO,NOINV,NO_CHASSIS,NO_POLISI,FS_KD_DEALER,TOTAL,STATUS,FS_LOG_DEL,FD_TGL_DEL) VALUES ('" & Format(isnow, "MM/DD/YYYY") & "', '" & Text(29) & "', '" & Text(0) & "', '" & Text(22) & "', '" & Text(17) & "', '" & Text(16) & "', '" & gKdCbg & "', " & CDbl(nilaiRupiah) & ", 'B','','01/01/3000')")
                        Call objHT.UpdateData_BengkelOnly("UPDATE [" & dbReplikasiHO & "]..MOUHDR SET BAYAR = BAYAR + " & CDbl(nilaiRupiah) & " WHERE NO_MOU = '" & Text(29) & "'")

                    End If
                End If
            
            If erTranAll = False Then
                Call objHT.CloseTransactionBengkel
                onTR = False
                Call saveLogMenu(gSA, ComptName, gVersi, "frmRegistrasi", "tblSimpan", 1)
            Else
                Call objHT.CloseTransactionBengkel_err
                onTR = False
                Call saveLogMenu(gSA, ComptName, gVersi, "frmRegistrasi", "tblSimpan", 2)
                MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
                Exit Sub
            End If
            
            MsgBox "WO no. " & Text(0) & " Telah diclosing "
            cmdGate_Click
            lblClose.Visible = True
            
            Text(0) = ""
            Text(0).SelLength = Len(Text(0))
            Text(0).SetFocus
        Else
            MsgBox "Isi alasan Kredit, dan detail pembayaran", vbInformation, "Kredit"
            cmdGateS_Click
        End If
    Else
        MsgBox "No WO salah !!"
        Text(0).SelLength = Len(Text(0))
        Text(0).SetFocus
    End If
    Exit Sub
    
vbError:
    Call objHT.CloseTransactionBengkel_err
    onTR = False
    Call saveLogMenu(gSA, ComptName, gVersi, "frmRegistrasi", "tblSimpan", 2)
    MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmRegistrasi"), vbInformation, "Gagal Simpan"
    Exit Sub
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
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
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "INV"
                gBantuan = "AND RPR_COND <> 'K' AND NOINV <> '' "
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As ADODB.Recordset

           
1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 0
3                 If Text(Index) <> "" Then
4                     Set rsSearch = objO.FetchData("SELECT * FROM jobDta WITH (NOLOCK) WHERE no_wo = '" & Text(Index) & "'")
5                     If rsSearch.EOF Then
6                         gMainKeyCari = Text(Index)
7                         gSearchOption = "INV"
8                         gBantuan = "AND (RPR_COND <> '' OR RPR_COND = 'S')"
9                         frmCari.show (vbModal)
10                    Else
11                        Call hapusForm
12                        Select Case Trim(rsSearch("RPR_COND"))
                              Case ""
13                                MsgBox "Unit Belum dikerjakan ! "
14                            Case "T"
15                                MsgBox "Pengerjaan unit dihentikan sementara ! "
16                            Case "P"
17                                MsgBox "Unit dalam Proses Pengerjaan ! "
18                            Case "S"
19                                If Trim(rsSearch("NOINV")) = "" Then
20                                    MsgBox "Unit Belum diinvoice ! "
21                                Else
22                                    Call panggilWO(Text(Index))
23                                End If
24                            Case "K"
25                                Call panggilWO(Text(Index))
26                        End Select
27                    End If
28                End If
29        End Select

30       On Error GoTo 0
31       Exit Sub

text_Validate_Error:

32        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmRegistrasi")
           
End Sub

Private Function returnJob(labelReturn As Object) As Boolean
1        On Error GoTo returnJob_Error

2         returnJob = False
3         If labelReturn.Visible Then
4             returnJob = True
5         End If

6        On Error GoTo 0
7        Exit Function

returnJob_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure returnJob of Form frmRegistrasi")
End Function

Private Function typeJob(statusJob As String) As String
          
1        On Error GoTo typeJob_Error

2         If statusJob = "W" Then
3             typeJob = "WARRANTY"
4         ElseIf statusJob = "P" Then
5             typeJob = "PROGRAM"
6         Else
7             typeJob = "REGULAR"
8         End If

9        On Error GoTo 0
10       Exit Function

typeJob_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure typeJob of Form frmRegistrasi")
End Function

