VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmBookingMenuOnline 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menu Booking Online"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16245
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   16245
   Begin TDBDate6Ctl.TDBDate tglAwal 
      Height          =   375
      Left            =   2400
      TabIndex        =   28
      Top             =   600
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calendar        =   "frmBookingMenuOnline.frx":0000
      Caption         =   "frmBookingMenuOnline.frx":012C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmBookingMenuOnline.frx":0198
      Keys            =   "frmBookingMenuOnline.frx":01B6
      Spin            =   "frmBookingMenuOnline.frx":0214
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
      Text            =   "07/08/2024"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   45511
      CenturyMode     =   0
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "JADI BOOKING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   24
      Top             =   600
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BATAL BOOKING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   23
      Top             =   600
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BELUM PROSES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   22
      Top             =   600
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Frame frameLegend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LEGEND WARNA"
      Height          =   1455
      Left            =   480
      TabIndex        =   18
      Top             =   6360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "JADI BOOKING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BATAL BOOKING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BELUM PROSES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3375
      End
   End
   Begin ucXPButton.XPButton btnLegend 
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   7800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBookingMenuOnline.frx":023C
      PICN            =   "frmBookingMenuOnline.frx":0258
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frameBatalBooking 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3360
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   5535
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3960
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
      Begin VB.TextBox txtAlasan 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   5295
      End
      Begin AIFCmp1.asxPowerButton btnSimpanBatal 
         Height          =   495
         Left            =   2640
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BackColor       =   33023
         Caption         =   "SIMPAN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16777215
      End
      Begin AIFCmp1.asxPowerButton btnBatalPembatalan 
         Height          =   495
         Left            =   4080
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BackColor       =   33023
         Caption         =   "BATAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16777215
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbCustomerMenolak 
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   1320
         Width           =   3015
         AllowInput      =   0   'False
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "-Pilih Alasan Customer Menolak -"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.OptionButton optTolak 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Customer Menolak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   3015
      End
      Begin VB.OptionButton optTolak 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tidak terhubung"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "PEMBATALAN BOOKING ONLINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   5055
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   5535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "PEMBATALAN BOOKING ONLINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   5055
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   15960
      TabIndex        =   1
      Top             =   120
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
   Begin TrueOleDBGrid70.TDBGrid gridTagih 
      Height          =   6015
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   10610
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
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
      Columns.Count   =   21
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=21"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=46,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=32,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=88,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=84,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=85,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=86,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=28,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=25,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=26,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=27,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=92,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=89,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=90,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=91,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=96,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=93,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=94,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=95,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=100,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=97,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=98,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=99,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=104,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=101,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=102,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=103,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=108,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=105,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=106,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=107,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=112,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=109,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=110,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=111,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=116,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=113,.parent=14"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=114,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=115,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=120,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=117,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=118,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=119,.parent=17"
      _StyleDefs(120) =   "Named:id=33:Normal"
      _StyleDefs(121) =   ":id=33,.parent=0"
      _StyleDefs(122) =   "Named:id=34:Heading"
      _StyleDefs(123) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(124) =   ":id=34,.wraptext=-1"
      _StyleDefs(125) =   "Named:id=35:Footing"
      _StyleDefs(126) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(127) =   "Named:id=36:Selected"
      _StyleDefs(128) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(129) =   "Named:id=37:Caption"
      _StyleDefs(130) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(131) =   "Named:id=38:HighlightRow"
      _StyleDefs(132) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(133) =   "Named:id=39:EvenRow"
      _StyleDefs(134) =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&"
      _StyleDefs(135) =   "Named:id=40:OddRow"
      _StyleDefs(136) =   ":id=40,.parent=33"
      _StyleDefs(137) =   "Named:id=41:RecordSelector"
      _StyleDefs(138) =   ":id=41,.parent=34"
      _StyleDefs(139) =   "Named:id=42:FilterBar"
      _StyleDefs(140) =   ":id=42,.parent=33"
      _StyleDefs(141) =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(142) =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(143) =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(144) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(145) =   "Named:id=87:SudahDitagih"
      _StyleDefs(146) =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(147) =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(148) =   ":id=87,.fontname=Tahoma"
      _StyleDefs(149) =   "Named:id=83:Booking"
      _StyleDefs(150) =   ":id=83,.parent=87,.fgcolor=&HFF00&"
   End
   Begin AIFCmp1.asxPowerButton tblImport 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "DATA ONLINE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton tblBatal 
      Height          =   495
      Left            =   12720
      TabIndex        =   7
      Top             =   7920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BackColor       =   33023
      Caption         =   "BATAL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton tblSimpan 
      Height          =   495
      Left            =   10440
      TabIndex        =   8
      Top             =   7920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "BUAT BOOKING"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      Height          =   495
      Left            =   7920
      TabIndex        =   9
      Top             =   7920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BackColor       =   255
      Caption         =   "BATAL BOOKING"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton3 
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BackColor       =   255
      Caption         =   "REFRESH "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin TDBDate6Ctl.TDBDate tglAkhir 
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   600
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calendar        =   "frmBookingMenuOnline.frx":4773
      Caption         =   "frmBookingMenuOnline.frx":489F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmBookingMenuOnline.frx":490B
      Keys            =   "frmBookingMenuOnline.frx":4929
      Spin            =   "frmBookingMenuOnline.frx":4987
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
      Text            =   "07/08/2024"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   45511
      CenturyMode     =   0
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "S / D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      Top             =   600
      Width           =   615
   End
   Begin VB.Label jmlTagih 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   9720
      TabIndex        =   6
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Yang Dipilih :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   3060
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: DAFTAR BOOKING ONLINE ::."
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
      Left            =   480
      TabIndex        =   2
      Top             =   1275
      Width           =   13545
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   13575
   End
End
Attribute VB_Name = "frmBookingMenuOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRecord As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns
Private jumlahBookingOnline As Integer
Dim nTotal As Integer

Private Sub FetchCombo()
    Adodc1.ConnectionString = SQLCONBengkel
    Adodc1.RecordSource = "select alasan,id from AlasanTolakanCust WITH (NOLOCK)"
    cbCustomerMenolak.DataSourceList = Adodc1
    cbCustomerMenolak.DataFieldList = "alasan"
    cbCustomerMenolak.Columns(0).Width = 4000
    cbCustomerMenolak.Columns(1).Visible = False
    cbCustomerMenolak.Columns(0).Caption = "Alasan Penolakan Customer"
End Sub

Private Sub asxPowerButton1_Click()
    If nTotal > 0 Then
        If nTotal > 1 Then
            MsgBox "Membatalkan booking hanya diperbolehkan 1 data", vbCritical
            Exit Sub
        Else
            If MsgBox("Apakan yakin membatalkan " & nTotal & " Pesan Booking?", vbYesNo) = vbYes Then
                frameBatalBooking.Visible = True
                If Log.hansaRenkeiBP = 1 Then
                    Call FetchCombo
                    txtAlasan.Visible = False
                    txtAlasan.Text = ""
                    btnSimpanBatal.Top = 1800
                    btnBatalPembatalan.Top = 1800
                    frameBatalBooking.Height = 2415
                Else
                    txtAlasan.Text = ""
                    btnSimpanBatal.Top = 2760
                    btnBatalPembatalan.Top = 2760
                    cbCustomerMenolak.Visible = False
                    frameBatalBooking.Height = 3375
                End If
            End If
        End If
    End If
End Sub

Public Sub asxPowerButton3_Click()
    If Log.RPLHistori = 0 Then
        Set rsRecord = objO.FetchDataHO("select count(*) from bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
    Else
        Set rsRecord = objO.FetchData("select count(*) from [" & dbReplikasiHO & "]..bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
    End If
    jumlahBookingOnline = 0
    If rsRecord.RecordCount > 0 Then
        jumlahBookingOnline = rsRecord.Fields(0).value
    End If
    tblImport.Caption = "DATA ONLINE ( " & jumlahBookingOnline & " )"
    Call hapusForm
    Call refreshGrid
    
End Sub

Private Sub btnBatalPembatalan_Click()
    txtAlasan.Text = ""
    frameBatalBooking.Visible = False
End Sub

Private Sub btnLegend_Click()
    If frameLegend.Visible Then
        frameLegend.Visible = False
    Else
        frameLegend.Visible = True
    End If
End Sub

Private Sub btnSimpanBatal_Click()
       Dim keteranganUpdate As String
       Dim evenID As String
       Dim cancelreason As String
       
       
1      If Log.hansaRenkeiBP = 1 Then
2         If optTolak(1).value And cbCustomerMenolak.value = "-Pilih Alasan Customer Menolak -" Then
3             MsgBox "Silahkan pilih alasan customer menolak", vbInformation
4             Exit Sub
5         End If
6      End If
       
      ' If Get_ConnectionProfil_Bengkel99 = False Then
      '    MsgBox "Koneksi HO terputus , proses tidak dapat dilanjutkan.", vbInformation
      '    Exit Sub
      ' End If
       
7      If Not rsRecord.EOF Then
8             rsRecord.MoveFirst
9             Do While Not rsRecord.EOF
10                If rsRecord("import") = True Then
11                    evenID = ""
12                    cancelreason = ""
13                    If NulSID(rsRecord!noBookingApps) <> "" And (Log.DiDxP1 = 1) Then
                          ''evenID = objO.FetchDataHO_String("select event_id from didx..bookingonline_didx_toSend where nobookingapps = '" & NulSID(rsRecord!noBookingApps) & "'")
14                        evenID = objO.FetchData_String("select event_id from [" & dbReplikasiHO & "]..bookingonline_didx_toSend where nobookingapps = '" & NulSID(rsRecord!noBookingApps) & "'")
15                    End If
                      
16                    If NulSID(rsRecord!noBookingApps) <> "" And (Log.p2Didx = 1) Then
                          ''evenID = objO.FetchDataHO_String("select event_id from didx..bookingonline_didx_toSend where nobookingapps = '" & NulSID(rsRecord!noBookingApps) & "'")
17                        evenID = objO.FetchData_String("select event_id from [" & dbReplikasiHO & "]..p2_bookingonline_didx where nobookingapps = '" & NulSID(rsRecord!noBookingApps) & "'")
18                    End If
                      
19                    If Log.hansaRenkeiBP = 1 Then
20                        If optTolak(0).value Then
21                            cancelreason = "TIDAK TERHUBUNG"
22                            Call objO.InsertData("UPDATE BOOKINGONLINE SET FPROSES = 3,HASILFU = '' WHERE ID = '" & rsRecord("Id") & "'")
23                            If Log.flagOffBookingLocal = 1 Then
24                                Call objO.InsertDataHO("UPDATE BENGKEL99.DBO.BOOKINGONLINE SET FPROSES = 3,HASILFU = '' WHERE ID = '" & rsRecord("Id") & "'")
25                            Else
26                                Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..BOOKINGONLINE SET FPROSES = 3,HASILFU = '' WHERE ID = '" & rsRecord("Id") & "'")
27                            End If
28                        Else
29                            cancelreason = NulSID(cbCustomerMenolak.Text)
30                            Call objO.InsertData("UPDATE BOOKINGONLINE SET FPROSES = 2,HASILFU = '" & NulSID(cbCustomerMenolak.Text) & "' WHERE ID = '" & rsRecord("Id") & "'")
31                            If Log.flagOffBookingLocal = 1 Then
32                                Call objO.InsertDataHO("UPDATE BENGKEL99.DBO.BOOKINGONLINE SET FPROSES = 2,HASILFU = '" & NulSID(cbCustomerMenolak.Text) & "' WHERE ID = '" & rsRecord("Id") & "'")
33                            Else
34                                Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..BOOKINGONLINE SET FPROSES = 2,HASILFU = '" & NulSID(cbCustomerMenolak.Text) & "' WHERE ID = '" & rsRecord("Id") & "'")
35                            End If
36                        End If
37                    Else
38                        cancelreason = NulSID(txtAlasan.Text)
39                        Call objO.InsertData("UPDATE BOOKINGONLINE SET FPROSES = 2,HASILFU = '" & Trim(txtAlasan.Text) & "' WHERE ID = '" & rsRecord("Id") & "'")
40                        Call objO.InsertDataHO("UPDATE BOOKINGONLINE SET FPROSES = 2,HASILFU = '" & Trim(txtAlasan.Text) & "' WHERE ID = '" & rsRecord("Id") & "'")
41                    End If
42                    If evenID <> "" And Log.DiDxP1 = 1 Then
      '                    Call objO.InsertDataHO("INSERT INTO didx..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & rsRecord("nobookingapps") & "','" & rsRecord("nobookingapps") & "','','" & Right(gDb2, 2) & "','CANCELLED','0','" & evenID & "',getdate())")
      '                    Call objO.InsertDataHO("UPDATE didx..bookingonline_didx_toSend SET booking_status = 'CANCELLED',cancellation_reason = '" & cancelreason & "' WHERE event_id = '" & evenID & "' ")
43                        Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & rsRecord("nobookingapps") & "','" & rsRecord("nobookingapps") & "','','" & Right(gDb2, 2) & "','CANCELLED','0','" & evenID & "',getdate())")
44                        Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..bookingonline_didx_toSend SET booking_status = 'CANCELLED',cancellation_reason = '" & cancelreason & "' WHERE event_id = '" & evenID & "' ")
45                    End If
                      
46                    If evenID <> "" And Log.p2Didx = 1 Then
47                        Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & rsRecord("nobookingapps") & "','" & rsRecord("nobookingapps") & "','','" & Right(gDb2, 2) & "','CANCELLED','0','" & evenID & "',getdate())")
48                        Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..p2_bookingonline_didx SET booking_status = 'CANCELLED',cancellation_reason = '" & cancelreason & "' WHERE event_id = '" & evenID & "' ")
49                    End If
50                End If
                  
51                rsRecord.MoveNext
52            Loop
              
53            MsgBox "Data Berhasil Diubah", vbInformation
54            frameBatalBooking.Visible = False
55            Call refreshGrid
56        End If
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Load()
    tglAwal = Format(isnow, "dd/mm/yyyy")
    tglAkhir = Format(isnow, "dd/mm/yyyy")
    
    If Log.RPLHistori = 0 Then
        Set rsRecord = objO.FetchDataHO("select count(*) from bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
    Else
        Set rsRecord = objO.FetchData("select count(*) from [" & dbReplikasiHO & "]..bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
    End If
    
    jumlahBookingOnline = 0
    If rsRecord.RecordCount > 0 Then
        jumlahBookingOnline = rsRecord.Fields(0).value
    End If
    tblImport.Caption = "DATA ONLINE ( " & jumlahBookingOnline & " )"
    Call hapusForm
    Call refreshGrid
    
End Sub

Private Sub hapusForm()
    nTagih = 0
    totTagih = 0
End Sub

Private Sub createRS()
    If rsRecord.State Then rsRecord.Close
    Set rsRecord = Nothing
    With rsRecord
        'If .State <> adStateClosed Then .Close
        .Fields.Append "No", adInteger
        .Fields.Append "Import", adBoolean
        .Fields.Append "Id", adInteger
        .Fields.Append "NoPolisi", adVarChar, 12
        .Fields.Append "NamaCustomer", adVarChar, 50
        .Fields.Append "JenisKendaraan", adVarChar, 50
        .Fields.Append "TanggalBooking", adVarChar, 25
        .Fields.Append "JamBooking", adVarChar, 25
        .Fields.Append "NoTel", adVarChar, 25
        .Fields.Append "Email", adVarChar, 50
        .Fields.Append "Kota", adVarChar, 50
        .Fields.Append "ProsesBooking", adVarChar, 25
        .Fields.Append "TanggalEntry", adVarChar, 25
        .Fields.Append "NoBooking", adVarChar, 25
        ' ==================================================================================
        ' 14 Feb 2017 | THANIA | update tambahan keluhan, job order & model kendaraan
        ' ==================================================================================
        .Fields.Append "keluhan", adVarChar, 200
        .Fields.Append "jenisService", adVarChar, 100
        ' ==================================================================================
        ' 23-04-2018 update kodenasmocoapp dan kodeprogramapp
        ' ----------------------------------------------------------------------------------
        .Fields.Append "nobookingapps", adVarChar, 100
        .Fields.Append "kodepromoapp", adVarChar, 100
        ' ==================================================================================
        ' 04-05-2018 update kode booking cars world
        ' ----------------------------------------------------------------------------------
        
        .Fields.Append "noBookCarsWorld", adVarChar, 100
        .Fields.Append "pic", adVarChar, 100
        .Fields.Append "sumber", adVarChar, 100
        .Fields.Append "carrier_name", adVarChar, 100
        .Fields.Append "carrier_phone_number", adVarChar, 100
        .Fields.Append "booking_source", adVarChar, 100
        .Fields.Append "status_mtoyota", adVarChar, 100
        .Fields.Append "slot_requested_datetime_start", adVarChar, 100
        .Fields.Append "service_pricing", adVarChar, 100
        .Open
    End With
    Set gridTagih.DataSource = rsRecord
    Call headColumns
End Sub

Private Sub headColumns()
    
    With gridTagih
        .Columns(0).Caption = "No"
        .Columns(0).DataField = "No"
        .Columns(0).Width = 600
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Caption = "Proses Booking"
        .Columns(1).DataField = "Import"
        .Columns(1).Width = 500
      '  .Splits(0).Columns(1).Locked = True
        
        .Columns(3).Caption = "No Polisi"
        .Columns(3).DataField = "NoPolisi"
        .Columns(3).Width = 1000
        .Splits(0).Columns(3).Locked = True
        
        .Columns(2).Caption = "ID Booking"
        .Columns(2).DataField = "ID"
        .Columns(2).Width = 1000
        .Splits(0).Columns(2).Locked = True
        
        .Columns(4).Caption = "Nama Customer"
        .Columns(4).DataField = "NamaCustomer"
        .Columns(4).Width = 2000
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).DataField = "NoTel"
        .Columns(5).Width = 1000
        .Columns(5).Caption = "No Telepon"
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Caption = "Jenis Kendaraan"
        .Columns(6).DataField = "JenisKendaraan"
        .Columns(6).Width = 2000
        .Splits(0).Columns(6).Locked = True
        
        .Columns(7).Caption = "Tanggal Booking"
        .Columns(7).DataField = "TanggalBooking"
        .Columns(7).Width = 1000
        .Splits(0).Columns(7).Locked = True
        
        .Columns(8).Caption = "Jam Booking"
        .Columns(8).DataField = "JamBooking"
        .Columns(8).Width = 1000
        .Splits(0).Columns(8).Locked = True
        
        .Columns(9).Caption = "Status"
        .Columns(9).DataField = "ProsesBooking"
        .Columns(9).Width = 2000
        .Splits(0).Columns(9).Locked = True
        
        .Columns(10).Caption = "No Booking"
        .Columns(10).DataField = "NoBooking"
        .Columns(10).Width = 1500
        .Splits(0).Columns(10).Locked = True
        
        .Columns(11).Visible = False
        .Columns(11).DataField = "ProsesBooking"
        .Columns(11).Width = 1500
        .Splits(0).Columns(11).Locked = True
        
        .Columns(12).Caption = "Tanggal Input"
        .Columns(12).DataField = "TanggalEntry"
        .Columns(12).Width = 1500
        .Splits(0).Columns(12).Locked = True
        
        .Columns(13).Caption = "Sumber Booking"
        .Columns(13).DataField = "sumber"
        .Columns(13).Width = 2500
        .Splits(0).Columns(13).Locked = True
        
        If Log.hansaRenkeiBP = 1 Then
            .Columns(14).Caption = "Pembuat"
            .Columns(14).DataField = "pic"
            .Columns(14).Width = 2500
            .Columns(14).Visible = True
            .Splits(0).Columns(14).Locked = True
            
        Else
            .Columns(14).Caption = "Sumber Booking"
            .Columns(14).DataField = "pic"
            .Columns(14).Visible = False
        End If
        
        
        .Columns(15).Caption = "Nama Pembawa"
        .Columns(15).DataField = "carrier_name"
        .Columns(15).Width = 2500
        .Splits(0).Columns(15).Locked = True
        
        .Columns(16).Caption = "NoTelp Pembawa"
        .Columns(16).DataField = "carrier_phone_number"
        .Columns(16).Width = 2500
        .Splits(0).Columns(16).Locked = True
        
        .Columns(17).Caption = "Sumber Booking"
        .Columns(17).DataField = "booking_source"
        .Columns(17).Width = 2500
        .Splits(0).Columns(17).Locked = True
        
        .Columns(18).Caption = "Status MToyota"
        .Columns(18).DataField = "status_mtoyota"
        .Columns(18).Width = 2500
        .Splits(0).Columns(18).Locked = True
        
        .Columns(19).Caption = "Tanggal Request MToyota"
        .Columns(19).DataField = "slot_requested_datetime_start"
        .Columns(19).Width = 2500
        .Splits(0).Columns(19).Locked = True
        .Columns(20).Caption = "Service Pricing"
        .Columns(20).DataField = "service_pricing"
        .Columns(20).Width = 1000
        .Splits(0).Columns(20).Locked = True
        
    End With
End Sub

Private Sub refreshGrid()
      Dim rsTagih As ADODB.Recordset
      Dim varTagih As String
      Dim kriteria As String
1     kriteria = "("

2     If Check1.value = vbChecked Then
3         kriteria = kriteria & "0,"
4     End If

5     If Check2.value = vbChecked Then
6         kriteria = kriteria & "2,3,"
7     End If

8     If Check3.value = vbChecked Then
9         kriteria = kriteria & "1,"
10    End If

11    If kriteria = "(" Then
12        Call createRS
13        Exit Sub
14    Else
15        kriteria = Mid(kriteria, 1, Len(kriteria) - 1) & ")"
16    End If
          
17        If Log.p2Didx = 1 Then
18            Set rsTagih = objO.FetchData("select nofak,fProses,replace(notel,'`','') as Notel,a.Email,Kota,convert(varchar,tanggal_booking,103) tanggalbooking,convert(varchar,jam_booking,108) jambooking," & _
                        "id,nopol as NoPolisi,nama_pelanggan as NamaCustomer,jenis_mobil as JenisKendaraan,a.tglentry, jenis_service, service_lain, keluhan, a.nobookingapps, kodepromoapp,noBookCarsWorld,carrier_name,carrier_phone_number,booking_source,b.booking_status,b.slot_requested_datetime_start,case when service_pricing_call_flag = '0' then 'False' else 'True' end service_pricing " & _
                        "from bookingonline a WITH (NOLOCK) " & _
                        "left join [" & dbReplikasiHO & "]..p2_bookingonline_didx b with (nolock) ON a.nobookingapps = b.nobookingapps and isnull(b.booking_source,'')  <> 'WALK_IN' where (tanggal_booking between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & " 23:59') and fproses in " & kriteria & " and dealer = '" & Log.KodeDealer & "' order by tanggal_booking,jam_booking")
19        Else
20            Set rsTagih = objO.FetchData("select nofak,fProses,replace(notel,'`','') as Notel,a.Email,Kota,convert(varchar,tanggal_booking,103) tanggalbooking,convert(varchar,jam_booking,108) jambooking," & _
                        "id,nopol as NoPolisi,nama_pelanggan as NamaCustomer,jenis_mobil as JenisKendaraan,a.tglentry, jenis_service, service_lain, keluhan, a.nobookingapps, kodepromoapp,noBookCarsWorld,carrier_name,carrier_phone_number,booking_source,'' as booking_status,'' as slot_requested_datetime_start,'False' as service_pricing " & _
                        "from bookingonline a WITH (NOLOCK) " & _
                        "left join [" & dbReplikasiHO & "]..bookingonline_didx_tosend b with (nolock) ON a.nobookingapps = b.nobookingapps and isnull(b.booking_source,'')  <> 'WALK_IN' where (tanggal_booking between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & " 23:59') and fproses in " & kriteria & " and dealer = '" & Log.KodeDealer & "' order by tanggal_booking,jam_booking")
21        End If
          
22        If rsTagih.RecordCount > 0 Then
23            Call createRS
              Dim hasilProsesBooking As String
24            i = 1
25            nTotal = 0
26            rsTagih.MoveFirst
27            Do While Not rsTagih.EOF
28                hasilProsesBooking = ""
29                If rsTagih("fProses") = 0 Then
30                    hasilProsesBooking = "Belum Proses"
                     ' nTotal = nTotal + 1
31                ElseIf rsTagih("fProses") = 1 Then
32                    hasilProsesBooking = "Booking"
33                ElseIf rsTagih("fProses") = 2 Or rsTagih("fProses") = 3 Then
34                    hasilProsesBooking = "Batal Booking"
35                End If
36                If rsTagih("booking_status") = "CANCELLED" Then
37                    hasilProsesBooking = "Batal Booking"
38                End If
39                With rsRecord
40                    .AddNew
41                    !No = i
42                    !Import = False
43                    !id = rsTagih("ID")
44                    !NoPolisi = Trim(rsTagih("NoPolisi"))
45                    !namacustomer = Trim(rsTagih("NamaCustomer"))
46                    !JenisKendaraan = Trim(rsTagih("JenisKendaraan"))
47                    !TanggalBooking = Trim(rsTagih("TanggalBooking"))
48                    !JamBooking = Trim(rsTagih("JamBooking"))
49                    !Notel = Trim(rsTagih("Notel"))
50                    !email = Trim(rsTagih("Email"))
51                    !Kota = Trim(rsTagih("Kota"))
52                    !ProsesBooking = hasilProsesBooking
53                    !TanggalEntry = rsTagih("tglentry")
54                    !NoBooking = Trim(rsTagih("NoFak"))
                      ' ==================================================================================
                      ' 14 Feb 2017 | THANIA | update tambahan keluhan, job order & model kendaraan
                      ' ==================================================================================
55                    !keluhan = Trim(rsTagih("Keluhan"))
56                    !jenisService = Trim(rsTagih("jenis_service")) & " " & Trim(rsTagih("Service_lain"))
                      ' ==================================================================================
                      
                      '' ---------------------------------------------------------------
                      '' 23-04-2018 penambahan no booking app dan kode promo
                      '' ---------------------------------------------------------------
57                    !kodePromoApp = NulSID(rsTagih!kodePromoApp)
58                    !noBookingApps = NulSID(rsTagih!noBookingApps)
                      ' ==================================================================================
                      ' 04-05-2018 update kode booking cars world
                      ' ----------------------------------------------------------------------------------
59                    !noBookCarsWorld = NulSID(rsTagih!noBookCarsWorld)
60                    If NulSID(rsTagih!noBookingApps) = "" Then
61                        !sumber = "Web"
62                    Else
63                        If Log.hansaRenkei = 1 Then
64                            If Len(NulSID(rsTagih!noBookCarsWorld)) > 5 Then
65                                If Left(NulSID(rsTagih!noBookCarsWorld), 5) = NulSID("sales") Then
66                                    !sumber = "Nasmoco Intra Apps"
67                                Else
68                                    !sumber = "Nasmoco Apps"
69                                End If
70                            Else
71                                !sumber = "Nasmoco Apps"
72                            End If
73                        Else
74                            !sumber = "Nasmoco Apps"
75                        End If
                          
76                        If Log.hansaRenkeiBP = 1 Then
77                            If Len(NulSID(rsTagih!noBookCarsWorld)) > 5 Then
78                                If Left(NulSID(rsTagih!noBookCarsWorld), 5) = NulSID("sales") Then
79                                    !sumber = "Nasmoco Intra Apps"
80                                    !PIC = Replace(Mid(NulSID(rsTagih!noBookCarsWorld), 7, Len(NulSID(rsTagih!noBookCarsWorld))), "_", " ")
81                                ElseIf Left(NulSID(rsTagih!noBookCarsWorld), 4) = NulSID("sabp") Then
82                                    !sumber = "NIS APC"
83                                    If InStr(NulSID(rsTagih!noBookCarsWorld), "SAGR") > 0 Then
84                                        !PIC = Left(Replace(Mid(NulSID(rsTagih!noBookCarsWorld), 6, Len(NulSID(rsTagih!noBookCarsWorld))), "_", " "), InStr(NulSID(rsTagih!noBookCarsWorld), "SAGR") - 9)
85                                    End If
86                                Else
87                                    !sumber = "Nasmoco Apps"
88                                End If
89                            Else
90                                !sumber = "Nasmoco Apps"
91                            End If
92                        Else
93                            !sumber = "Nasmoco Apps"
94                        End If
                          
95                        If Len(NulSID(rsTagih!noBookingApps)) > 5 Then
96                            If Left(NulSID(rsTagih!noBookingApps), 4) = NulSID("didx") Then
97                                !sumber = "M-Toyota DIDX"
98                            End If
99                        End If
                          
100                       !carrier_name = NulSID(rsTagih!carrier_name)
101                       !carrier_phone_number = NulSID(rsTagih!carrier_phone_number)
102                       !booking_source = NulSID(rsTagih!booking_source)
103                       !status_mtoyota = NulSID(rsTagih!booking_status)
104                       !slot_requested_datetime_start = Format(NulSID(rsTagih!slot_requested_datetime_start), "dd-mm-yyyy hh:mm")
105                       !service_pricing = NulSID(rsTagih!service_pricing)
106                   End If
107                   .Update
108               End With
109               i = i + 1
110               rsTagih.MoveNext
111           Loop
112           gridTagih.DataSource = rsRecord
113           jmlTagih.Caption = nTotal
114           gridTagih.FetchRowStyle = True
115           Call headColumns
116           rsRecord.MoveFirst
117       Else
118           Call createRS
119       End If
          
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

Private Sub gridTagih_AfterColUpdate(ByVal ColIndex As Integer)
    With gridTagih
        Select Case ColIndex
            Case 1
                If .Columns(ColIndex) Then
                    nTotal = nTotal + 1
                Else
                    nTotal = nTotal - 1
                End If
               
        End Select
        
        jmlTagih.Caption = nTotal
    End With
End Sub

Private Sub gridTagih_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsRecord.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    
    If Trim(rs.Fields(11).value) = "Batal Booking" Then
        RowStyle = "SudahDitagih"
    ElseIf Trim(rs.Fields(11).value) = "Booking" Then
        RowStyle = "Booking"
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
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

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

Private Sub tblBatal_Click()
    Unload Me
End Sub

Private Sub tblImport_Click()
    Dim f As New frmBookingImportDataOnline '.Show vbModal
    f.show vbModal
End Sub

Private Sub tblSimpan_Click()
    Dim nopol, noTelpBookingOnline, namaCustomerBookingOnline, kotaBookingOnline, tanggalBookingOnline, jamBookingOnline, modelBO, keluhanBO, permintaanBO, noBookingApps, kodePromoApp, noBookCarsWorld, nikSales As String
    Dim IDBookingOnline As Double
    
    'If Not rsRecord.EOF Then
    If nTotal > 0 Then
        If nTotal > 1 Then
            MsgBox "Proses pembuatan booking hanya diperbolehkan satu per satu", vbInformation
            Exit Sub
        End If
        Dim rsCloneBooking As New ADODB.Recordset
        Set rsCloneBooking = rsRecord
        
        rsCloneBooking.MoveFirst
        nopol = ""
        Do While Not rsCloneBooking.EOF
            If rsCloneBooking("import") And rsCloneBooking("ProsesBooking") = "Belum Proses" Then
                nopol = rsCloneBooking("NoPolisi")
                IDBookingOnline = CDbl(rsCloneBooking("ID"))
                noTelpBookingOnline = Trim(rsCloneBooking("Notel"))
                namaCustomerBookingOnline = Trim(rsCloneBooking("NamaCustomer"))
                kotaBookingOnline = Trim(rsCloneBooking("Kota"))
                tanggalBookingOnline = Trim(rsCloneBooking("TanggalBooking"))
                jamBookingOnline = Trim(rsCloneBooking("JamBooking"))
                modelBO = Trim(rsCloneBooking("jeniskendaraan"))
                keluhanBO = Trim(rsCloneBooking("keluhan"))
                permintaanBO = Trim(rsCloneBooking("jenisservice"))
                
                '' ---------------------------------------------------
                '' 23-04-2018 penambahan nobookingapp dan kode promo
                '' ---------------------------------------------------
                noBookingApps = NulSID(rsCloneBooking!noBookingApps)
                kodePromoApp = NulSID(rsCloneBooking!kodePromoApp)
                
                ' ==================================================================================
                ' 04-05-2018 update kode booking cars world
                ' ----------------------------------------------------------------------------------
                noBookCarsWorld = NulSID(rsCloneBooking!noBookCarsWorld)
                nikSales = ""
                If Log.hansaRenkei = 1 Then
                    If Len(noBookCarsWorld) > 5 Then
                        If Left(NulSID(noBookCarsWorld), 5) = NulSID("sales") Then
                            noBookCarsWorld = Replace(noBookCarsWorld, " ", "")
                            nikSales = Mid(noBookCarsWorld, 7, Len(noBookCarsWorld))
                        End If
                    End If
                End If
            End If
            rsCloneBooking.MoveNext
        Loop
     '   rsTempImport.MoveFirst
        If nopol <> "" Then
            If MsgBox("Apakah yakkin membuat Booking dari No Polisi = " & nopol & " ?", vbYesNo) = vbYes Then
                If objO.FetchData_Count("SELECT ID FROM BOOKINGONLINE WHERE ID = '" & IDBookingOnline & "' AND FPROSES = 0") = 0 Then
                    MsgBox "Unit dengan No Polisi = " & nopol & " telah dilakukan booking / batal booking", vbInformation
                Else
                    If gFiskal = Format(Month(isnow), "00") Then
                            '-----------------------------------
                            '.:: SET AS CHILD ::.
                            '-----------------------------------
                            TOChild Layar.hWnd, frmBooking.hWnd
                            '-----------------------------------
                            '-----------------------------------
                            isBookByMrs = False
                            frmBooking.passingISBookingOnline = True
                            frmBooking.passingNoPolBookingOnline = nopol
    
                            frmBooking.passingIDBookingOnline = IDBookingOnline
    
                            frmBooking.passingNoTelpBookingOnline = noTelpBookingOnline
                            frmBooking.passingNamaCustomerBookingOnline = namaCustomerBookingOnline
                            frmBooking.passingKotaBookingOnline = kotaBookingOnline
                            frmBooking.passingTanggalBookingOnline = tanggalBookingOnline
                            frmBooking.passingJamBookingOnline = jamBookingOnline
                            frmBooking.passingmdlKendaraanOnline = modelBO
                            frmBooking.passingKeluhanOnline = keluhanBO
                            frmBooking.passingPermintaanOnline = permintaanBO
                            frmBooking.Height = frmBooking.asxPowerBanner1.Height
                            
                            frmBooking.passingNoBookingAppOnline = noBookingApps
                            frmBooking.passingKodePromosiAppOnline = kodePromoApp
                            If Log.hansaRenkei = 1 Then
                                frmBooking.passingNisSALES = nikSales
                            End If
                            frmBooking.show
                            refreshGrid
                    Else
                        MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
                    End If
                End If
            End If
        Else
            MsgBox "Tidak ada No Polisi yang dapat dibuat Booking", vbInformation
            refreshGrid
        End If
        
    End If
    
End Sub
