VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmSIDSearchCust 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   14550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameDijual 
      Height          =   3855
      Left            =   2040
      TabIndex        =   68
      Top             =   3360
      Visible         =   0   'False
      Width           =   8175
      Begin ucXPButton.XPButton cmdAktif 
         Height          =   495
         Left            =   5640
         TabIndex        =   95
         Top             =   3240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "XPButton1"
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
         MICON           =   "frmSIDSearchCust.frx":0000
         PICN            =   "frmSIDSearchCust.frx":001C
         PICH            =   "frmSIDSearchCust.frx":6E96
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdTutup 
         Height          =   495
         Left            =   6840
         TabIndex        =   93
         Top             =   3240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "XPButton1"
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
         MICON           =   "frmSIDSearchCust.frx":E210
         PICN            =   "frmSIDSearchCust.frx":E22C
         PICH            =   "frmSIDSearchCust.frx":11554
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapus 
         Height          =   495
         Left            =   5640
         TabIndex        =   94
         Top             =   3240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "XPButton1"
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
         MICON           =   "frmSIDSearchCust.frx":14946
         PICN            =   "frmSIDSearchCust.frx":14962
         PICH            =   "frmSIDSearchCust.frx":1B4D5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label labelAlamat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1920
         TabIndex        =   92
         Top             =   2640
         Width           =   6015
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   120
         TabIndex        =   91
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1725
         TabIndex        =   90
         Top             =   2640
         Width           =   165
      End
      Begin VB.Label labelNamaCustomer 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   89
         Top             =   2400
         Width           =   6015
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   120
         TabIndex        =   88
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1725
         TabIndex        =   87
         Top             =   2400
         Width           =   165
      End
      Begin VB.Label labelIDCustomer 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   86
         Top             =   2160
         Width           =   6015
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   120
         TabIndex        =   85
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1725
         TabIndex        =   84
         Top             =   2160
         Width           =   165
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "akan DIHAPUSKAN dari data Customer ini :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   120
         TabIndex        =   83
         Top             =   1920
         Width           =   7935
      End
      Begin VB.Label labelWarna 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   82
         Top             =   1560
         Width           =   6015
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   120
         TabIndex        =   81
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1725
         TabIndex        =   80
         Top             =   1560
         Width           =   165
      End
      Begin VB.Label labelModel 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   79
         Top             =   1320
         Width           =   6015
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Model "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   120
         TabIndex        =   78
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1725
         TabIndex        =   77
         Top             =   1320
         Width           =   165
      End
      Begin VB.Label labelNoPolisi 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   76
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   75
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1725
         TabIndex        =   74
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label labelNoRangka 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   73
         Top             =   840
         Width           =   6015
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Apakah Unit dengan data berikut :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   72
         Top             =   600
         Width           =   7935
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   71
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1725
         TabIndex        =   70
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblKeteranganDijual 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Konfirmasi Unit Dijual / Pindah Tangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   69
         Top             =   120
         Width           =   8175
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   8175
      End
   End
   Begin VB.Frame frameDuplikasi 
      Height          =   8655
      Left            =   0
      TabIndex        =   56
      Top             =   360
      Width           =   12495
      Begin VB.Frame frameUploadKTPMerge 
         Caption         =   "Proses Upload KTP Merge"
         Height          =   8535
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   12495
         Begin TrueOleDBGrid70.TDBGrid gridCustomerFix 
            Height          =   1095
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   1931
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
            Columns.Count   =   16
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   953
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   15790320
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=16"
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
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(45)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(49)=   "Column(12).Width=2725"
            Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=2646"
            Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(53)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(57)=   "Column(14).Width=2725"
            Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=2646"
            Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(61)=   "Column(15).Width=2725"
            Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=2646"
            Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "DATA CUSTOMER MERGE"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   15790320
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
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
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
            _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
            _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
            _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
            _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
            _StyleDefs(100) =   "Named:id=33:Normal"
            _StyleDefs(101) =   ":id=33,.parent=0"
            _StyleDefs(102) =   "Named:id=34:Heading"
            _StyleDefs(103) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(104) =   ":id=34,.wraptext=-1"
            _StyleDefs(105) =   "Named:id=35:Footing"
            _StyleDefs(106) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(107) =   "Named:id=36:Selected"
            _StyleDefs(108) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(109) =   "Named:id=37:Caption"
            _StyleDefs(110) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(111) =   "Named:id=38:HighlightRow"
            _StyleDefs(112) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(113) =   "Named:id=39:EvenRow"
            _StyleDefs(114) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(115) =   "Named:id=40:OddRow"
            _StyleDefs(116) =   ":id=40,.parent=33"
            _StyleDefs(117) =   "Named:id=41:RecordSelector"
            _StyleDefs(118) =   ":id=41,.parent=34"
            _StyleDefs(119) =   "Named:id=42:FilterBar"
            _StyleDefs(120) =   ":id=42,.parent=33"
         End
         Begin VB.TextBox txtUploadKTP 
            BackColor       =   &H00C0E0FF&
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   1560
            Width           =   7455
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   6600
            Top             =   4680
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ucXPButton.XPButton cmdUploadKTP 
            Height          =   495
            Left            =   120
            TabIndex        =   62
            Top             =   1470
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   ""
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
            MICON           =   "frmSIDSearchCust.frx":2226F
            PICN            =   "frmSIDSearchCust.frx":2228B
            PICH            =   "frmSIDSearchCust.frx":259C9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdSimpanMerge 
            Height          =   495
            Left            =   9360
            TabIndex        =   63
            Top             =   1470
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&a"
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
            MICON           =   "frmSIDSearchCust.frx":292AD
            PICN            =   "frmSIDSearchCust.frx":292C9
            PICH            =   "frmSIDSearchCust.frx":2FCE6
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
            Height          =   3735
            Left            =   120
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   6015
         End
      End
      Begin TrueOleDBGrid70.TDBGrid gridDuplikat 
         Height          =   7575
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   13361
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ID Merge"
         Columns(0).DataField=   "idmerge"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   4
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Pilih Merge"
         Columns(1).DataField=   "pilihMerge"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ID Cust"
         Columns(2).DataField=   "idcust"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   "pilihNama"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nama Customer"
         Columns(4).DataField=   "Nama"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   4
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   "pilihTglLahir"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Tgl Lahir"
         Columns(6).DataField=   "tglLahir"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "pilihAgama"
         Columns(7).DataField=   "pilihAgama"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Agama"
         Columns(8).DataField=   "Agama"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   4
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "pilihPekerjaan"
         Columns(9).DataField=   "pilihPekerjaan"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Pekerjaan"
         Columns(10).DataField=   "pekerjaan"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   4
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "pilihHobi"
         Columns(11).DataField=   "pilihHobi"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Hobi"
         Columns(12).DataField=   "Hobi"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   4
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "pilihOlahraga"
         Columns(13).DataField=   "pilihOlahraga"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Olahraga"
         Columns(14).DataField=   "Olahraga"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   4
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "pilihPropinsi"
         Columns(15).DataField=   "pilihPropinsi"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Propinsi"
         Columns(16).DataField=   "Propinsi"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   4
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "pilihKota"
         Columns(17).DataField=   "pilihKota"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "Kota"
         Columns(18).DataField=   "Kota"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   4
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "pilihKecamatan"
         Columns(19).DataField=   "pilihKecamatan"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "Kecamatan"
         Columns(20).DataField=   "Kecamatan"
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   4
         Columns(21)._MaxComboItems=   5
         Columns(21).Caption=   "pilihKelurahan"
         Columns(21).DataField=   "pilihKelurahan"
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(22)._VlistStyle=   0
         Columns(22)._MaxComboItems=   5
         Columns(22).Caption=   "Kelurahan"
         Columns(22).DataField=   "Kelurahan"
         Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(23)._VlistStyle=   4
         Columns(23)._MaxComboItems=   5
         Columns(23).Caption=   "pilihHP"
         Columns(23).DataField=   "pilihHP"
         Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(24)._VlistStyle=   0
         Columns(24)._MaxComboItems=   5
         Columns(24).Caption=   "HP"
         Columns(24).DataField=   "HP"
         Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(25)._VlistStyle=   4
         Columns(25)._MaxComboItems=   5
         Columns(25).Caption=   "pilihKTP"
         Columns(25).DataField=   "pilihKTP"
         Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(26)._VlistStyle=   0
         Columns(26)._MaxComboItems=   5
         Columns(26).Caption=   "No KTP"
         Columns(26).DataField=   "NoKTP"
         Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(27)._VlistStyle=   4
         Columns(27)._MaxComboItems=   5
         Columns(27).Caption=   "pilihNPWP"
         Columns(27).DataField=   "pilihNPWP"
         Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(28)._VlistStyle=   0
         Columns(28)._MaxComboItems=   5
         Columns(28).Caption=   "NPWP"
         Columns(28).DataField=   "noNPWP"
         Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(29)._VlistStyle=   4
         Columns(29)._MaxComboItems=   5
         Columns(29).DataField=   ""
         Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   30
         Splits(0)._UserFlags=   0
         Splits(0).SizeMode=   2
         Splits(0).Size  =   2
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=30"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1323"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1588"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1508"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8708"
         Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=8708"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=529"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=450"
         Splits(0)._ColumnProps(31)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(33)=   "Column(6).Width=2117"
         Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=2037"
         Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=8708"
         Splits(0)._ColumnProps(37)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(39)=   "Column(7).Width=529"
         Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=450"
         Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(43)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(45)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(8)._ColStyle=8708"
         Splits(0)._ColumnProps(49)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(51)=   "Column(9).Width=529"
         Splits(0)._ColumnProps(52)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(9)._WidthInPix=450"
         Splits(0)._ColumnProps(54)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(55)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(56)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(59)=   "Column(10)._ColStyle=8708"
         Splits(0)._ColumnProps(60)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(61)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(62)=   "Column(11).Width=529"
         Splits(0)._ColumnProps(63)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(64)=   "Column(11)._WidthInPix=450"
         Splits(0)._ColumnProps(65)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(66)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(67)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(68)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(70)=   "Column(12)._ColStyle=8708"
         Splits(0)._ColumnProps(71)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(72)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(73)=   "Column(13).Width=529"
         Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=450"
         Splits(0)._ColumnProps(76)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(77)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(78)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(79)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(80)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(81)=   "Column(14)._ColStyle=8708"
         Splits(0)._ColumnProps(82)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(83)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(84)=   "Column(15).Width=529"
         Splits(0)._ColumnProps(85)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(86)=   "Column(15)._WidthInPix=450"
         Splits(0)._ColumnProps(87)=   "Column(15).Visible=0"
         Splits(0)._ColumnProps(88)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(89)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(90)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(91)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(92)=   "Column(16)._ColStyle=8708"
         Splits(0)._ColumnProps(93)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(94)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(95)=   "Column(17).Width=529"
         Splits(0)._ColumnProps(96)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(97)=   "Column(17)._WidthInPix=450"
         Splits(0)._ColumnProps(98)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(99)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(100)=   "Column(18).Width=2831"
         Splits(0)._ColumnProps(101)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(102)=   "Column(18)._WidthInPix=2752"
         Splits(0)._ColumnProps(103)=   "Column(18)._ColStyle=8708"
         Splits(0)._ColumnProps(104)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(105)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(106)=   "Column(19).Width=529"
         Splits(0)._ColumnProps(107)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(108)=   "Column(19)._WidthInPix=450"
         Splits(0)._ColumnProps(109)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(110)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(111)=   "Column(20).Width=2725"
         Splits(0)._ColumnProps(112)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(113)=   "Column(20)._WidthInPix=2646"
         Splits(0)._ColumnProps(114)=   "Column(20)._ColStyle=8708"
         Splits(0)._ColumnProps(115)=   "Column(20).Visible=0"
         Splits(0)._ColumnProps(116)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(117)=   "Column(21).Width=529"
         Splits(0)._ColumnProps(118)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(119)=   "Column(21)._WidthInPix=450"
         Splits(0)._ColumnProps(120)=   "Column(21).Visible=0"
         Splits(0)._ColumnProps(121)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(122)=   "Column(22).Width=2725"
         Splits(0)._ColumnProps(123)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(124)=   "Column(22)._WidthInPix=2646"
         Splits(0)._ColumnProps(125)=   "Column(22)._ColStyle=8708"
         Splits(0)._ColumnProps(126)=   "Column(22).Visible=0"
         Splits(0)._ColumnProps(127)=   "Column(22).Order=23"
         Splits(0)._ColumnProps(128)=   "Column(23).Width=529"
         Splits(0)._ColumnProps(129)=   "Column(23).DividerColor=0"
         Splits(0)._ColumnProps(130)=   "Column(23)._WidthInPix=450"
         Splits(0)._ColumnProps(131)=   "Column(23).Visible=0"
         Splits(0)._ColumnProps(132)=   "Column(23).Order=24"
         Splits(0)._ColumnProps(133)=   "Column(24).Width=2725"
         Splits(0)._ColumnProps(134)=   "Column(24).DividerColor=0"
         Splits(0)._ColumnProps(135)=   "Column(24)._WidthInPix=2646"
         Splits(0)._ColumnProps(136)=   "Column(24)._ColStyle=8708"
         Splits(0)._ColumnProps(137)=   "Column(24).Visible=0"
         Splits(0)._ColumnProps(138)=   "Column(24).Order=25"
         Splits(0)._ColumnProps(139)=   "Column(25).Width=529"
         Splits(0)._ColumnProps(140)=   "Column(25).DividerColor=0"
         Splits(0)._ColumnProps(141)=   "Column(25)._WidthInPix=450"
         Splits(0)._ColumnProps(142)=   "Column(25).Visible=0"
         Splits(0)._ColumnProps(143)=   "Column(25).Order=26"
         Splits(0)._ColumnProps(144)=   "Column(26).Width=2725"
         Splits(0)._ColumnProps(145)=   "Column(26).DividerColor=0"
         Splits(0)._ColumnProps(146)=   "Column(26)._WidthInPix=2646"
         Splits(0)._ColumnProps(147)=   "Column(26)._ColStyle=8708"
         Splits(0)._ColumnProps(148)=   "Column(26).Visible=0"
         Splits(0)._ColumnProps(149)=   "Column(26).Order=27"
         Splits(0)._ColumnProps(150)=   "Column(27).Width=529"
         Splits(0)._ColumnProps(151)=   "Column(27).DividerColor=0"
         Splits(0)._ColumnProps(152)=   "Column(27)._WidthInPix=450"
         Splits(0)._ColumnProps(153)=   "Column(27).Visible=0"
         Splits(0)._ColumnProps(154)=   "Column(27).Order=28"
         Splits(0)._ColumnProps(155)=   "Column(28).Width=2725"
         Splits(0)._ColumnProps(156)=   "Column(28).DividerColor=0"
         Splits(0)._ColumnProps(157)=   "Column(28)._WidthInPix=2646"
         Splits(0)._ColumnProps(158)=   "Column(28)._ColStyle=8708"
         Splits(0)._ColumnProps(159)=   "Column(28).Visible=0"
         Splits(0)._ColumnProps(160)=   "Column(28).Order=29"
         Splits(0)._ColumnProps(161)=   "Column(29).Width=529"
         Splits(0)._ColumnProps(162)=   "Column(29).DividerColor=0"
         Splits(0)._ColumnProps(163)=   "Column(29)._WidthInPix=450"
         Splits(0)._ColumnProps(164)=   "Column(29).Visible=0"
         Splits(0)._ColumnProps(165)=   "Column(29).Order=30"
         Splits(1)._UserFlags=   0
         Splits(1).RecordSelectorWidth=   688
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1).DividerColor=   14215660
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=30"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=1402"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1323"
         Splits(1)._ColumnProps(4)=   "Column(0)._ColStyle=1"
         Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(7)=   "Column(1).Width=1588"
         Splits(1)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(9)=   "Column(1)._WidthInPix=1508"
         Splits(1)._ColumnProps(10)=   "Column(1)._ColStyle=1"
         Splits(1)._ColumnProps(11)=   "Column(1).Visible=0"
         Splits(1)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(13)=   "Column(2).Width=2725"
         Splits(1)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
         Splits(1)._ColumnProps(16)=   "Column(2)._ColStyle=8708"
         Splits(1)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(18)=   "Column(3).Width=529"
         Splits(1)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(20)=   "Column(3)._WidthInPix=450"
         Splits(1)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(22)=   "Column(4).Width=2725"
         Splits(1)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
         Splits(1)._ColumnProps(25)=   "Column(4)._ColStyle=8708"
         Splits(1)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(27)=   "Column(5).Width=529"
         Splits(1)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(29)=   "Column(5)._WidthInPix=450"
         Splits(1)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(31)=   "Column(6).Width=2117"
         Splits(1)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(33)=   "Column(6)._WidthInPix=2037"
         Splits(1)._ColumnProps(34)=   "Column(6)._ColStyle=8708"
         Splits(1)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(36)=   "Column(7).Width=529"
         Splits(1)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(38)=   "Column(7)._WidthInPix=450"
         Splits(1)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(1)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(1)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(1)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(1)._ColumnProps(44)=   "Column(8)._ColStyle=8708"
         Splits(1)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(46)=   "Column(9).Width=529"
         Splits(1)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(48)=   "Column(9)._WidthInPix=450"
         Splits(1)._ColumnProps(49)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(50)=   "Column(10).Width=2725"
         Splits(1)._ColumnProps(51)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(52)=   "Column(10)._WidthInPix=2646"
         Splits(1)._ColumnProps(53)=   "Column(10)._ColStyle=8708"
         Splits(1)._ColumnProps(54)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(55)=   "Column(11).Width=529"
         Splits(1)._ColumnProps(56)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(57)=   "Column(11)._WidthInPix=450"
         Splits(1)._ColumnProps(58)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(59)=   "Column(12).Width=2725"
         Splits(1)._ColumnProps(60)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(61)=   "Column(12)._WidthInPix=2646"
         Splits(1)._ColumnProps(62)=   "Column(12)._ColStyle=8708"
         Splits(1)._ColumnProps(63)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(64)=   "Column(13).Width=529"
         Splits(1)._ColumnProps(65)=   "Column(13).DividerColor=0"
         Splits(1)._ColumnProps(66)=   "Column(13)._WidthInPix=450"
         Splits(1)._ColumnProps(67)=   "Column(13).Order=14"
         Splits(1)._ColumnProps(68)=   "Column(14).Width=2725"
         Splits(1)._ColumnProps(69)=   "Column(14).DividerColor=0"
         Splits(1)._ColumnProps(70)=   "Column(14)._WidthInPix=2646"
         Splits(1)._ColumnProps(71)=   "Column(14)._ColStyle=8708"
         Splits(1)._ColumnProps(72)=   "Column(14).Order=15"
         Splits(1)._ColumnProps(73)=   "Column(15).Width=529"
         Splits(1)._ColumnProps(74)=   "Column(15).DividerColor=0"
         Splits(1)._ColumnProps(75)=   "Column(15)._WidthInPix=450"
         Splits(1)._ColumnProps(76)=   "Column(15).Order=16"
         Splits(1)._ColumnProps(77)=   "Column(16).Width=2725"
         Splits(1)._ColumnProps(78)=   "Column(16).DividerColor=0"
         Splits(1)._ColumnProps(79)=   "Column(16)._WidthInPix=2646"
         Splits(1)._ColumnProps(80)=   "Column(16)._ColStyle=8708"
         Splits(1)._ColumnProps(81)=   "Column(16).Order=17"
         Splits(1)._ColumnProps(82)=   "Column(17).Width=529"
         Splits(1)._ColumnProps(83)=   "Column(17).DividerColor=0"
         Splits(1)._ColumnProps(84)=   "Column(17)._WidthInPix=450"
         Splits(1)._ColumnProps(85)=   "Column(17).Order=18"
         Splits(1)._ColumnProps(86)=   "Column(18).Width=2831"
         Splits(1)._ColumnProps(87)=   "Column(18).DividerColor=0"
         Splits(1)._ColumnProps(88)=   "Column(18)._WidthInPix=2752"
         Splits(1)._ColumnProps(89)=   "Column(18)._ColStyle=8708"
         Splits(1)._ColumnProps(90)=   "Column(18).Order=19"
         Splits(1)._ColumnProps(91)=   "Column(19).Width=529"
         Splits(1)._ColumnProps(92)=   "Column(19).DividerColor=0"
         Splits(1)._ColumnProps(93)=   "Column(19)._WidthInPix=450"
         Splits(1)._ColumnProps(94)=   "Column(19).Order=20"
         Splits(1)._ColumnProps(95)=   "Column(20).Width=2725"
         Splits(1)._ColumnProps(96)=   "Column(20).DividerColor=0"
         Splits(1)._ColumnProps(97)=   "Column(20)._WidthInPix=2646"
         Splits(1)._ColumnProps(98)=   "Column(20)._ColStyle=8708"
         Splits(1)._ColumnProps(99)=   "Column(20).Order=21"
         Splits(1)._ColumnProps(100)=   "Column(21).Width=529"
         Splits(1)._ColumnProps(101)=   "Column(21).DividerColor=0"
         Splits(1)._ColumnProps(102)=   "Column(21)._WidthInPix=450"
         Splits(1)._ColumnProps(103)=   "Column(21).Order=22"
         Splits(1)._ColumnProps(104)=   "Column(22).Width=2725"
         Splits(1)._ColumnProps(105)=   "Column(22).DividerColor=0"
         Splits(1)._ColumnProps(106)=   "Column(22)._WidthInPix=2646"
         Splits(1)._ColumnProps(107)=   "Column(22)._ColStyle=8708"
         Splits(1)._ColumnProps(108)=   "Column(22).Order=23"
         Splits(1)._ColumnProps(109)=   "Column(23).Width=529"
         Splits(1)._ColumnProps(110)=   "Column(23).DividerColor=0"
         Splits(1)._ColumnProps(111)=   "Column(23)._WidthInPix=450"
         Splits(1)._ColumnProps(112)=   "Column(23).Order=24"
         Splits(1)._ColumnProps(113)=   "Column(24).Width=2725"
         Splits(1)._ColumnProps(114)=   "Column(24).DividerColor=0"
         Splits(1)._ColumnProps(115)=   "Column(24)._WidthInPix=2646"
         Splits(1)._ColumnProps(116)=   "Column(24)._ColStyle=8708"
         Splits(1)._ColumnProps(117)=   "Column(24).Order=25"
         Splits(1)._ColumnProps(118)=   "Column(25).Width=529"
         Splits(1)._ColumnProps(119)=   "Column(25).DividerColor=0"
         Splits(1)._ColumnProps(120)=   "Column(25)._WidthInPix=450"
         Splits(1)._ColumnProps(121)=   "Column(25).Order=26"
         Splits(1)._ColumnProps(122)=   "Column(26).Width=2725"
         Splits(1)._ColumnProps(123)=   "Column(26).DividerColor=0"
         Splits(1)._ColumnProps(124)=   "Column(26)._WidthInPix=2646"
         Splits(1)._ColumnProps(125)=   "Column(26)._ColStyle=8708"
         Splits(1)._ColumnProps(126)=   "Column(26).Order=27"
         Splits(1)._ColumnProps(127)=   "Column(27).Width=529"
         Splits(1)._ColumnProps(128)=   "Column(27).DividerColor=0"
         Splits(1)._ColumnProps(129)=   "Column(27)._WidthInPix=450"
         Splits(1)._ColumnProps(130)=   "Column(27).Order=28"
         Splits(1)._ColumnProps(131)=   "Column(28).Width=2725"
         Splits(1)._ColumnProps(132)=   "Column(28).DividerColor=0"
         Splits(1)._ColumnProps(133)=   "Column(28)._WidthInPix=2646"
         Splits(1)._ColumnProps(134)=   "Column(28)._ColStyle=8708"
         Splits(1)._ColumnProps(135)=   "Column(28).Order=29"
         Splits(1)._ColumnProps(136)=   "Column(29).Width=529"
         Splits(1)._ColumnProps(137)=   "Column(29).DividerColor=0"
         Splits(1)._ColumnProps(138)=   "Column(29)._WidthInPix=450"
         Splits(1)._ColumnProps(139)=   "Column(29).Visible=0"
         Splits(1)._ColumnProps(140)=   "Column(29).Order=30"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "DATA POTENSIAL DUPLIKASI"
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
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=94,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=91,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=90,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=87,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=88,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=89,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=86,.parent=13,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=14,.alignment=2"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=282,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=279,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=280,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=281,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=82,.parent=13,.locked=-1"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=79,.parent=14,.alignment=2"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=80,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=81,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=78,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=74,.parent=13,.locked=-1"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14,.alignment=2"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14,.alignment=2"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14,.alignment=2"
         _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=62,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14,.alignment=2"
         _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=54,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=51,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=52,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=53,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=48,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=49,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=46,.parent=13"
         _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=43,.parent=14"
         _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=44,.parent=15"
         _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=45,.parent=17"
         _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=30,.parent=15"
         _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=31,.parent=17"
         _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=28,.parent=13"
         _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=25,.parent=14"
         _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=26,.parent=15"
         _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=27,.parent=17"
         _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=98,.parent=13,.locked=-1"
         _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14,.alignment=2"
         _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
         _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
         _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=102,.parent=13"
         _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=99,.parent=14"
         _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=100,.parent=15"
         _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=101,.parent=17"
         _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=106,.parent=13,.locked=-1"
         _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=103,.parent=14,.alignment=2"
         _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=104,.parent=15"
         _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=105,.parent=17"
         _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=110,.parent=13"
         _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=107,.parent=14"
         _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=108,.parent=15"
         _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=109,.parent=17"
         _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=114,.parent=13,.locked=-1"
         _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=111,.parent=14,.alignment=2"
         _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=112,.parent=15"
         _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=113,.parent=17"
         _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=118,.parent=13"
         _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=115,.parent=14"
         _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=116,.parent=15"
         _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=117,.parent=17"
         _StyleDefs(126) =   "Splits(0).Columns(22).Style:id=122,.parent=13,.locked=-1"
         _StyleDefs(127) =   "Splits(0).Columns(22).HeadingStyle:id=119,.parent=14,.alignment=2"
         _StyleDefs(128) =   "Splits(0).Columns(22).FooterStyle:id=120,.parent=15"
         _StyleDefs(129) =   "Splits(0).Columns(22).EditorStyle:id=121,.parent=17"
         _StyleDefs(130) =   "Splits(0).Columns(23).Style:id=126,.parent=13"
         _StyleDefs(131) =   "Splits(0).Columns(23).HeadingStyle:id=123,.parent=14"
         _StyleDefs(132) =   "Splits(0).Columns(23).FooterStyle:id=124,.parent=15"
         _StyleDefs(133) =   "Splits(0).Columns(23).EditorStyle:id=125,.parent=17"
         _StyleDefs(134) =   "Splits(0).Columns(24).Style:id=130,.parent=13,.locked=-1"
         _StyleDefs(135) =   "Splits(0).Columns(24).HeadingStyle:id=127,.parent=14,.alignment=2"
         _StyleDefs(136) =   "Splits(0).Columns(24).FooterStyle:id=128,.parent=15"
         _StyleDefs(137) =   "Splits(0).Columns(24).EditorStyle:id=129,.parent=17"
         _StyleDefs(138) =   "Splits(0).Columns(25).Style:id=134,.parent=13"
         _StyleDefs(139) =   "Splits(0).Columns(25).HeadingStyle:id=131,.parent=14"
         _StyleDefs(140) =   "Splits(0).Columns(25).FooterStyle:id=132,.parent=15"
         _StyleDefs(141) =   "Splits(0).Columns(25).EditorStyle:id=133,.parent=17"
         _StyleDefs(142) =   "Splits(0).Columns(26).Style:id=138,.parent=13,.locked=-1"
         _StyleDefs(143) =   "Splits(0).Columns(26).HeadingStyle:id=135,.parent=14,.alignment=2"
         _StyleDefs(144) =   "Splits(0).Columns(26).FooterStyle:id=136,.parent=15"
         _StyleDefs(145) =   "Splits(0).Columns(26).EditorStyle:id=137,.parent=17"
         _StyleDefs(146) =   "Splits(0).Columns(27).Style:id=142,.parent=13"
         _StyleDefs(147) =   "Splits(0).Columns(27).HeadingStyle:id=139,.parent=14"
         _StyleDefs(148) =   "Splits(0).Columns(27).FooterStyle:id=140,.parent=15"
         _StyleDefs(149) =   "Splits(0).Columns(27).EditorStyle:id=141,.parent=17"
         _StyleDefs(150) =   "Splits(0).Columns(28).Style:id=146,.parent=13,.locked=-1"
         _StyleDefs(151) =   "Splits(0).Columns(28).HeadingStyle:id=143,.parent=14,.alignment=2"
         _StyleDefs(152) =   "Splits(0).Columns(28).FooterStyle:id=144,.parent=15"
         _StyleDefs(153) =   "Splits(0).Columns(28).EditorStyle:id=145,.parent=17"
         _StyleDefs(154) =   "Splits(0).Columns(29).Style:id=150,.parent=13"
         _StyleDefs(155) =   "Splits(0).Columns(29).HeadingStyle:id=147,.parent=14"
         _StyleDefs(156) =   "Splits(0).Columns(29).FooterStyle:id=148,.parent=15"
         _StyleDefs(157) =   "Splits(0).Columns(29).EditorStyle:id=149,.parent=17"
         _StyleDefs(158) =   "Splits(1).Style:id=151,.parent=1"
         _StyleDefs(159) =   "Splits(1).CaptionStyle:id=160,.parent=4"
         _StyleDefs(160) =   "Splits(1).HeadingStyle:id=152,.parent=2"
         _StyleDefs(161) =   "Splits(1).FooterStyle:id=153,.parent=3"
         _StyleDefs(162) =   "Splits(1).InactiveStyle:id=154,.parent=5"
         _StyleDefs(163) =   "Splits(1).SelectedStyle:id=156,.parent=6"
         _StyleDefs(164) =   "Splits(1).EditorStyle:id=155,.parent=7"
         _StyleDefs(165) =   "Splits(1).HighlightRowStyle:id=157,.parent=8"
         _StyleDefs(166) =   "Splits(1).EvenRowStyle:id=158,.parent=9"
         _StyleDefs(167) =   "Splits(1).OddRowStyle:id=159,.parent=10"
         _StyleDefs(168) =   "Splits(1).RecordSelectorStyle:id=161,.parent=11"
         _StyleDefs(169) =   "Splits(1).FilterBarStyle:id=162,.parent=12"
         _StyleDefs(170) =   "Splits(1).Columns(0).Style:id=166,.parent=151,.alignment=2"
         _StyleDefs(171) =   "Splits(1).Columns(0).HeadingStyle:id=163,.parent=152"
         _StyleDefs(172) =   "Splits(1).Columns(0).FooterStyle:id=164,.parent=153"
         _StyleDefs(173) =   "Splits(1).Columns(0).EditorStyle:id=165,.parent=155"
         _StyleDefs(174) =   "Splits(1).Columns(1).Style:id=170,.parent=151,.alignment=2"
         _StyleDefs(175) =   "Splits(1).Columns(1).HeadingStyle:id=167,.parent=152"
         _StyleDefs(176) =   "Splits(1).Columns(1).FooterStyle:id=168,.parent=153"
         _StyleDefs(177) =   "Splits(1).Columns(1).EditorStyle:id=169,.parent=155"
         _StyleDefs(178) =   "Splits(1).Columns(2).Style:id=174,.parent=151,.locked=-1"
         _StyleDefs(179) =   "Splits(1).Columns(2).HeadingStyle:id=171,.parent=152,.alignment=2"
         _StyleDefs(180) =   "Splits(1).Columns(2).FooterStyle:id=172,.parent=153"
         _StyleDefs(181) =   "Splits(1).Columns(2).EditorStyle:id=173,.parent=155"
         _StyleDefs(182) =   "Splits(1).Columns(3).Style:id=286,.parent=151"
         _StyleDefs(183) =   "Splits(1).Columns(3).HeadingStyle:id=283,.parent=152"
         _StyleDefs(184) =   "Splits(1).Columns(3).FooterStyle:id=284,.parent=153"
         _StyleDefs(185) =   "Splits(1).Columns(3).EditorStyle:id=285,.parent=155"
         _StyleDefs(186) =   "Splits(1).Columns(4).Style:id=178,.parent=151,.locked=-1"
         _StyleDefs(187) =   "Splits(1).Columns(4).HeadingStyle:id=175,.parent=152,.alignment=2"
         _StyleDefs(188) =   "Splits(1).Columns(4).FooterStyle:id=176,.parent=153"
         _StyleDefs(189) =   "Splits(1).Columns(4).EditorStyle:id=177,.parent=155"
         _StyleDefs(190) =   "Splits(1).Columns(5).Style:id=182,.parent=151"
         _StyleDefs(191) =   "Splits(1).Columns(5).HeadingStyle:id=179,.parent=152"
         _StyleDefs(192) =   "Splits(1).Columns(5).FooterStyle:id=180,.parent=153"
         _StyleDefs(193) =   "Splits(1).Columns(5).EditorStyle:id=181,.parent=155"
         _StyleDefs(194) =   "Splits(1).Columns(6).Style:id=186,.parent=151,.locked=-1"
         _StyleDefs(195) =   "Splits(1).Columns(6).HeadingStyle:id=183,.parent=152,.alignment=2"
         _StyleDefs(196) =   "Splits(1).Columns(6).FooterStyle:id=184,.parent=153"
         _StyleDefs(197) =   "Splits(1).Columns(6).EditorStyle:id=185,.parent=155"
         _StyleDefs(198) =   "Splits(1).Columns(7).Style:id=190,.parent=151"
         _StyleDefs(199) =   "Splits(1).Columns(7).HeadingStyle:id=187,.parent=152,.alignment=2"
         _StyleDefs(200) =   "Splits(1).Columns(7).FooterStyle:id=188,.parent=153"
         _StyleDefs(201) =   "Splits(1).Columns(7).EditorStyle:id=189,.parent=155"
         _StyleDefs(202) =   "Splits(1).Columns(8).Style:id=194,.parent=151,.locked=-1"
         _StyleDefs(203) =   "Splits(1).Columns(8).HeadingStyle:id=191,.parent=152,.alignment=2"
         _StyleDefs(204) =   "Splits(1).Columns(8).FooterStyle:id=192,.parent=153"
         _StyleDefs(205) =   "Splits(1).Columns(8).EditorStyle:id=193,.parent=155"
         _StyleDefs(206) =   "Splits(1).Columns(9).Style:id=198,.parent=151"
         _StyleDefs(207) =   "Splits(1).Columns(9).HeadingStyle:id=195,.parent=152"
         _StyleDefs(208) =   "Splits(1).Columns(9).FooterStyle:id=196,.parent=153"
         _StyleDefs(209) =   "Splits(1).Columns(9).EditorStyle:id=197,.parent=155"
         _StyleDefs(210) =   "Splits(1).Columns(10).Style:id=202,.parent=151,.locked=-1"
         _StyleDefs(211) =   "Splits(1).Columns(10).HeadingStyle:id=199,.parent=152,.alignment=2"
         _StyleDefs(212) =   "Splits(1).Columns(10).FooterStyle:id=200,.parent=153"
         _StyleDefs(213) =   "Splits(1).Columns(10).EditorStyle:id=201,.parent=155"
         _StyleDefs(214) =   "Splits(1).Columns(11).Style:id=206,.parent=151"
         _StyleDefs(215) =   "Splits(1).Columns(11).HeadingStyle:id=203,.parent=152"
         _StyleDefs(216) =   "Splits(1).Columns(11).FooterStyle:id=204,.parent=153"
         _StyleDefs(217) =   "Splits(1).Columns(11).EditorStyle:id=205,.parent=155"
         _StyleDefs(218) =   "Splits(1).Columns(12).Style:id=210,.parent=151,.locked=-1"
         _StyleDefs(219) =   "Splits(1).Columns(12).HeadingStyle:id=207,.parent=152,.alignment=2"
         _StyleDefs(220) =   "Splits(1).Columns(12).FooterStyle:id=208,.parent=153"
         _StyleDefs(221) =   "Splits(1).Columns(12).EditorStyle:id=209,.parent=155"
         _StyleDefs(222) =   "Splits(1).Columns(13).Style:id=214,.parent=151"
         _StyleDefs(223) =   "Splits(1).Columns(13).HeadingStyle:id=211,.parent=152"
         _StyleDefs(224) =   "Splits(1).Columns(13).FooterStyle:id=212,.parent=153"
         _StyleDefs(225) =   "Splits(1).Columns(13).EditorStyle:id=213,.parent=155"
         _StyleDefs(226) =   "Splits(1).Columns(14).Style:id=218,.parent=151,.locked=-1"
         _StyleDefs(227) =   "Splits(1).Columns(14).HeadingStyle:id=215,.parent=152,.alignment=2"
         _StyleDefs(228) =   "Splits(1).Columns(14).FooterStyle:id=216,.parent=153"
         _StyleDefs(229) =   "Splits(1).Columns(14).EditorStyle:id=217,.parent=155"
         _StyleDefs(230) =   "Splits(1).Columns(15).Style:id=222,.parent=151"
         _StyleDefs(231) =   "Splits(1).Columns(15).HeadingStyle:id=219,.parent=152"
         _StyleDefs(232) =   "Splits(1).Columns(15).FooterStyle:id=220,.parent=153"
         _StyleDefs(233) =   "Splits(1).Columns(15).EditorStyle:id=221,.parent=155"
         _StyleDefs(234) =   "Splits(1).Columns(16).Style:id=226,.parent=151,.locked=-1"
         _StyleDefs(235) =   "Splits(1).Columns(16).HeadingStyle:id=223,.parent=152,.alignment=2"
         _StyleDefs(236) =   "Splits(1).Columns(16).FooterStyle:id=224,.parent=153"
         _StyleDefs(237) =   "Splits(1).Columns(16).EditorStyle:id=225,.parent=155"
         _StyleDefs(238) =   "Splits(1).Columns(17).Style:id=230,.parent=151"
         _StyleDefs(239) =   "Splits(1).Columns(17).HeadingStyle:id=227,.parent=152"
         _StyleDefs(240) =   "Splits(1).Columns(17).FooterStyle:id=228,.parent=153"
         _StyleDefs(241) =   "Splits(1).Columns(17).EditorStyle:id=229,.parent=155"
         _StyleDefs(242) =   "Splits(1).Columns(18).Style:id=234,.parent=151,.locked=-1"
         _StyleDefs(243) =   "Splits(1).Columns(18).HeadingStyle:id=231,.parent=152,.alignment=2"
         _StyleDefs(244) =   "Splits(1).Columns(18).FooterStyle:id=232,.parent=153"
         _StyleDefs(245) =   "Splits(1).Columns(18).EditorStyle:id=233,.parent=155"
         _StyleDefs(246) =   "Splits(1).Columns(19).Style:id=238,.parent=151"
         _StyleDefs(247) =   "Splits(1).Columns(19).HeadingStyle:id=235,.parent=152"
         _StyleDefs(248) =   "Splits(1).Columns(19).FooterStyle:id=236,.parent=153"
         _StyleDefs(249) =   "Splits(1).Columns(19).EditorStyle:id=237,.parent=155"
         _StyleDefs(250) =   "Splits(1).Columns(20).Style:id=242,.parent=151,.locked=-1"
         _StyleDefs(251) =   "Splits(1).Columns(20).HeadingStyle:id=239,.parent=152,.alignment=2"
         _StyleDefs(252) =   "Splits(1).Columns(20).FooterStyle:id=240,.parent=153"
         _StyleDefs(253) =   "Splits(1).Columns(20).EditorStyle:id=241,.parent=155"
         _StyleDefs(254) =   "Splits(1).Columns(21).Style:id=246,.parent=151"
         _StyleDefs(255) =   "Splits(1).Columns(21).HeadingStyle:id=243,.parent=152"
         _StyleDefs(256) =   "Splits(1).Columns(21).FooterStyle:id=244,.parent=153"
         _StyleDefs(257) =   "Splits(1).Columns(21).EditorStyle:id=245,.parent=155"
         _StyleDefs(258) =   "Splits(1).Columns(22).Style:id=250,.parent=151,.locked=-1"
         _StyleDefs(259) =   "Splits(1).Columns(22).HeadingStyle:id=247,.parent=152,.alignment=2"
         _StyleDefs(260) =   "Splits(1).Columns(22).FooterStyle:id=248,.parent=153"
         _StyleDefs(261) =   "Splits(1).Columns(22).EditorStyle:id=249,.parent=155"
         _StyleDefs(262) =   "Splits(1).Columns(23).Style:id=254,.parent=151"
         _StyleDefs(263) =   "Splits(1).Columns(23).HeadingStyle:id=251,.parent=152"
         _StyleDefs(264) =   "Splits(1).Columns(23).FooterStyle:id=252,.parent=153"
         _StyleDefs(265) =   "Splits(1).Columns(23).EditorStyle:id=253,.parent=155"
         _StyleDefs(266) =   "Splits(1).Columns(24).Style:id=258,.parent=151,.locked=-1"
         _StyleDefs(267) =   "Splits(1).Columns(24).HeadingStyle:id=255,.parent=152,.alignment=2"
         _StyleDefs(268) =   "Splits(1).Columns(24).FooterStyle:id=256,.parent=153"
         _StyleDefs(269) =   "Splits(1).Columns(24).EditorStyle:id=257,.parent=155"
         _StyleDefs(270) =   "Splits(1).Columns(25).Style:id=262,.parent=151"
         _StyleDefs(271) =   "Splits(1).Columns(25).HeadingStyle:id=259,.parent=152"
         _StyleDefs(272) =   "Splits(1).Columns(25).FooterStyle:id=260,.parent=153"
         _StyleDefs(273) =   "Splits(1).Columns(25).EditorStyle:id=261,.parent=155"
         _StyleDefs(274) =   "Splits(1).Columns(26).Style:id=266,.parent=151,.locked=-1"
         _StyleDefs(275) =   "Splits(1).Columns(26).HeadingStyle:id=263,.parent=152,.alignment=2"
         _StyleDefs(276) =   "Splits(1).Columns(26).FooterStyle:id=264,.parent=153"
         _StyleDefs(277) =   "Splits(1).Columns(26).EditorStyle:id=265,.parent=155"
         _StyleDefs(278) =   "Splits(1).Columns(27).Style:id=270,.parent=151"
         _StyleDefs(279) =   "Splits(1).Columns(27).HeadingStyle:id=267,.parent=152"
         _StyleDefs(280) =   "Splits(1).Columns(27).FooterStyle:id=268,.parent=153"
         _StyleDefs(281) =   "Splits(1).Columns(27).EditorStyle:id=269,.parent=155"
         _StyleDefs(282) =   "Splits(1).Columns(28).Style:id=274,.parent=151,.locked=-1"
         _StyleDefs(283) =   "Splits(1).Columns(28).HeadingStyle:id=271,.parent=152,.alignment=2"
         _StyleDefs(284) =   "Splits(1).Columns(28).FooterStyle:id=272,.parent=153"
         _StyleDefs(285) =   "Splits(1).Columns(28).EditorStyle:id=273,.parent=155"
         _StyleDefs(286) =   "Splits(1).Columns(29).Style:id=278,.parent=151"
         _StyleDefs(287) =   "Splits(1).Columns(29).HeadingStyle:id=275,.parent=152"
         _StyleDefs(288) =   "Splits(1).Columns(29).FooterStyle:id=276,.parent=153"
         _StyleDefs(289) =   "Splits(1).Columns(29).EditorStyle:id=277,.parent=155"
         _StyleDefs(290) =   "Named:id=33:Normal"
         _StyleDefs(291) =   ":id=33,.parent=0"
         _StyleDefs(292) =   "Named:id=34:Heading"
         _StyleDefs(293) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(294) =   ":id=34,.wraptext=-1"
         _StyleDefs(295) =   "Named:id=35:Footing"
         _StyleDefs(296) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(297) =   "Named:id=36:Selected"
         _StyleDefs(298) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(299) =   "Named:id=37:Caption"
         _StyleDefs(300) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(301) =   "Named:id=38:HighlightRow"
         _StyleDefs(302) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(303) =   "Named:id=39:EvenRow"
         _StyleDefs(304) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(305) =   "Named:id=40:OddRow"
         _StyleDefs(306) =   ":id=40,.parent=33"
         _StyleDefs(307) =   "Named:id=41:RecordSelector"
         _StyleDefs(308) =   ":id=41,.parent=34"
         _StyleDefs(309) =   "Named:id=42:FilterBar"
         _StyleDefs(310) =   ":id=42,.parent=33"
      End
      Begin ucXPButton.XPButton cmdProses 
         Height          =   495
         Left            =   9960
         TabIndex        =   58
         Top             =   7920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
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
         MICON           =   "frmSIDSearchCust.frx":36903
         PICN            =   "frmSIDSearchCust.frx":3691F
         PICH            =   "frmSIDSearchCust.frx":3D33C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdTutupDuplikasi 
         Height          =   495
         Left            =   11160
         TabIndex        =   59
         Top             =   7920
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
         MICON           =   "frmSIDSearchCust.frx":43F59
         PICN            =   "frmSIDSearchCust.frx":43F75
         PICH            =   "frmSIDSearchCust.frx":44449
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   12510
      Begin ucXPButton.XPButton cmdSplit 
         Height          =   495
         Left            =   7320
         TabIndex        =   65
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
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
         MICON           =   "frmSIDSearchCust.frx":44994
         PICN            =   "frmSIDSearchCust.frx":449B0
         PICH            =   "frmSIDSearchCust.frx":47F9D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame frameLoading 
         BackColor       =   &H80000016&
         Height          =   1095
         Left            =   0
         TabIndex        =   52
         Top             =   2160
         Visible         =   0   'False
         Width           =   12135
         Begin ARProgBarCtrl.ARProgressBar progBar 
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   661
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Border          =   2
            Caption         =   " "
         End
         Begin VB.Label lblKeteranganLoading 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
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
            TabIndex        =   54
            Top             =   120
            Width           =   11895
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   6630
         TabIndex        =   26
         Top             =   90
         Width           =   5820
         Begin ucXPButton.XPButton cmdBuatIDBaru 
            Height          =   495
            Left            =   3960
            TabIndex        =   66
            Top             =   2760
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
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
            MICON           =   "frmSIDSearchCust.frx":4B656
            PICN            =   "frmSIDSearchCust.frx":4B672
            PICH            =   "frmSIDSearchCust.frx":4EE46
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdDijual 
            Height          =   495
            Left            =   4320
            TabIndex        =   67
            Top             =   2760
            Width           =   1140
            _ExtentX        =   2011
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
            MICON           =   "frmSIDSearchCust.frx":52862
            PICN            =   "frmSIDSearchCust.frx":5287E
            PICH            =   "frmSIDSearchCust.frx":55C44
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   11
            Left            =   1785
            MaxLength       =   160
            TabIndex        =   49
            Top             =   2640
            Width           =   3225
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   9
            Left            =   1785
            MaxLength       =   50
            TabIndex        =   27
            Top             =   45
            Width           =   3435
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   8
            Left            =   1785
            MaxLength       =   20
            TabIndex        =   28
            Top             =   375
            Width           =   3225
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   7
            Left            =   1785
            MaxLength       =   25
            TabIndex        =   29
            Top             =   705
            Width           =   3225
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   6
            Left            =   1785
            MaxLength       =   160
            TabIndex        =   30
            Top             =   1035
            Width           =   3225
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Index           =   5
            Left            =   1785
            MaxLength       =   13
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   1710
            Width           =   3630
         End
         Begin TDBDate6Ctl.TDBDate tglLahirx 
            Height          =   330
            Left            =   1785
            TabIndex        =   31
            Top             =   1365
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   582
            Calendar        =   "frmSIDSearchCust.frx":591B4
            Caption         =   "frmSIDSearchCust.frx":592E0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSIDSearchCust.frx":5934C
            Keys            =   "frmSIDSearchCust.frx":5936A
            Spin            =   "frmSIDSearchCust.frx":593C8
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
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   51
            Top             =   2640
            Width           =   165
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "No Telp"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   50
            Top             =   2655
            Width           =   1620
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H000080FF&
            Height          =   3000
            Left            =   0
            Top             =   0
            Width           =   5460
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "No ID"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   44
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "No NPWP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   43
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Code Customer"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Customer "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   41
            Top             =   1050
            Width           =   1620
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Lahir  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   40
            Top             =   1380
            Width           =   1635
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   39
            Top             =   1725
            Width           =   1620
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1695
            TabIndex        =   38
            Top             =   60
            Width           =   165
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1695
            TabIndex        =   37
            Top             =   385
            Width           =   165
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1695
            TabIndex        =   36
            Top             =   715
            Width           =   165
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1695
            TabIndex        =   35
            Top             =   1050
            Width           =   165
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1695
            TabIndex        =   34
            Top             =   1380
            Width           =   165
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1695
            TabIndex        =   33
            Top             =   1720
            Width           =   165
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2955
         Left            =   75
         TabIndex        =   1
         Top             =   120
         Width           =   6360
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   10
            Left            =   1860
            MaxLength       =   15
            TabIndex        =   46
            Top             =   2040
            Width           =   2970
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   4
            Left            =   1860
            MaxLength       =   13
            TabIndex        =   12
            Top             =   1710
            Width           =   2970
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   3
            Left            =   1860
            MaxLength       =   160
            TabIndex        =   11
            Top             =   1035
            Width           =   2970
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   2
            Left            =   1860
            MaxLength       =   25
            TabIndex        =   10
            Top             =   705
            Width           =   2970
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   1
            Left            =   1860
            MaxLength       =   20
            TabIndex        =   9
            Top             =   375
            Width           =   2970
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   0
            Left            =   1860
            MaxLength       =   12
            TabIndex        =   3
            Top             =   45
            Width           =   1740
         End
         Begin TDBDate6Ctl.TDBDate tglLahir 
            Height          =   330
            Left            =   1860
            TabIndex        =   13
            Top             =   1365
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   582
            Calendar        =   "frmSIDSearchCust.frx":593F0
            Caption         =   "frmSIDSearchCust.frx":5951C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSIDSearchCust.frx":59588
            Keys            =   "frmSIDSearchCust.frx":595A6
            Spin            =   "frmSIDSearchCust.frx":59604
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
         Begin ucXPButton.XPButton btnCari 
            Height          =   495
            Left            =   45
            TabIndex        =   21
            Top             =   2400
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&s"
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
            MICON           =   "frmSIDSearchCust.frx":5962C
            PICN            =   "frmSIDSearchCust.frx":59648
            PICH            =   "frmSIDSearchCust.frx":59BB6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton btnTambah 
            Height          =   495
            Left            =   2460
            TabIndex        =   22
            Top             =   2400
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&a"
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
            MICON           =   "frmSIDSearchCust.frx":5A1D3
            PICN            =   "frmSIDSearchCust.frx":5A1EF
            PICH            =   "frmSIDSearchCust.frx":5A63B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton btnKeluar 
            Height          =   495
            Left            =   3655
            TabIndex        =   23
            Top             =   2400
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
            MICON           =   "frmSIDSearchCust.frx":5AB03
            PICN            =   "frmSIDSearchCust.frx":5AB1F
            PICH            =   "frmSIDSearchCust.frx":5AFF3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton btnClear 
            Height          =   495
            Left            =   1245
            TabIndex        =   45
            Top             =   2400
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&s"
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
            MICON           =   "frmSIDSearchCust.frx":5B53E
            PICN            =   "frmSIDSearchCust.frx":5B55A
            PICH            =   "frmSIDSearchCust.frx":5BAAB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdDuplikasi 
            Height          =   495
            Left            =   4800
            TabIndex        =   55
            Top             =   2400
            Width           =   1545
            _ExtentX        =   2725
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
            MICON           =   "frmSIDSearchCust.frx":5C07C
            PICN            =   "frmSIDSearchCust.frx":5C098
            PICH            =   "frmSIDSearchCust.frx":5F76C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1725
            TabIndex        =   48
            Top             =   2040
            Width           =   165
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "No Telp"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   47
            Top             =   2055
            Width           =   1620
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1725
            TabIndex        =   20
            Top             =   1720
            Width           =   165
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1725
            TabIndex        =   19
            Top             =   1380
            Width           =   165
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1725
            TabIndex        =   18
            Top             =   1045
            Width           =   165
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1725
            TabIndex        =   17
            Top             =   715
            Width           =   165
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1725
            TabIndex        =   16
            Top             =   385
            Width           =   165
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1725
            TabIndex        =   15
            Top             =   60
            Width           =   165
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Customer "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   8
            Top             =   1720
            Width           =   1620
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Lahir  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   7
            Top             =   1380
            Width           =   1635
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Customer "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Top             =   1045
            Width           =   1620
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "No KTP / NPWP "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   715
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No Rangka "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   4
            Top             =   385
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor Polisi"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   2
            Top             =   60
            Width           =   1305
         End
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H000080FF&
         Height          =   3015
         Left            =   0
         Top             =   90
         Width           =   6450
      End
   End
   Begin ARProgBarCtrl.ARProgressBar ARProgressBar1 
      Height          =   210
      Left            =   60
      TabIndex        =   24
      Top             =   3495
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   370
      Value           =   0
      ForeColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0 %"
   End
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   5250
      Left            =   60
      TabIndex        =   25
      Top             =   3720
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   9260
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Detail"
      Columns(0).DataField=   "Detail"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Proses"
      Columns(1).DataField=   "Proses"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Kode Customer"
      Columns(2).DataField=   "IDCust"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Title"
      Columns(3).DataField=   "Title"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Customer"
      Columns(4).DataField=   "NamaCust"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Tgl Lahir"
      Columns(5).DataField=   "TglLahir"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Agama"
      Columns(6).DataField=   "Agama"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Pekerjaan"
      Columns(7).DataField=   "Pekerjaan"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Propinsi"
      Columns(8).DataField=   "Propinsi"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Kecamatan"
      Columns(9).DataField=   "Kecamatan"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Kelurahan"
      Columns(10).DataField=   "Kelurahan"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Kode Pos"
      Columns(11).DataField=   "KodePos"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "No Polisi"
      Columns(12).DataField=   "Nopolisi"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Alamat"
      Columns(13).DataField=   "Alamat"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Kota"
      Columns(14).DataField=   "Kota"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "RT"
      Columns(15).DataField=   "RT"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "RW"
      Columns(16).DataField=   "RW"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Kode Telp"
      Columns(17).DataField=   "KodeTelp"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Telp"
      Columns(18).DataField=   "Telp"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "HP"
      Columns(19).DataField=   "HP"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "HP2"
      Columns(20).DataField=   "HP2"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Email"
      Columns(21).DataField=   "Email"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "No. ID/NPWP"
      Columns(22).DataField=   "NoID"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Hobi"
      Columns(23).DataField=   "Hobi"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Olah Raga"
      Columns(24).DataField=   "Olahraga"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "Suku"
      Columns(25).DataField=   "Suku"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "Preferensi"
      Columns(26).DataField=   "Preferensi"
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "No Polish"
      Columns(27).DataField=   "NoPolis"
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "Tgl Awal Polish"
      Columns(28).DataField=   "TglAwalPolis"
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "Tgl Akhir Polish"
      Columns(29).DataField=   "TglAkhirPolis"
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "No Rangka"
      Columns(30).DataField=   "NoRangka"
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "No Engine"
      Columns(31).DataField=   "NoEngine"
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "No Key"
      Columns(32).DataField=   "NoKey"
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "Model"
      Columns(33).DataField=   "Mdl"
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "Class"
      Columns(34).DataField=   "Class"
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "Full Model"
      Columns(35).DataField=   "Fmdl"
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).Caption=   "Color"
      Columns(36).DataField=   "Color"
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).Caption=   "Kode Dealer"
      Columns(37).DataField=   "KdDlrPenjual"
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).Caption=   "Nama Dealer"
      Columns(38).DataField=   "NmDlrPenjual"
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).Caption=   "DDate"
      Columns(39).DataField=   "DDate"
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).Caption=   "NONPWP"
      Columns(40).DataField=   "NONPWP"
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).Caption=   "NOKTP"
      Columns(41).DataField=   "NOKTP"
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(42)._VlistStyle=   0
      Columns(42)._MaxComboItems=   5
      Columns(42).Caption=   "Split"
      Columns(42).DataField=   "split"
      Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(43)._VlistStyle=   0
      Columns(43)._MaxComboItems=   5
      Columns(43).Caption=   "Status Unit"
      Columns(43).DataField=   "StatusUnit"
      Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   44
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).SizeMode=   2
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=44"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0).ButtonText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).ButtonAlways=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1667"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(12)=   "Column(1).Button=1"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(1).ButtonText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).ButtonAlways=1"
      Splits(0)._ColumnProps(16)=   "Column(1).ButtonHeader=1"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=2752"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=8708"
      Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=8708"
      Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=4657"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=4551"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=8708"
      Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(35)=   "Column(5).Width=2275"
      Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=2170"
      Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=8708"
      Splits(0)._ColumnProps(39)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(41)=   "Column(6).Width=159"
      Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._WidthInPix=53"
      Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=8708"
      Splits(0)._ColumnProps(45)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(47)=   "Column(7).Width=3360"
      Splits(0)._ColumnProps(48)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._WidthInPix=3254"
      Splits(0)._ColumnProps(50)=   "Column(7)._ColStyle=8708"
      Splits(0)._ColumnProps(51)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(52)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(53)=   "Column(8).Width=3281"
      Splits(0)._ColumnProps(54)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(56)=   "Column(8)._ColStyle=8708"
      Splits(0)._ColumnProps(57)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(58)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(59)=   "Column(9).Width=2646"
      Splits(0)._ColumnProps(60)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(9)._WidthInPix=2540"
      Splits(0)._ColumnProps(62)=   "Column(9)._ColStyle=8708"
      Splits(0)._ColumnProps(63)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(64)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(65)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(10)._ColStyle=8708"
      Splits(0)._ColumnProps(69)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(71)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(72)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(74)=   "Column(11)._ColStyle=8708"
      Splits(0)._ColumnProps(75)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(76)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(77)=   "Column(12).Width=1746"
      Splits(0)._ColumnProps(78)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(12)._WidthInPix=1640"
      Splits(0)._ColumnProps(80)=   "Column(12)._ColStyle=8708"
      Splits(0)._ColumnProps(81)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(83)=   "Column(13).Width=6271"
      Splits(0)._ColumnProps(84)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(13)._WidthInPix=6165"
      Splits(0)._ColumnProps(86)=   "Column(13)._ColStyle=8708"
      Splits(0)._ColumnProps(87)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(88)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(89)=   "Column(14).Width=2963"
      Splits(0)._ColumnProps(90)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(14)._WidthInPix=2858"
      Splits(0)._ColumnProps(92)=   "Column(14)._ColStyle=8708"
      Splits(0)._ColumnProps(93)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(94)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(95)=   "Column(15).Width=7064"
      Splits(0)._ColumnProps(96)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(15)._WidthInPix=6959"
      Splits(0)._ColumnProps(98)=   "Column(15)._ColStyle=8708"
      Splits(0)._ColumnProps(99)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(100)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(101)=   "Column(16).Width=2461"
      Splits(0)._ColumnProps(102)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(16)._WidthInPix=2355"
      Splits(0)._ColumnProps(104)=   "Column(16)._ColStyle=8708"
      Splits(0)._ColumnProps(105)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(106)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(107)=   "Column(17).Width=2646"
      Splits(0)._ColumnProps(108)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(109)=   "Column(17)._WidthInPix=2540"
      Splits(0)._ColumnProps(110)=   "Column(17)._ColStyle=8708"
      Splits(0)._ColumnProps(111)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(112)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(113)=   "Column(18).Width=3281"
      Splits(0)._ColumnProps(114)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(18)._WidthInPix=3175"
      Splits(0)._ColumnProps(116)=   "Column(18)._ColStyle=8708"
      Splits(0)._ColumnProps(117)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(118)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(119)=   "Column(19).Width=3784"
      Splits(0)._ColumnProps(120)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(19)._WidthInPix=3678"
      Splits(0)._ColumnProps(122)=   "Column(19)._ColStyle=8708"
      Splits(0)._ColumnProps(123)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(124)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(125)=   "Column(20).Width=2646"
      Splits(0)._ColumnProps(126)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(20)._WidthInPix=2540"
      Splits(0)._ColumnProps(128)=   "Column(20)._ColStyle=8708"
      Splits(0)._ColumnProps(129)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(130)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(131)=   "Column(21).Width=3281"
      Splits(0)._ColumnProps(132)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(21)._WidthInPix=3175"
      Splits(0)._ColumnProps(134)=   "Column(21)._ColStyle=8708"
      Splits(0)._ColumnProps(135)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(136)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(137)=   "Column(22).Width=4789"
      Splits(0)._ColumnProps(138)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(22)._WidthInPix=4683"
      Splits(0)._ColumnProps(140)=   "Column(22)._ColStyle=8708"
      Splits(0)._ColumnProps(141)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(142)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(143)=   "Column(23).Width=873"
      Splits(0)._ColumnProps(144)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(23)._WidthInPix=767"
      Splits(0)._ColumnProps(146)=   "Column(23)._ColStyle=8708"
      Splits(0)._ColumnProps(147)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(148)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(149)=   "Column(24).Width=1773"
      Splits(0)._ColumnProps(150)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(24)._WidthInPix=1667"
      Splits(0)._ColumnProps(152)=   "Column(24)._ColStyle=8708"
      Splits(0)._ColumnProps(153)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(154)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(155)=   "Column(25).Width=3281"
      Splits(0)._ColumnProps(156)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(25)._WidthInPix=3175"
      Splits(0)._ColumnProps(158)=   "Column(25)._ColStyle=8708"
      Splits(0)._ColumnProps(159)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(160)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(161)=   "Column(26).Width=3281"
      Splits(0)._ColumnProps(162)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(26)._WidthInPix=3175"
      Splits(0)._ColumnProps(164)=   "Column(26)._ColStyle=8196"
      Splits(0)._ColumnProps(165)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(166)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(167)=   "Column(27).Width=3519"
      Splits(0)._ColumnProps(168)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(169)=   "Column(27)._WidthInPix=3413"
      Splits(0)._ColumnProps(170)=   "Column(27)._ColStyle=8708"
      Splits(0)._ColumnProps(171)=   "Column(27).Visible=0"
      Splits(0)._ColumnProps(172)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(173)=   "Column(28).Width=3519"
      Splits(0)._ColumnProps(174)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(28)._WidthInPix=3413"
      Splits(0)._ColumnProps(176)=   "Column(28)._ColStyle=8708"
      Splits(0)._ColumnProps(177)=   "Column(28).Visible=0"
      Splits(0)._ColumnProps(178)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(179)=   "Column(29).Width=3519"
      Splits(0)._ColumnProps(180)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(181)=   "Column(29)._WidthInPix=3413"
      Splits(0)._ColumnProps(182)=   "Column(29)._ColStyle=8708"
      Splits(0)._ColumnProps(183)=   "Column(29).Visible=0"
      Splits(0)._ColumnProps(184)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(185)=   "Column(30).Width=4366"
      Splits(0)._ColumnProps(186)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(30)._WidthInPix=4260"
      Splits(0)._ColumnProps(188)=   "Column(30)._ColStyle=8708"
      Splits(0)._ColumnProps(189)=   "Column(30).Visible=0"
      Splits(0)._ColumnProps(190)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(191)=   "Column(31).Width=873"
      Splits(0)._ColumnProps(192)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(193)=   "Column(31)._WidthInPix=767"
      Splits(0)._ColumnProps(194)=   "Column(31)._ColStyle=8708"
      Splits(0)._ColumnProps(195)=   "Column(31).Visible=0"
      Splits(0)._ColumnProps(196)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(197)=   "Column(32).Width=3519"
      Splits(0)._ColumnProps(198)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(199)=   "Column(32)._WidthInPix=3413"
      Splits(0)._ColumnProps(200)=   "Column(32)._ColStyle=8708"
      Splits(0)._ColumnProps(201)=   "Column(32).Visible=0"
      Splits(0)._ColumnProps(202)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(203)=   "Column(33).Width=3519"
      Splits(0)._ColumnProps(204)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(205)=   "Column(33)._WidthInPix=3413"
      Splits(0)._ColumnProps(206)=   "Column(33)._ColStyle=8708"
      Splits(0)._ColumnProps(207)=   "Column(33).Visible=0"
      Splits(0)._ColumnProps(208)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(209)=   "Column(34).Width=5292"
      Splits(0)._ColumnProps(210)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(34)._WidthInPix=5186"
      Splits(0)._ColumnProps(212)=   "Column(34)._ColStyle=8708"
      Splits(0)._ColumnProps(213)=   "Column(34).Visible=0"
      Splits(0)._ColumnProps(214)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(215)=   "Column(35).Width=3281"
      Splits(0)._ColumnProps(216)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(217)=   "Column(35)._WidthInPix=3175"
      Splits(0)._ColumnProps(218)=   "Column(35)._ColStyle=8708"
      Splits(0)._ColumnProps(219)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(220)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(221)=   "Column(36).Width=3281"
      Splits(0)._ColumnProps(222)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(223)=   "Column(36)._WidthInPix=3175"
      Splits(0)._ColumnProps(224)=   "Column(36)._ColStyle=8708"
      Splits(0)._ColumnProps(225)=   "Column(36).Visible=0"
      Splits(0)._ColumnProps(226)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(227)=   "Column(37).Width=5292"
      Splits(0)._ColumnProps(228)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(229)=   "Column(37)._WidthInPix=5186"
      Splits(0)._ColumnProps(230)=   "Column(37)._ColStyle=8708"
      Splits(0)._ColumnProps(231)=   "Column(37).Visible=0"
      Splits(0)._ColumnProps(232)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(233)=   "Column(38).Width=3281"
      Splits(0)._ColumnProps(234)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(235)=   "Column(38)._WidthInPix=3175"
      Splits(0)._ColumnProps(236)=   "Column(38)._ColStyle=8708"
      Splits(0)._ColumnProps(237)=   "Column(38).Visible=0"
      Splits(0)._ColumnProps(238)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(239)=   "Column(39).Width=3281"
      Splits(0)._ColumnProps(240)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(241)=   "Column(39)._WidthInPix=3175"
      Splits(0)._ColumnProps(242)=   "Column(39)._ColStyle=8708"
      Splits(0)._ColumnProps(243)=   "Column(39).Visible=0"
      Splits(0)._ColumnProps(244)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(245)=   "Column(40).Width=3281"
      Splits(0)._ColumnProps(246)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(247)=   "Column(40)._WidthInPix=3175"
      Splits(0)._ColumnProps(248)=   "Column(40).Visible=0"
      Splits(0)._ColumnProps(249)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(250)=   "Column(41).Width=3281"
      Splits(0)._ColumnProps(251)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(252)=   "Column(41)._WidthInPix=3175"
      Splits(0)._ColumnProps(253)=   "Column(41).Visible=0"
      Splits(0)._ColumnProps(254)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(255)=   "Column(42).Width=3281"
      Splits(0)._ColumnProps(256)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(257)=   "Column(42)._WidthInPix=3175"
      Splits(0)._ColumnProps(258)=   "Column(42).Visible=0"
      Splits(0)._ColumnProps(259)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(260)=   "Column(42).ButtonText=1"
      Splits(0)._ColumnProps(261)=   "Column(43).Width=3281"
      Splits(0)._ColumnProps(262)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(263)=   "Column(43)._WidthInPix=3175"
      Splits(0)._ColumnProps(264)=   "Column(43).Visible=0"
      Splits(0)._ColumnProps(265)=   "Column(43).Order=44"
      Splits(1)._UserFlags=   0
      Splits(1).AllowSizing=   -1  'True
      Splits(1).RecordSelectorWidth=   688
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).AllowColSelect=   0   'False
      Splits(1).FetchRowStyle=   -1  'True
      Splits(1).AlternatingRowStyle=   -1  'True
      Splits(1).DividerColor=   14215660
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=44"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
      Splits(1)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(7)=   "Column(0).ButtonText=1"
      Splits(1)._ColumnProps(8)=   "Column(0).ButtonAlways=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=1773"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1667"
      Splits(1)._ColumnProps(12)=   "Column(1)._ColStyle=513"
      Splits(1)._ColumnProps(13)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(14)=   "Column(1).Button=1"
      Splits(1)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(16)=   "Column(1).ButtonText=1"
      Splits(1)._ColumnProps(17)=   "Column(1).ButtonAlways=1"
      Splits(1)._ColumnProps(18)=   "Column(1).ButtonHeader=1"
      Splits(1)._ColumnProps(19)=   "Column(2).Width=2752"
      Splits(1)._ColumnProps(20)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._WidthInPix=2646"
      Splits(1)._ColumnProps(22)=   "Column(2)._ColStyle=8708"
      Splits(1)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(24)=   "Column(3).Width=3281"
      Splits(1)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(26)=   "Column(3)._WidthInPix=3175"
      Splits(1)._ColumnProps(27)=   "Column(3)._ColStyle=8708"
      Splits(1)._ColumnProps(28)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(30)=   "Column(4).Width=4657"
      Splits(1)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(32)=   "Column(4)._WidthInPix=4551"
      Splits(1)._ColumnProps(33)=   "Column(4)._ColStyle=8708"
      Splits(1)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(35)=   "Column(5).Width=2275"
      Splits(1)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(37)=   "Column(5)._WidthInPix=2170"
      Splits(1)._ColumnProps(38)=   "Column(5)._ColStyle=8708"
      Splits(1)._ColumnProps(39)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(40)=   "Column(6).Width=159"
      Splits(1)._ColumnProps(41)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(42)=   "Column(6)._WidthInPix=53"
      Splits(1)._ColumnProps(43)=   "Column(6)._ColStyle=8708"
      Splits(1)._ColumnProps(44)=   "Column(6).Visible=0"
      Splits(1)._ColumnProps(45)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(46)=   "Column(7).Width=3360"
      Splits(1)._ColumnProps(47)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(48)=   "Column(7)._WidthInPix=3254"
      Splits(1)._ColumnProps(49)=   "Column(7)._ColStyle=8708"
      Splits(1)._ColumnProps(50)=   "Column(7).Visible=0"
      Splits(1)._ColumnProps(51)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(52)=   "Column(8).Width=3281"
      Splits(1)._ColumnProps(53)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(54)=   "Column(8)._WidthInPix=3175"
      Splits(1)._ColumnProps(55)=   "Column(8)._ColStyle=8708"
      Splits(1)._ColumnProps(56)=   "Column(8).Visible=0"
      Splits(1)._ColumnProps(57)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(58)=   "Column(9).Width=2646"
      Splits(1)._ColumnProps(59)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(60)=   "Column(9)._WidthInPix=2540"
      Splits(1)._ColumnProps(61)=   "Column(9)._ColStyle=8708"
      Splits(1)._ColumnProps(62)=   "Column(9).Visible=0"
      Splits(1)._ColumnProps(63)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(64)=   "Column(10).Width=3281"
      Splits(1)._ColumnProps(65)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(66)=   "Column(10)._WidthInPix=3175"
      Splits(1)._ColumnProps(67)=   "Column(10)._ColStyle=8708"
      Splits(1)._ColumnProps(68)=   "Column(10).Visible=0"
      Splits(1)._ColumnProps(69)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(70)=   "Column(11).Width=3281"
      Splits(1)._ColumnProps(71)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(72)=   "Column(11)._WidthInPix=3175"
      Splits(1)._ColumnProps(73)=   "Column(11)._ColStyle=8708"
      Splits(1)._ColumnProps(74)=   "Column(11).Visible=0"
      Splits(1)._ColumnProps(75)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(76)=   "Column(12).Width=1746"
      Splits(1)._ColumnProps(77)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(78)=   "Column(12)._WidthInPix=1640"
      Splits(1)._ColumnProps(79)=   "Column(12)._ColStyle=8708"
      Splits(1)._ColumnProps(80)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(81)=   "Column(13).Width=6271"
      Splits(1)._ColumnProps(82)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(13)._WidthInPix=6165"
      Splits(1)._ColumnProps(84)=   "Column(13)._ColStyle=8708"
      Splits(1)._ColumnProps(85)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(86)=   "Column(14).Width=2963"
      Splits(1)._ColumnProps(87)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(88)=   "Column(14)._WidthInPix=2858"
      Splits(1)._ColumnProps(89)=   "Column(14)._ColStyle=8708"
      Splits(1)._ColumnProps(90)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(91)=   "Column(15).Width=7064"
      Splits(1)._ColumnProps(92)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(93)=   "Column(15)._WidthInPix=6959"
      Splits(1)._ColumnProps(94)=   "Column(15)._ColStyle=8708"
      Splits(1)._ColumnProps(95)=   "Column(15).Visible=0"
      Splits(1)._ColumnProps(96)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(97)=   "Column(16).Width=2461"
      Splits(1)._ColumnProps(98)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(99)=   "Column(16)._WidthInPix=2355"
      Splits(1)._ColumnProps(100)=   "Column(16)._ColStyle=8708"
      Splits(1)._ColumnProps(101)=   "Column(16).Visible=0"
      Splits(1)._ColumnProps(102)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(103)=   "Column(17).Width=2646"
      Splits(1)._ColumnProps(104)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(105)=   "Column(17)._WidthInPix=2540"
      Splits(1)._ColumnProps(106)=   "Column(17)._ColStyle=8708"
      Splits(1)._ColumnProps(107)=   "Column(17).Visible=0"
      Splits(1)._ColumnProps(108)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(109)=   "Column(18).Width=3281"
      Splits(1)._ColumnProps(110)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(111)=   "Column(18)._WidthInPix=3175"
      Splits(1)._ColumnProps(112)=   "Column(18)._ColStyle=8708"
      Splits(1)._ColumnProps(113)=   "Column(18).Visible=0"
      Splits(1)._ColumnProps(114)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(115)=   "Column(19).Width=3784"
      Splits(1)._ColumnProps(116)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(117)=   "Column(19)._WidthInPix=3678"
      Splits(1)._ColumnProps(118)=   "Column(19)._ColStyle=8708"
      Splits(1)._ColumnProps(119)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(120)=   "Column(20).Width=2646"
      Splits(1)._ColumnProps(121)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(122)=   "Column(20)._WidthInPix=2540"
      Splits(1)._ColumnProps(123)=   "Column(20)._ColStyle=8708"
      Splits(1)._ColumnProps(124)=   "Column(20).Visible=0"
      Splits(1)._ColumnProps(125)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(126)=   "Column(21).Width=3281"
      Splits(1)._ColumnProps(127)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(128)=   "Column(21)._WidthInPix=3175"
      Splits(1)._ColumnProps(129)=   "Column(21)._ColStyle=8708"
      Splits(1)._ColumnProps(130)=   "Column(21).Visible=0"
      Splits(1)._ColumnProps(131)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(132)=   "Column(22).Width=4789"
      Splits(1)._ColumnProps(133)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(134)=   "Column(22)._WidthInPix=4683"
      Splits(1)._ColumnProps(135)=   "Column(22)._ColStyle=8708"
      Splits(1)._ColumnProps(136)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(137)=   "Column(23).Width=873"
      Splits(1)._ColumnProps(138)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(139)=   "Column(23)._WidthInPix=767"
      Splits(1)._ColumnProps(140)=   "Column(23)._ColStyle=8708"
      Splits(1)._ColumnProps(141)=   "Column(23).Visible=0"
      Splits(1)._ColumnProps(142)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(143)=   "Column(24).Width=1773"
      Splits(1)._ColumnProps(144)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(145)=   "Column(24)._WidthInPix=1667"
      Splits(1)._ColumnProps(146)=   "Column(24)._ColStyle=8708"
      Splits(1)._ColumnProps(147)=   "Column(24).Visible=0"
      Splits(1)._ColumnProps(148)=   "Column(24).Order=25"
      Splits(1)._ColumnProps(149)=   "Column(25).Width=3281"
      Splits(1)._ColumnProps(150)=   "Column(25).DividerColor=0"
      Splits(1)._ColumnProps(151)=   "Column(25)._WidthInPix=3175"
      Splits(1)._ColumnProps(152)=   "Column(25)._ColStyle=8708"
      Splits(1)._ColumnProps(153)=   "Column(25).Visible=0"
      Splits(1)._ColumnProps(154)=   "Column(25).Order=26"
      Splits(1)._ColumnProps(155)=   "Column(26).Width=3281"
      Splits(1)._ColumnProps(156)=   "Column(26).DividerColor=0"
      Splits(1)._ColumnProps(157)=   "Column(26)._WidthInPix=3175"
      Splits(1)._ColumnProps(158)=   "Column(26)._ColStyle=8196"
      Splits(1)._ColumnProps(159)=   "Column(26).Visible=0"
      Splits(1)._ColumnProps(160)=   "Column(26).Order=27"
      Splits(1)._ColumnProps(161)=   "Column(27).Width=3519"
      Splits(1)._ColumnProps(162)=   "Column(27).DividerColor=0"
      Splits(1)._ColumnProps(163)=   "Column(27)._WidthInPix=3413"
      Splits(1)._ColumnProps(164)=   "Column(27)._ColStyle=8708"
      Splits(1)._ColumnProps(165)=   "Column(27).Visible=0"
      Splits(1)._ColumnProps(166)=   "Column(27).Order=28"
      Splits(1)._ColumnProps(167)=   "Column(28).Width=3519"
      Splits(1)._ColumnProps(168)=   "Column(28).DividerColor=0"
      Splits(1)._ColumnProps(169)=   "Column(28)._WidthInPix=3413"
      Splits(1)._ColumnProps(170)=   "Column(28)._ColStyle=8708"
      Splits(1)._ColumnProps(171)=   "Column(28).Visible=0"
      Splits(1)._ColumnProps(172)=   "Column(28).Order=29"
      Splits(1)._ColumnProps(173)=   "Column(29).Width=3519"
      Splits(1)._ColumnProps(174)=   "Column(29).DividerColor=0"
      Splits(1)._ColumnProps(175)=   "Column(29)._WidthInPix=3413"
      Splits(1)._ColumnProps(176)=   "Column(29)._ColStyle=8708"
      Splits(1)._ColumnProps(177)=   "Column(29).Visible=0"
      Splits(1)._ColumnProps(178)=   "Column(29).Order=30"
      Splits(1)._ColumnProps(179)=   "Column(30).Width=4366"
      Splits(1)._ColumnProps(180)=   "Column(30).DividerColor=0"
      Splits(1)._ColumnProps(181)=   "Column(30)._WidthInPix=4260"
      Splits(1)._ColumnProps(182)=   "Column(30)._ColStyle=8708"
      Splits(1)._ColumnProps(183)=   "Column(30).Order=31"
      Splits(1)._ColumnProps(184)=   "Column(31).Width=873"
      Splits(1)._ColumnProps(185)=   "Column(31).DividerColor=0"
      Splits(1)._ColumnProps(186)=   "Column(31)._WidthInPix=767"
      Splits(1)._ColumnProps(187)=   "Column(31)._ColStyle=8708"
      Splits(1)._ColumnProps(188)=   "Column(31).Visible=0"
      Splits(1)._ColumnProps(189)=   "Column(31).Order=32"
      Splits(1)._ColumnProps(190)=   "Column(32).Width=3519"
      Splits(1)._ColumnProps(191)=   "Column(32).DividerColor=0"
      Splits(1)._ColumnProps(192)=   "Column(32)._WidthInPix=3413"
      Splits(1)._ColumnProps(193)=   "Column(32)._ColStyle=8708"
      Splits(1)._ColumnProps(194)=   "Column(32).Visible=0"
      Splits(1)._ColumnProps(195)=   "Column(32).Order=33"
      Splits(1)._ColumnProps(196)=   "Column(33).Width=3519"
      Splits(1)._ColumnProps(197)=   "Column(33).DividerColor=0"
      Splits(1)._ColumnProps(198)=   "Column(33)._WidthInPix=3413"
      Splits(1)._ColumnProps(199)=   "Column(33)._ColStyle=8708"
      Splits(1)._ColumnProps(200)=   "Column(33).Visible=0"
      Splits(1)._ColumnProps(201)=   "Column(33).Order=34"
      Splits(1)._ColumnProps(202)=   "Column(34).Width=5292"
      Splits(1)._ColumnProps(203)=   "Column(34).DividerColor=0"
      Splits(1)._ColumnProps(204)=   "Column(34)._WidthInPix=5186"
      Splits(1)._ColumnProps(205)=   "Column(34)._ColStyle=8708"
      Splits(1)._ColumnProps(206)=   "Column(34).Visible=0"
      Splits(1)._ColumnProps(207)=   "Column(34).Order=35"
      Splits(1)._ColumnProps(208)=   "Column(35).Width=3281"
      Splits(1)._ColumnProps(209)=   "Column(35).DividerColor=0"
      Splits(1)._ColumnProps(210)=   "Column(35)._WidthInPix=3175"
      Splits(1)._ColumnProps(211)=   "Column(35)._ColStyle=8708"
      Splits(1)._ColumnProps(212)=   "Column(35).Visible=0"
      Splits(1)._ColumnProps(213)=   "Column(35).Order=36"
      Splits(1)._ColumnProps(214)=   "Column(36).Width=3281"
      Splits(1)._ColumnProps(215)=   "Column(36).DividerColor=0"
      Splits(1)._ColumnProps(216)=   "Column(36)._WidthInPix=3175"
      Splits(1)._ColumnProps(217)=   "Column(36)._ColStyle=8708"
      Splits(1)._ColumnProps(218)=   "Column(36).Visible=0"
      Splits(1)._ColumnProps(219)=   "Column(36).Order=37"
      Splits(1)._ColumnProps(220)=   "Column(37).Width=5292"
      Splits(1)._ColumnProps(221)=   "Column(37).DividerColor=0"
      Splits(1)._ColumnProps(222)=   "Column(37)._WidthInPix=5186"
      Splits(1)._ColumnProps(223)=   "Column(37)._ColStyle=8708"
      Splits(1)._ColumnProps(224)=   "Column(37).Visible=0"
      Splits(1)._ColumnProps(225)=   "Column(37).Order=38"
      Splits(1)._ColumnProps(226)=   "Column(38).Width=3281"
      Splits(1)._ColumnProps(227)=   "Column(38).DividerColor=0"
      Splits(1)._ColumnProps(228)=   "Column(38)._WidthInPix=3175"
      Splits(1)._ColumnProps(229)=   "Column(38)._ColStyle=8708"
      Splits(1)._ColumnProps(230)=   "Column(38).Visible=0"
      Splits(1)._ColumnProps(231)=   "Column(38).Order=39"
      Splits(1)._ColumnProps(232)=   "Column(39).Width=3281"
      Splits(1)._ColumnProps(233)=   "Column(39).DividerColor=0"
      Splits(1)._ColumnProps(234)=   "Column(39)._WidthInPix=3175"
      Splits(1)._ColumnProps(235)=   "Column(39)._ColStyle=8708"
      Splits(1)._ColumnProps(236)=   "Column(39).Visible=0"
      Splits(1)._ColumnProps(237)=   "Column(39).Order=40"
      Splits(1)._ColumnProps(238)=   "Column(40).Width=3281"
      Splits(1)._ColumnProps(239)=   "Column(40).DividerColor=0"
      Splits(1)._ColumnProps(240)=   "Column(40)._WidthInPix=3175"
      Splits(1)._ColumnProps(241)=   "Column(40).Visible=0"
      Splits(1)._ColumnProps(242)=   "Column(40).Order=41"
      Splits(1)._ColumnProps(243)=   "Column(41).Width=3281"
      Splits(1)._ColumnProps(244)=   "Column(41).DividerColor=0"
      Splits(1)._ColumnProps(245)=   "Column(41)._WidthInPix=3175"
      Splits(1)._ColumnProps(246)=   "Column(41).Visible=0"
      Splits(1)._ColumnProps(247)=   "Column(41).Order=42"
      Splits(1)._ColumnProps(248)=   "Column(42).Width=1773"
      Splits(1)._ColumnProps(249)=   "Column(42).DividerColor=0"
      Splits(1)._ColumnProps(250)=   "Column(42)._WidthInPix=1667"
      Splits(1)._ColumnProps(251)=   "Column(42)._ColStyle=1"
      Splits(1)._ColumnProps(252)=   "Column(42).Order=43"
      Splits(1)._ColumnProps(253)=   "Column(42).ButtonText=1"
      Splits(1)._ColumnProps(254)=   "Column(42).ButtonAlways=1"
      Splits(1)._ColumnProps(255)=   "Column(43).Width=3281"
      Splits(1)._ColumnProps(256)=   "Column(43).DividerColor=0"
      Splits(1)._ColumnProps(257)=   "Column(43)._WidthInPix=3175"
      Splits(1)._ColumnProps(258)=   "Column(43).Order=44"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   14215660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=71,.parent=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=71,.strikethrough=0,.charset=0"
      _StyleDefs(26)  =   ":id=71,.fontname=Verdana"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=136,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=72,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=73,.parent=3,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(30)  =   ":id=73,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(31)  =   ":id=73,.fontname=Verdana"
      _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=74,.parent=5"
      _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=84,.parent=6"
      _StyleDefs(34)  =   "Splits(0).EditorStyle:id=83,.parent=7"
      _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=85,.parent=8"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=86,.parent=9,.bgcolor=&HC1EDFF&"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=135,.parent=10"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=137,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=138,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=190,.parent=71,.alignment=2,.fgcolor=&H0&"
      _StyleDefs(41)  =   ":id=190,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=187,.parent=72,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=188,.parent=73"
      _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=189,.parent=83"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=194,.parent=71,.alignment=2,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=191,.parent=72,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=192,.parent=73"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=193,.parent=83"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=198,.parent=71,.locked=-1"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=195,.parent=72,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=196,.parent=73"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=197,.parent=83"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=222,.parent=71,.locked=-1"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=219,.parent=72,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=220,.parent=73"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=221,.parent=83"
      _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=226,.parent=71,.locked=-1"
      _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=223,.parent=72,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=224,.parent=73"
      _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=225,.parent=83"
      _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=230,.parent=71,.locked=-1"
      _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=227,.parent=72,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=228,.parent=73"
      _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=229,.parent=83"
      _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=234,.parent=71,.locked=-1"
      _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=231,.parent=72,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=232,.parent=73"
      _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=233,.parent=83"
      _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=238,.parent=71,.locked=-1"
      _StyleDefs(70)  =   "Splits(0).Columns(7).HeadingStyle:id=235,.parent=72,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(7).FooterStyle:id=236,.parent=73"
      _StyleDefs(72)  =   "Splits(0).Columns(7).EditorStyle:id=237,.parent=83"
      _StyleDefs(73)  =   "Splits(0).Columns(8).Style:id=242,.parent=71,.locked=-1"
      _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=239,.parent=72,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=240,.parent=73"
      _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=241,.parent=83"
      _StyleDefs(77)  =   "Splits(0).Columns(9).Style:id=246,.parent=71,.locked=-1"
      _StyleDefs(78)  =   "Splits(0).Columns(9).HeadingStyle:id=243,.parent=72,.alignment=2"
      _StyleDefs(79)  =   "Splits(0).Columns(9).FooterStyle:id=244,.parent=73"
      _StyleDefs(80)  =   "Splits(0).Columns(9).EditorStyle:id=245,.parent=83"
      _StyleDefs(81)  =   "Splits(0).Columns(10).Style:id=250,.parent=71,.locked=-1"
      _StyleDefs(82)  =   "Splits(0).Columns(10).HeadingStyle:id=247,.parent=72,.alignment=2"
      _StyleDefs(83)  =   "Splits(0).Columns(10).FooterStyle:id=248,.parent=73"
      _StyleDefs(84)  =   "Splits(0).Columns(10).EditorStyle:id=249,.parent=83"
      _StyleDefs(85)  =   "Splits(0).Columns(11).Style:id=254,.parent=71,.locked=-1"
      _StyleDefs(86)  =   "Splits(0).Columns(11).HeadingStyle:id=251,.parent=72,.alignment=2"
      _StyleDefs(87)  =   "Splits(0).Columns(11).FooterStyle:id=252,.parent=73"
      _StyleDefs(88)  =   "Splits(0).Columns(11).EditorStyle:id=253,.parent=83"
      _StyleDefs(89)  =   "Splits(0).Columns(12).Style:id=258,.parent=71,.locked=-1"
      _StyleDefs(90)  =   "Splits(0).Columns(12).HeadingStyle:id=255,.parent=72,.alignment=2"
      _StyleDefs(91)  =   "Splits(0).Columns(12).FooterStyle:id=256,.parent=73"
      _StyleDefs(92)  =   "Splits(0).Columns(12).EditorStyle:id=257,.parent=83"
      _StyleDefs(93)  =   "Splits(0).Columns(13).Style:id=262,.parent=71,.locked=-1"
      _StyleDefs(94)  =   "Splits(0).Columns(13).HeadingStyle:id=259,.parent=72,.alignment=2"
      _StyleDefs(95)  =   "Splits(0).Columns(13).FooterStyle:id=260,.parent=73"
      _StyleDefs(96)  =   "Splits(0).Columns(13).EditorStyle:id=261,.parent=83"
      _StyleDefs(97)  =   "Splits(0).Columns(14).Style:id=266,.parent=71,.locked=-1"
      _StyleDefs(98)  =   "Splits(0).Columns(14).HeadingStyle:id=263,.parent=72,.alignment=2"
      _StyleDefs(99)  =   "Splits(0).Columns(14).FooterStyle:id=264,.parent=73"
      _StyleDefs(100) =   "Splits(0).Columns(14).EditorStyle:id=265,.parent=83"
      _StyleDefs(101) =   "Splits(0).Columns(15).Style:id=270,.parent=71,.locked=-1"
      _StyleDefs(102) =   "Splits(0).Columns(15).HeadingStyle:id=267,.parent=72,.alignment=2"
      _StyleDefs(103) =   "Splits(0).Columns(15).FooterStyle:id=268,.parent=73"
      _StyleDefs(104) =   "Splits(0).Columns(15).EditorStyle:id=269,.parent=83"
      _StyleDefs(105) =   "Splits(0).Columns(16).Style:id=274,.parent=71,.locked=-1"
      _StyleDefs(106) =   "Splits(0).Columns(16).HeadingStyle:id=271,.parent=72,.alignment=2"
      _StyleDefs(107) =   "Splits(0).Columns(16).FooterStyle:id=272,.parent=73"
      _StyleDefs(108) =   "Splits(0).Columns(16).EditorStyle:id=273,.parent=83"
      _StyleDefs(109) =   "Splits(0).Columns(17).Style:id=278,.parent=71,.locked=-1"
      _StyleDefs(110) =   "Splits(0).Columns(17).HeadingStyle:id=275,.parent=72,.alignment=2"
      _StyleDefs(111) =   "Splits(0).Columns(17).FooterStyle:id=276,.parent=73"
      _StyleDefs(112) =   "Splits(0).Columns(17).EditorStyle:id=277,.parent=83"
      _StyleDefs(113) =   "Splits(0).Columns(18).Style:id=282,.parent=71,.locked=-1"
      _StyleDefs(114) =   "Splits(0).Columns(18).HeadingStyle:id=279,.parent=72,.alignment=2"
      _StyleDefs(115) =   "Splits(0).Columns(18).FooterStyle:id=280,.parent=73"
      _StyleDefs(116) =   "Splits(0).Columns(18).EditorStyle:id=281,.parent=83"
      _StyleDefs(117) =   "Splits(0).Columns(19).Style:id=286,.parent=71,.locked=-1"
      _StyleDefs(118) =   "Splits(0).Columns(19).HeadingStyle:id=283,.parent=72,.alignment=2"
      _StyleDefs(119) =   "Splits(0).Columns(19).FooterStyle:id=284,.parent=73"
      _StyleDefs(120) =   "Splits(0).Columns(19).EditorStyle:id=285,.parent=83"
      _StyleDefs(121) =   "Splits(0).Columns(20).Style:id=290,.parent=71,.locked=-1"
      _StyleDefs(122) =   "Splits(0).Columns(20).HeadingStyle:id=287,.parent=72,.alignment=2"
      _StyleDefs(123) =   "Splits(0).Columns(20).FooterStyle:id=288,.parent=73"
      _StyleDefs(124) =   "Splits(0).Columns(20).EditorStyle:id=289,.parent=83"
      _StyleDefs(125) =   "Splits(0).Columns(21).Style:id=294,.parent=71,.locked=-1"
      _StyleDefs(126) =   "Splits(0).Columns(21).HeadingStyle:id=291,.parent=72,.alignment=2"
      _StyleDefs(127) =   "Splits(0).Columns(21).FooterStyle:id=292,.parent=73"
      _StyleDefs(128) =   "Splits(0).Columns(21).EditorStyle:id=293,.parent=83"
      _StyleDefs(129) =   "Splits(0).Columns(22).Style:id=298,.parent=71,.locked=-1"
      _StyleDefs(130) =   "Splits(0).Columns(22).HeadingStyle:id=295,.parent=72,.alignment=2"
      _StyleDefs(131) =   "Splits(0).Columns(22).FooterStyle:id=296,.parent=73"
      _StyleDefs(132) =   "Splits(0).Columns(22).EditorStyle:id=297,.parent=83"
      _StyleDefs(133) =   "Splits(0).Columns(23).Style:id=302,.parent=71,.locked=-1"
      _StyleDefs(134) =   "Splits(0).Columns(23).HeadingStyle:id=299,.parent=72,.alignment=2"
      _StyleDefs(135) =   "Splits(0).Columns(23).FooterStyle:id=300,.parent=73"
      _StyleDefs(136) =   "Splits(0).Columns(23).EditorStyle:id=301,.parent=83"
      _StyleDefs(137) =   "Splits(0).Columns(24).Style:id=306,.parent=71,.locked=-1"
      _StyleDefs(138) =   "Splits(0).Columns(24).HeadingStyle:id=303,.parent=72,.alignment=2"
      _StyleDefs(139) =   "Splits(0).Columns(24).FooterStyle:id=304,.parent=73"
      _StyleDefs(140) =   "Splits(0).Columns(24).EditorStyle:id=305,.parent=83"
      _StyleDefs(141) =   "Splits(0).Columns(25).Style:id=310,.parent=71,.locked=-1"
      _StyleDefs(142) =   "Splits(0).Columns(25).HeadingStyle:id=307,.parent=72,.alignment=2"
      _StyleDefs(143) =   "Splits(0).Columns(25).FooterStyle:id=308,.parent=73"
      _StyleDefs(144) =   "Splits(0).Columns(25).EditorStyle:id=309,.parent=83"
      _StyleDefs(145) =   "Splits(0).Columns(26).Style:id=314,.parent=71,.locked=-1"
      _StyleDefs(146) =   "Splits(0).Columns(26).HeadingStyle:id=311,.parent=72"
      _StyleDefs(147) =   "Splits(0).Columns(26).FooterStyle:id=312,.parent=73"
      _StyleDefs(148) =   "Splits(0).Columns(26).EditorStyle:id=313,.parent=83"
      _StyleDefs(149) =   "Splits(0).Columns(27).Style:id=318,.parent=71,.locked=-1"
      _StyleDefs(150) =   "Splits(0).Columns(27).HeadingStyle:id=315,.parent=72,.alignment=2"
      _StyleDefs(151) =   "Splits(0).Columns(27).FooterStyle:id=316,.parent=73"
      _StyleDefs(152) =   "Splits(0).Columns(27).EditorStyle:id=317,.parent=83"
      _StyleDefs(153) =   "Splits(0).Columns(28).Style:id=322,.parent=71,.locked=-1"
      _StyleDefs(154) =   "Splits(0).Columns(28).HeadingStyle:id=319,.parent=72,.alignment=2"
      _StyleDefs(155) =   "Splits(0).Columns(28).FooterStyle:id=320,.parent=73"
      _StyleDefs(156) =   "Splits(0).Columns(28).EditorStyle:id=321,.parent=83"
      _StyleDefs(157) =   "Splits(0).Columns(29).Style:id=326,.parent=71,.locked=-1"
      _StyleDefs(158) =   "Splits(0).Columns(29).HeadingStyle:id=323,.parent=72,.alignment=2"
      _StyleDefs(159) =   "Splits(0).Columns(29).FooterStyle:id=324,.parent=73"
      _StyleDefs(160) =   "Splits(0).Columns(29).EditorStyle:id=325,.parent=83"
      _StyleDefs(161) =   "Splits(0).Columns(30).Style:id=330,.parent=71,.locked=-1"
      _StyleDefs(162) =   "Splits(0).Columns(30).HeadingStyle:id=327,.parent=72,.alignment=2"
      _StyleDefs(163) =   "Splits(0).Columns(30).FooterStyle:id=328,.parent=73"
      _StyleDefs(164) =   "Splits(0).Columns(30).EditorStyle:id=329,.parent=83"
      _StyleDefs(165) =   "Splits(0).Columns(31).Style:id=334,.parent=71,.locked=-1"
      _StyleDefs(166) =   "Splits(0).Columns(31).HeadingStyle:id=331,.parent=72,.alignment=2"
      _StyleDefs(167) =   "Splits(0).Columns(31).FooterStyle:id=332,.parent=73"
      _StyleDefs(168) =   "Splits(0).Columns(31).EditorStyle:id=333,.parent=83"
      _StyleDefs(169) =   "Splits(0).Columns(32).Style:id=338,.parent=71,.locked=-1"
      _StyleDefs(170) =   "Splits(0).Columns(32).HeadingStyle:id=335,.parent=72,.alignment=2"
      _StyleDefs(171) =   "Splits(0).Columns(32).FooterStyle:id=336,.parent=73"
      _StyleDefs(172) =   "Splits(0).Columns(32).EditorStyle:id=337,.parent=83"
      _StyleDefs(173) =   "Splits(0).Columns(33).Style:id=342,.parent=71,.locked=-1"
      _StyleDefs(174) =   "Splits(0).Columns(33).HeadingStyle:id=339,.parent=72,.alignment=2"
      _StyleDefs(175) =   "Splits(0).Columns(33).FooterStyle:id=340,.parent=73"
      _StyleDefs(176) =   "Splits(0).Columns(33).EditorStyle:id=341,.parent=83"
      _StyleDefs(177) =   "Splits(0).Columns(34).Style:id=346,.parent=71,.locked=-1"
      _StyleDefs(178) =   "Splits(0).Columns(34).HeadingStyle:id=343,.parent=72,.alignment=2"
      _StyleDefs(179) =   "Splits(0).Columns(34).FooterStyle:id=344,.parent=73"
      _StyleDefs(180) =   "Splits(0).Columns(34).EditorStyle:id=345,.parent=83"
      _StyleDefs(181) =   "Splits(0).Columns(35).Style:id=350,.parent=71,.locked=-1"
      _StyleDefs(182) =   "Splits(0).Columns(35).HeadingStyle:id=347,.parent=72,.alignment=2"
      _StyleDefs(183) =   "Splits(0).Columns(35).FooterStyle:id=348,.parent=73"
      _StyleDefs(184) =   "Splits(0).Columns(35).EditorStyle:id=349,.parent=83"
      _StyleDefs(185) =   "Splits(0).Columns(36).Style:id=354,.parent=71,.locked=-1"
      _StyleDefs(186) =   "Splits(0).Columns(36).HeadingStyle:id=351,.parent=72,.alignment=2"
      _StyleDefs(187) =   "Splits(0).Columns(36).FooterStyle:id=352,.parent=73"
      _StyleDefs(188) =   "Splits(0).Columns(36).EditorStyle:id=353,.parent=83"
      _StyleDefs(189) =   "Splits(0).Columns(37).Style:id=358,.parent=71,.locked=-1"
      _StyleDefs(190) =   "Splits(0).Columns(37).HeadingStyle:id=355,.parent=72,.alignment=2"
      _StyleDefs(191) =   "Splits(0).Columns(37).FooterStyle:id=356,.parent=73"
      _StyleDefs(192) =   "Splits(0).Columns(37).EditorStyle:id=357,.parent=83"
      _StyleDefs(193) =   "Splits(0).Columns(38).Style:id=362,.parent=71,.locked=-1"
      _StyleDefs(194) =   "Splits(0).Columns(38).HeadingStyle:id=359,.parent=72,.alignment=2"
      _StyleDefs(195) =   "Splits(0).Columns(38).FooterStyle:id=360,.parent=73"
      _StyleDefs(196) =   "Splits(0).Columns(38).EditorStyle:id=361,.parent=83"
      _StyleDefs(197) =   "Splits(0).Columns(39).Style:id=366,.parent=71,.locked=-1"
      _StyleDefs(198) =   "Splits(0).Columns(39).HeadingStyle:id=363,.parent=72,.alignment=2"
      _StyleDefs(199) =   "Splits(0).Columns(39).FooterStyle:id=364,.parent=73"
      _StyleDefs(200) =   "Splits(0).Columns(39).EditorStyle:id=365,.parent=83"
      _StyleDefs(201) =   "Splits(0).Columns(40).Style:id=378,.parent=71"
      _StyleDefs(202) =   "Splits(0).Columns(40).HeadingStyle:id=375,.parent=72"
      _StyleDefs(203) =   "Splits(0).Columns(40).FooterStyle:id=376,.parent=73"
      _StyleDefs(204) =   "Splits(0).Columns(40).EditorStyle:id=377,.parent=83"
      _StyleDefs(205) =   "Splits(0).Columns(41).Style:id=370,.parent=71"
      _StyleDefs(206) =   "Splits(0).Columns(41).HeadingStyle:id=367,.parent=72"
      _StyleDefs(207) =   "Splits(0).Columns(41).FooterStyle:id=368,.parent=73"
      _StyleDefs(208) =   "Splits(0).Columns(41).EditorStyle:id=369,.parent=83"
      _StyleDefs(209) =   "Splits(0).Columns(42).Style:id=386,.parent=71"
      _StyleDefs(210) =   "Splits(0).Columns(42).HeadingStyle:id=383,.parent=72"
      _StyleDefs(211) =   "Splits(0).Columns(42).FooterStyle:id=384,.parent=73"
      _StyleDefs(212) =   "Splits(0).Columns(42).EditorStyle:id=385,.parent=83"
      _StyleDefs(213) =   "Splits(0).Columns(43).Style:id=394,.parent=71"
      _StyleDefs(214) =   "Splits(0).Columns(43).HeadingStyle:id=391,.parent=72"
      _StyleDefs(215) =   "Splits(0).Columns(43).FooterStyle:id=392,.parent=73"
      _StyleDefs(216) =   "Splits(0).Columns(43).EditorStyle:id=393,.parent=83"
      _StyleDefs(217) =   "Splits(1).Style:id=13,.parent=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(218) =   ":id=13,.strikethrough=0,.charset=0"
      _StyleDefs(219) =   ":id=13,.fontname=Verdana"
      _StyleDefs(220) =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(221) =   "Splits(1).HeadingStyle:id=14,.parent=2"
      _StyleDefs(222) =   "Splits(1).FooterStyle:id=15,.parent=3,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(223) =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(224) =   ":id=15,.fontname=Verdana"
      _StyleDefs(225) =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(226) =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(227) =   "Splits(1).EditorStyle:id=17,.parent=7"
      _StyleDefs(228) =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(229) =   "Splits(1).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
      _StyleDefs(230) =   "Splits(1).OddRowStyle:id=21,.parent=10"
      _StyleDefs(231) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(232) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(233) =   "Splits(1).Columns(0).Style:id=28,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(234) =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(235) =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(236) =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(237) =   "Splits(1).Columns(1).Style:id=32,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(238) =   "Splits(1).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(239) =   "Splits(1).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(240) =   "Splits(1).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(241) =   "Splits(1).Columns(2).Style:id=186,.parent=13,.locked=-1"
      _StyleDefs(242) =   "Splits(1).Columns(2).HeadingStyle:id=183,.parent=14,.alignment=2"
      _StyleDefs(243) =   "Splits(1).Columns(2).FooterStyle:id=184,.parent=15"
      _StyleDefs(244) =   "Splits(1).Columns(2).EditorStyle:id=185,.parent=17"
      _StyleDefs(245) =   "Splits(1).Columns(3).Style:id=182,.parent=13,.locked=-1"
      _StyleDefs(246) =   "Splits(1).Columns(3).HeadingStyle:id=179,.parent=14,.alignment=2"
      _StyleDefs(247) =   "Splits(1).Columns(3).FooterStyle:id=180,.parent=15"
      _StyleDefs(248) =   "Splits(1).Columns(3).EditorStyle:id=181,.parent=17"
      _StyleDefs(249) =   "Splits(1).Columns(4).Style:id=58,.parent=13,.locked=-1"
      _StyleDefs(250) =   "Splits(1).Columns(4).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(251) =   "Splits(1).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(252) =   "Splits(1).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(253) =   "Splits(1).Columns(5).Style:id=62,.parent=13,.locked=-1"
      _StyleDefs(254) =   "Splits(1).Columns(5).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(255) =   "Splits(1).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(256) =   "Splits(1).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(257) =   "Splits(1).Columns(6).Style:id=66,.parent=13,.locked=-1"
      _StyleDefs(258) =   "Splits(1).Columns(6).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(259) =   "Splits(1).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(260) =   "Splits(1).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(261) =   "Splits(1).Columns(7).Style:id=70,.parent=13,.locked=-1"
      _StyleDefs(262) =   "Splits(1).Columns(7).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(263) =   "Splits(1).Columns(7).FooterStyle:id=68,.parent=15"
      _StyleDefs(264) =   "Splits(1).Columns(7).EditorStyle:id=69,.parent=17"
      _StyleDefs(265) =   "Splits(1).Columns(8).Style:id=202,.parent=13,.locked=-1"
      _StyleDefs(266) =   "Splits(1).Columns(8).HeadingStyle:id=199,.parent=14,.alignment=2"
      _StyleDefs(267) =   "Splits(1).Columns(8).FooterStyle:id=200,.parent=15"
      _StyleDefs(268) =   "Splits(1).Columns(8).EditorStyle:id=201,.parent=17"
      _StyleDefs(269) =   "Splits(1).Columns(9).Style:id=82,.parent=13,.locked=-1"
      _StyleDefs(270) =   "Splits(1).Columns(9).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(271) =   "Splits(1).Columns(9).FooterStyle:id=80,.parent=15"
      _StyleDefs(272) =   "Splits(1).Columns(9).EditorStyle:id=81,.parent=17"
      _StyleDefs(273) =   "Splits(1).Columns(10).Style:id=178,.parent=13,.locked=-1"
      _StyleDefs(274) =   "Splits(1).Columns(10).HeadingStyle:id=175,.parent=14,.alignment=2"
      _StyleDefs(275) =   "Splits(1).Columns(10).FooterStyle:id=176,.parent=15"
      _StyleDefs(276) =   "Splits(1).Columns(10).EditorStyle:id=177,.parent=17"
      _StyleDefs(277) =   "Splits(1).Columns(11).Style:id=214,.parent=13,.locked=-1"
      _StyleDefs(278) =   "Splits(1).Columns(11).HeadingStyle:id=211,.parent=14,.alignment=2"
      _StyleDefs(279) =   "Splits(1).Columns(11).FooterStyle:id=212,.parent=15"
      _StyleDefs(280) =   "Splits(1).Columns(11).EditorStyle:id=213,.parent=17"
      _StyleDefs(281) =   "Splits(1).Columns(12).Style:id=90,.parent=13,.locked=-1"
      _StyleDefs(282) =   "Splits(1).Columns(12).HeadingStyle:id=87,.parent=14,.alignment=2"
      _StyleDefs(283) =   "Splits(1).Columns(12).FooterStyle:id=88,.parent=15"
      _StyleDefs(284) =   "Splits(1).Columns(12).EditorStyle:id=89,.parent=17"
      _StyleDefs(285) =   "Splits(1).Columns(13).Style:id=94,.parent=13,.locked=-1"
      _StyleDefs(286) =   "Splits(1).Columns(13).HeadingStyle:id=91,.parent=14,.alignment=2"
      _StyleDefs(287) =   "Splits(1).Columns(13).FooterStyle:id=92,.parent=15"
      _StyleDefs(288) =   "Splits(1).Columns(13).EditorStyle:id=93,.parent=17"
      _StyleDefs(289) =   "Splits(1).Columns(14).Style:id=78,.parent=13,.locked=-1"
      _StyleDefs(290) =   "Splits(1).Columns(14).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(291) =   "Splits(1).Columns(14).FooterStyle:id=76,.parent=15"
      _StyleDefs(292) =   "Splits(1).Columns(14).EditorStyle:id=77,.parent=17"
      _StyleDefs(293) =   "Splits(1).Columns(15).Style:id=98,.parent=13,.locked=-1"
      _StyleDefs(294) =   "Splits(1).Columns(15).HeadingStyle:id=95,.parent=14,.alignment=2"
      _StyleDefs(295) =   "Splits(1).Columns(15).FooterStyle:id=96,.parent=15"
      _StyleDefs(296) =   "Splits(1).Columns(15).EditorStyle:id=97,.parent=17"
      _StyleDefs(297) =   "Splits(1).Columns(16).Style:id=102,.parent=13,.locked=-1"
      _StyleDefs(298) =   "Splits(1).Columns(16).HeadingStyle:id=99,.parent=14,.alignment=2"
      _StyleDefs(299) =   "Splits(1).Columns(16).FooterStyle:id=100,.parent=15"
      _StyleDefs(300) =   "Splits(1).Columns(16).EditorStyle:id=101,.parent=17"
      _StyleDefs(301) =   "Splits(1).Columns(17).Style:id=106,.parent=13,.locked=-1"
      _StyleDefs(302) =   "Splits(1).Columns(17).HeadingStyle:id=103,.parent=14,.alignment=2"
      _StyleDefs(303) =   "Splits(1).Columns(17).FooterStyle:id=104,.parent=15"
      _StyleDefs(304) =   "Splits(1).Columns(17).EditorStyle:id=105,.parent=17"
      _StyleDefs(305) =   "Splits(1).Columns(18).Style:id=206,.parent=13,.locked=-1"
      _StyleDefs(306) =   "Splits(1).Columns(18).HeadingStyle:id=203,.parent=14,.alignment=2"
      _StyleDefs(307) =   "Splits(1).Columns(18).FooterStyle:id=204,.parent=15"
      _StyleDefs(308) =   "Splits(1).Columns(18).EditorStyle:id=205,.parent=17"
      _StyleDefs(309) =   "Splits(1).Columns(19).Style:id=110,.parent=13,.locked=-1"
      _StyleDefs(310) =   "Splits(1).Columns(19).HeadingStyle:id=107,.parent=14,.alignment=2"
      _StyleDefs(311) =   "Splits(1).Columns(19).FooterStyle:id=108,.parent=15"
      _StyleDefs(312) =   "Splits(1).Columns(19).EditorStyle:id=109,.parent=17"
      _StyleDefs(313) =   "Splits(1).Columns(20).Style:id=114,.parent=13,.locked=-1"
      _StyleDefs(314) =   "Splits(1).Columns(20).HeadingStyle:id=111,.parent=14,.alignment=2"
      _StyleDefs(315) =   "Splits(1).Columns(20).FooterStyle:id=112,.parent=15"
      _StyleDefs(316) =   "Splits(1).Columns(20).EditorStyle:id=113,.parent=17"
      _StyleDefs(317) =   "Splits(1).Columns(21).Style:id=210,.parent=13,.locked=-1"
      _StyleDefs(318) =   "Splits(1).Columns(21).HeadingStyle:id=207,.parent=14,.alignment=2"
      _StyleDefs(319) =   "Splits(1).Columns(21).FooterStyle:id=208,.parent=15"
      _StyleDefs(320) =   "Splits(1).Columns(21).EditorStyle:id=209,.parent=17"
      _StyleDefs(321) =   "Splits(1).Columns(22).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(322) =   "Splits(1).Columns(22).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(323) =   "Splits(1).Columns(22).FooterStyle:id=48,.parent=15"
      _StyleDefs(324) =   "Splits(1).Columns(22).EditorStyle:id=49,.parent=17"
      _StyleDefs(325) =   "Splits(1).Columns(23).Style:id=118,.parent=13,.locked=-1"
      _StyleDefs(326) =   "Splits(1).Columns(23).HeadingStyle:id=115,.parent=14,.alignment=2"
      _StyleDefs(327) =   "Splits(1).Columns(23).FooterStyle:id=116,.parent=15"
      _StyleDefs(328) =   "Splits(1).Columns(23).EditorStyle:id=117,.parent=17"
      _StyleDefs(329) =   "Splits(1).Columns(24).Style:id=122,.parent=13,.locked=-1"
      _StyleDefs(330) =   "Splits(1).Columns(24).HeadingStyle:id=119,.parent=14,.alignment=2"
      _StyleDefs(331) =   "Splits(1).Columns(24).FooterStyle:id=120,.parent=15"
      _StyleDefs(332) =   "Splits(1).Columns(24).EditorStyle:id=121,.parent=17"
      _StyleDefs(333) =   "Splits(1).Columns(25).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(334) =   "Splits(1).Columns(25).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(335) =   "Splits(1).Columns(25).FooterStyle:id=44,.parent=15"
      _StyleDefs(336) =   "Splits(1).Columns(25).EditorStyle:id=45,.parent=17"
      _StyleDefs(337) =   "Splits(1).Columns(26).Style:id=218,.parent=13,.locked=-1"
      _StyleDefs(338) =   "Splits(1).Columns(26).HeadingStyle:id=215,.parent=14"
      _StyleDefs(339) =   "Splits(1).Columns(26).FooterStyle:id=216,.parent=15"
      _StyleDefs(340) =   "Splits(1).Columns(26).EditorStyle:id=217,.parent=17"
      _StyleDefs(341) =   "Splits(1).Columns(27).Style:id=126,.parent=13,.locked=-1"
      _StyleDefs(342) =   "Splits(1).Columns(27).HeadingStyle:id=123,.parent=14,.alignment=2"
      _StyleDefs(343) =   "Splits(1).Columns(27).FooterStyle:id=124,.parent=15"
      _StyleDefs(344) =   "Splits(1).Columns(27).EditorStyle:id=125,.parent=17"
      _StyleDefs(345) =   "Splits(1).Columns(28).Style:id=130,.parent=13,.locked=-1"
      _StyleDefs(346) =   "Splits(1).Columns(28).HeadingStyle:id=127,.parent=14,.alignment=2"
      _StyleDefs(347) =   "Splits(1).Columns(28).FooterStyle:id=128,.parent=15"
      _StyleDefs(348) =   "Splits(1).Columns(28).EditorStyle:id=129,.parent=17"
      _StyleDefs(349) =   "Splits(1).Columns(29).Style:id=134,.parent=13,.locked=-1"
      _StyleDefs(350) =   "Splits(1).Columns(29).HeadingStyle:id=131,.parent=14,.alignment=2"
      _StyleDefs(351) =   "Splits(1).Columns(29).FooterStyle:id=132,.parent=15"
      _StyleDefs(352) =   "Splits(1).Columns(29).EditorStyle:id=133,.parent=17"
      _StyleDefs(353) =   "Splits(1).Columns(30).Style:id=54,.parent=13,.locked=-1"
      _StyleDefs(354) =   "Splits(1).Columns(30).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(355) =   "Splits(1).Columns(30).FooterStyle:id=52,.parent=15"
      _StyleDefs(356) =   "Splits(1).Columns(30).EditorStyle:id=53,.parent=17"
      _StyleDefs(357) =   "Splits(1).Columns(31).Style:id=142,.parent=13,.locked=-1"
      _StyleDefs(358) =   "Splits(1).Columns(31).HeadingStyle:id=139,.parent=14,.alignment=2"
      _StyleDefs(359) =   "Splits(1).Columns(31).FooterStyle:id=140,.parent=15"
      _StyleDefs(360) =   "Splits(1).Columns(31).EditorStyle:id=141,.parent=17"
      _StyleDefs(361) =   "Splits(1).Columns(32).Style:id=146,.parent=13,.locked=-1"
      _StyleDefs(362) =   "Splits(1).Columns(32).HeadingStyle:id=143,.parent=14,.alignment=2"
      _StyleDefs(363) =   "Splits(1).Columns(32).FooterStyle:id=144,.parent=15"
      _StyleDefs(364) =   "Splits(1).Columns(32).EditorStyle:id=145,.parent=17"
      _StyleDefs(365) =   "Splits(1).Columns(33).Style:id=150,.parent=13,.locked=-1"
      _StyleDefs(366) =   "Splits(1).Columns(33).HeadingStyle:id=147,.parent=14,.alignment=2"
      _StyleDefs(367) =   "Splits(1).Columns(33).FooterStyle:id=148,.parent=15"
      _StyleDefs(368) =   "Splits(1).Columns(33).EditorStyle:id=149,.parent=17"
      _StyleDefs(369) =   "Splits(1).Columns(34).Style:id=154,.parent=13,.locked=-1"
      _StyleDefs(370) =   "Splits(1).Columns(34).HeadingStyle:id=151,.parent=14,.alignment=2"
      _StyleDefs(371) =   "Splits(1).Columns(34).FooterStyle:id=152,.parent=15"
      _StyleDefs(372) =   "Splits(1).Columns(34).EditorStyle:id=153,.parent=17"
      _StyleDefs(373) =   "Splits(1).Columns(35).Style:id=158,.parent=13,.locked=-1"
      _StyleDefs(374) =   "Splits(1).Columns(35).HeadingStyle:id=155,.parent=14,.alignment=2"
      _StyleDefs(375) =   "Splits(1).Columns(35).FooterStyle:id=156,.parent=15"
      _StyleDefs(376) =   "Splits(1).Columns(35).EditorStyle:id=157,.parent=17"
      _StyleDefs(377) =   "Splits(1).Columns(36).Style:id=162,.parent=13,.locked=-1"
      _StyleDefs(378) =   "Splits(1).Columns(36).HeadingStyle:id=159,.parent=14,.alignment=2"
      _StyleDefs(379) =   "Splits(1).Columns(36).FooterStyle:id=160,.parent=15"
      _StyleDefs(380) =   "Splits(1).Columns(36).EditorStyle:id=161,.parent=17"
      _StyleDefs(381) =   "Splits(1).Columns(37).Style:id=166,.parent=13,.locked=-1"
      _StyleDefs(382) =   "Splits(1).Columns(37).HeadingStyle:id=163,.parent=14,.alignment=2"
      _StyleDefs(383) =   "Splits(1).Columns(37).FooterStyle:id=164,.parent=15"
      _StyleDefs(384) =   "Splits(1).Columns(37).EditorStyle:id=165,.parent=17"
      _StyleDefs(385) =   "Splits(1).Columns(38).Style:id=170,.parent=13,.locked=-1"
      _StyleDefs(386) =   "Splits(1).Columns(38).HeadingStyle:id=167,.parent=14,.alignment=2"
      _StyleDefs(387) =   "Splits(1).Columns(38).FooterStyle:id=168,.parent=15"
      _StyleDefs(388) =   "Splits(1).Columns(38).EditorStyle:id=169,.parent=17"
      _StyleDefs(389) =   "Splits(1).Columns(39).Style:id=174,.parent=13,.locked=-1"
      _StyleDefs(390) =   "Splits(1).Columns(39).HeadingStyle:id=171,.parent=14,.alignment=2"
      _StyleDefs(391) =   "Splits(1).Columns(39).FooterStyle:id=172,.parent=15"
      _StyleDefs(392) =   "Splits(1).Columns(39).EditorStyle:id=173,.parent=17"
      _StyleDefs(393) =   "Splits(1).Columns(40).Style:id=382,.parent=13"
      _StyleDefs(394) =   "Splits(1).Columns(40).HeadingStyle:id=379,.parent=14"
      _StyleDefs(395) =   "Splits(1).Columns(40).FooterStyle:id=380,.parent=15"
      _StyleDefs(396) =   "Splits(1).Columns(40).EditorStyle:id=381,.parent=17"
      _StyleDefs(397) =   "Splits(1).Columns(41).Style:id=374,.parent=13"
      _StyleDefs(398) =   "Splits(1).Columns(41).HeadingStyle:id=371,.parent=14"
      _StyleDefs(399) =   "Splits(1).Columns(41).FooterStyle:id=372,.parent=15"
      _StyleDefs(400) =   "Splits(1).Columns(41).EditorStyle:id=373,.parent=17"
      _StyleDefs(401) =   "Splits(1).Columns(42).Style:id=390,.parent=13,.alignment=2"
      _StyleDefs(402) =   "Splits(1).Columns(42).HeadingStyle:id=387,.parent=14"
      _StyleDefs(403) =   "Splits(1).Columns(42).FooterStyle:id=388,.parent=15"
      _StyleDefs(404) =   "Splits(1).Columns(42).EditorStyle:id=389,.parent=17"
      _StyleDefs(405) =   "Splits(1).Columns(43).Style:id=398,.parent=13"
      _StyleDefs(406) =   "Splits(1).Columns(43).HeadingStyle:id=395,.parent=14"
      _StyleDefs(407) =   "Splits(1).Columns(43).FooterStyle:id=396,.parent=15"
      _StyleDefs(408) =   "Splits(1).Columns(43).EditorStyle:id=397,.parent=17"
      _StyleDefs(409) =   "Named:id=33:Normal"
      _StyleDefs(410) =   ":id=33,.parent=0"
      _StyleDefs(411) =   "Named:id=34:Heading"
      _StyleDefs(412) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(413) =   ":id=34,.wraptext=-1"
      _StyleDefs(414) =   "Named:id=35:Footing"
      _StyleDefs(415) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(416) =   "Named:id=36:Selected"
      _StyleDefs(417) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H0&"
      _StyleDefs(418) =   "Named:id=37:Caption"
      _StyleDefs(419) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(420) =   "Named:id=38:HighlightRow"
      _StyleDefs(421) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(422) =   "Named:id=39:EvenRow"
      _StyleDefs(423) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(424) =   "Named:id=40:OddRow"
      _StyleDefs(425) =   ":id=40,.parent=33"
      _StyleDefs(426) =   "Named:id=41:RecordSelector"
      _StyleDefs(427) =   ":id=41,.parent=34"
      _StyleDefs(428) =   "Named:id=42:FilterBar"
      _StyleDefs(429) =   ":id=42,.parent=33"
      _StyleDefs(430) =   "Named:id=0:"
      _StyleDefs(431) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(432) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(433) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(434) =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(435) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(436) =   ":id=0,.fontname=MS Sans Serif"
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: SEARCH DATA CUSTOMER  ::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4155
      TabIndex        =   14
      Top             =   75
      Width           =   2985
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   60
      X2              =   11200
      Y1              =   8970
      Y2              =   8970
   End
End
Attribute VB_Name = "frmSIDSearchCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim isTambah    As Boolean
Dim rsTemp      As New ADODB.Recordset
Dim rsHlp       As New ADODB.Recordset
Dim rsGrid      As New ADODB.Recordset
Dim rsDuplikat  As New ADODB.Recordset
Dim rsMergeFix  As New ADODB.Recordset
Dim rsCust      As New ADODB.Recordset
Dim rsTempCust  As New ADODB.Recordset

Dim strWhere    As String
Dim strWhere2   As String
Dim outPut()    As String
Dim strTemp     As String
Dim strTemp2    As String
Dim strTemp3    As String
Dim strTemp4    As String
Dim strTemp5    As String
'Dim SQLConBP    As String

Private sSQL As String
Private MySearch    As String
Private isPassing   As Boolean
Private IsAct       As String
Private MyRs        As ADODB.Recordset
Private rsPel       As ADODB.Recordset

Private cari        As String
Private cari2       As String
Private cari3       As String
Private cari4       As String
Private cari5       As String
Private cari6       As String
Private carix       As String
Private cari2x      As String
Private cari3x      As String
Private cari4x      As String
Private cari5x      As String
Private cari6x      As String

Private daftarIDCust    As String
Private isLoad          As Boolean
Private idCustTerpilih  As String
Private idTempKolom     As String
Private kolomTerpilih   As String
Private dataDuplikat(20) As String
Private rsIdTerpilih    As ADODB.Recordset
Private idCustSplit     As String
Private kodeLgnLamaSplit    As String
Private namaCustomerSplit   As String
Private alamatSplit         As String
Private noKTPSplit          As String
Private noNPWPSplit         As String
Private NoHPSplit           As String
Private totalData       As Double

Dim cols As TrueOleDBGrid70.Columns

Public Property Let SetSearch(value As String)
    MySearch = value
End Property

Public Property Get IsPass() As Boolean
    IsPass = isPassing
End Property

Public Property Get rs() As ADODB.Recordset
    Set rs = MyRs
End Property

Public Property Get Action() As String
    Action = IsAct
End Property

Private Sub btnClear_Click()
    Call clearForm
End Sub

Private Sub cmdBuatIDBaru_Click()
    
    Dim Nama() As String
    If MsgBox("Apakah yakin untuk memisahkan ID Customer", vbYesNo) = vbYes Then
        Call clearForm
        If Log.monitoringScoring = 0 Then
            MsgBox "Project merge dan split single id belum diaktifkan", vbInformation
            
        Else
            With frmSIDMasterPelanggan
                .setIDCustSplit = kodeLgnLamaSplit
                .isSplitCustomer (True)
    '            .isSplitCustomer = True
                .isEdit False
                Nama() = Split(NulSID(namaCustomerSplit), " ")
                For i = 0 To UBound(Nama)
                    If Nama(i) <> "" And i = 0 Then
                        .txt(4) = NulSID(Nama(i))
                    ElseIf Nama(i) <> "" And i = 1 Then
                        .txt(5) = NulSID(Nama(i))
                    Else
                        .txt(6) = .txt(6) & NulSID(Nama(i))
                    End If
                Next
                .txt(3) = NulSID(noNPWPSplit)
                
                ''pecah notelepon
                .txt(14) = NulSID(NoHPSplit)
                
                .txt(8) = NulSID(alamatSplit)
                
                .show vbModal
            End With
        End If
        
        cmdBuatIDBaru.Visible = False
    End If
End Sub

Private Sub cmdDijual_Click()
    If NulSID(rsGrid!noRangka) = "" Then
        MsgBox "Tidak ada data no rangka untuk data ini.", vbInformation
        Exit Sub
    End If
    If NulSID(rsGrid!statusUnit) = NulSID("Active") Then
        Label30(9).Caption = "akan DIHAPUSKAN dari data Customer ini :"
        cmdHapus.Visible = True
        cmdAktif.Visible = False
    Else
        Label30(9).Caption = "akan DIKEMBALIKAN ke data Customer ini :"
        cmdHapus.Visible = False
        cmdAktif.Visible = True
    End If
    frameDijual.Visible = True
    labelNoRangka.Caption = NulSID(rsGrid!noRangka)
    labelNoPolisi.Caption = NulSID(rsGrid!NoPolisi)
    sSQL = "SELECT fmdl FROM [singleidori]..unitcustomer WITH (NOLOCK) WHERE norangka = '" & NulSID(rsGrid!noRangka) & "' "
    labelModel.Caption = objO.FetchData_String(sSQL)
    sSQL = "SELECT color FROM [singleidori]..unitcustomer WITH (NOLOCK) WHERE norangka = '" & NulSID(rsGrid!noRangka) & "' "
    labelWarna.Caption = objO.FetchData_String(sSQL)
    
    labelNamaCustomer.Caption = NulSID(rsGrid!namaCust)
    labelIDCustomer.Caption = NulSID(rsGrid!idcust)
    labelAlamat.Caption = NulSID(rsGrid!Alamat)
End Sub

Private Sub cmdDuplikasi_Click()
    frameDuplikasi.Top = 360
    frameDuplikasi.Visible = True
    gridDuplikat.DataSource = rsDuplikat
    gridDuplikat.ReBind
End Sub

Private Sub cmdHapus_Click()
    sSQL = "SELECT * FROM [singleidori]..DataDijual WITH (NOLOCK) WHERE idcust = '" & NulSID(labelIDCustomer.Caption) & "' AND norangka = '" & NulSID(labelNoRangka.Caption) & "' "
    Set rsHlp = objO.FetchData(sSQL)
    If Not rsHlp.EOF Then
        sSQL = "UPDATE [singleidori]..DataDijual SET dijual = 1,tglEntry = getdate(),userEntry = '" & gSA & "',PCEntry = '" & ComptName & "',AppEntry = 'SVC',DealerEntry = '" & Right(gDb, 2) & "' WHERE idcust = '" & NulSID(labelIDCustomer.Caption) & "' AND norangka = '" & NulSID(labelNoRangka.Caption) & "' "
    Else
        sSQL = "INSERT INTO [singleidori]..DataDijual (idCust,noRangka,dijual,tglEntry,userEntry,PCEntry,AppEntry,DealerEntry) VALUES ('" & NulSID(labelIDCustomer.Caption) & "','" & NulSID(labelNoRangka.Caption) & "',1,getdate(),'" & gSA & "','" & ComptName & "','SVC','" & Right(gDb, 2) & "')"
    End If
    Call objO.InsertData(sSQL)
    MsgBox "Data berhasil disimpan", vbInformation
    frameDijual.Visible = False
    Call btnCari_Click
End Sub

Private Sub cmdAktif_Click()
    sSQL = "UPDATE [singleidori]..DataDijual SET dijual = 0,tglEntry = getdate(),userEntry = '" & gSA & "',PCEntry = '" & ComptName & "',AppEntry = 'SVC',DealerEntry = '" & Right(gDb, 2) & "' WHERE idcust = '" & NulSID(labelIDCustomer.Caption) & "' AND norangka = '" & NulSID(labelNoRangka.Caption) & "' "
    Call objO.InsertData(sSQL)
    MsgBox "Data berhasil disimpan", vbInformation
    frameDijual.Visible = False
    Call btnCari_Click
End Sub


Private Sub cmdProses_Click()
    rsDuplikat.filter = "idMerge = " & True
    If Not rsDuplikat.EOF Then
        For i = 0 To 20
            dataDuplikat(i) = ""
        Next
        idCustTerpilih = rsDuplikat!idcust
        rsDuplikat.filter = ""
        
        sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & idCustTerpilih & "' "
        Set rsIdTerpilih = objO.FetchData(sSQL)
        If Not rsIdTerpilih.EOF Then
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihNama = " & True
            idTempKolom = ""
            If Not rsDuplikat.EOF Then
                idTempKolom = rsDuplikat!idcust
            Else
                idTempKolom = idCustTerpilih
            End If
            dataDuplikat(0) = objO.FetchData_String("SELECT NamaCust1 FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & idTempKolom & "'")
            dataDuplikat(1) = objO.FetchData_String("SELECT NamaCust2 FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & idTempKolom & "'")
            dataDuplikat(2) = objO.FetchData_String("SELECT NamaCust3 FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & idTempKolom & "'")
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihTglLahir = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(4) = Format(NulSID(rsDuplikat!tglLahir), "yyyy-mm-dd")
            Else
                dataDuplikat(4) = NulSID(Format(rsIdTerpilih!tglLahir, "yyyy-mm-dd"))
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihAgama = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(5) = NulSID(rsDuplikat!Agama)
            Else
                dataDuplikat(5) = NulSID(rsIdTerpilih!Agama)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihPekerjaan = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(6) = NulSID(rsDuplikat!pekerjaan)
            Else
                dataDuplikat(6) = NulSID(rsIdTerpilih!pekerjaan)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihHobi = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(7) = NulSID(rsDuplikat!hobi)
            Else
                dataDuplikat(7) = NulSID(rsIdTerpilih!hobi)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihOlahraga = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(8) = NulSID(rsDuplikat!Olahraga)
            Else
                dataDuplikat(8) = NulSID(rsIdTerpilih!Olahraga)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihAlamat = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(9) = NulSID(rsDuplikat!Alamat)
            Else
                dataDuplikat(9) = NulSID(rsIdTerpilih!Alamat)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihKota = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(10) = NulSID(rsDuplikat!Kota)
            Else
                dataDuplikat(10) = NulSID(rsIdTerpilih!Kota)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihPropinsi = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(11) = NulSID(rsDuplikat!Propinsi)
            Else
                dataDuplikat(11) = NulSID(rsIdTerpilih!Propinsi)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihKecamatan = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(12) = NulSID(rsDuplikat!kecamatan)
            Else
                dataDuplikat(12) = NulSID(rsIdTerpilih!kecamatan)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihKelurahan = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(13) = NulSID(rsDuplikat!kelurahan)
            Else
                dataDuplikat(13) = NulSID(rsIdTerpilih!kelurahan)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihHP = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(14) = NulSID(rsDuplikat!HP)
            Else
                dataDuplikat(14) = NulSID(rsIdTerpilih!hp1)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihNPWP = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(15) = NulSID(rsDuplikat!noNPWP)
            Else
                dataDuplikat(15) = NulSID(rsIdTerpilih!noNPWP)
            End If
            
            rsDuplikat.filter = ""
            rsDuplikat.filter = "pilihKTP = " & True
            If Not rsDuplikat.EOF Then
                dataDuplikat(16) = NulSID(rsDuplikat!noKTP)
            Else
                dataDuplikat(16) = NulSID(rsIdTerpilih!noKTP)
            End If
            
            Call createRsMergeFix
            With rsMergeFix
                .AddNew
                !idcust = NulSID(idTempKolom)
                !Nama = NulSID(dataDuplikat(0) & " " & dataDuplikat(1) & " " & dataDuplikat(2))
                !tglLahir = NulSID(dataDuplikat(4))
                !Agama = NulSID(dataDuplikat(5))
                !pekerjaan = NulSID(dataDuplikat(6))
                !hobi = NulSID(dataDuplikat(7))
                !Olahraga = NulSID(dataDuplikat(8))
                !Alamat = NulSID(dataDuplikat(9))
                !Kota = NulSID(dataDuplikat(10))
                !Propinsi = NulSID(dataDuplikat(11))
                !kecamatan = NulSID(dataDuplikat(12))
                !kelurahan = NulSID(dataDuplikat(13))
                !HP = NulSID(dataDuplikat(14))
                !noNPWP = NulSID(dataDuplikat(15))
                !noKTP = NulSID(dataDuplikat(16))
                .Update
            End With
            Call createHeaderMergeFix
            gridCustomerFix.DataSource = rsMergeFix
            frameUploadKTPMerge.Left = 0
            frameUploadKTPMerge.Visible = True
        Else
            rsDuplikat.filter = ""
            MsgBox "Data tidak ditemukan.", vbInformation
        End If
    Else
        rsDuplikat.filter = ""
        MsgBox "Pilih data yang diproses merge", vbInformation
    End If
End Sub

Private Sub cmdSimpanMerge_Click()
    
1        On Error GoTo cmdSimpanMerge_Click_Error
    Dim rsCek As ADODB.Recordset
    Dim namacustomer As String
    Dim isBegin As Boolean
2         isBegin = False
3         rsDuplikat.filter = "pilihMerge = " & True
4         If rsDuplikat.EOF Then
5       MsgBox "Tidak ada id yang akan di merge!", vbInformation
6       Exit Sub
7         End If
8         rsDuplikat.MoveFirst
    
9         If NulSID(txtUploadKTP.Text) <> "" Then
10      hasilKTP = simpanImageKTP(idTempKolom)
11      If hasilKTP = False Then
12          MsgBox "Gagal Upload KTP", vbInformation
13          Exit Sub
14      End If
15        Else
16      MsgBox "Foto KTP wajib diupload.", vbInformation
17      Exit Sub
18        End If
     
19        Call objHT.OpenTransactionBengkel
20        isBegin = True
    
21        Do While Not rsDuplikat.EOF
        ''simpan data usetemp
        ''------------------------------------
22      sSQL = "SELECT NAMACUST1,NAMACUST2,NAMACUST3 FROM [SingleIDOri].[dbo].newcustomer WITH (NOLOCK) WHERE IDCUST = '" & NulSID(rsDuplikat!idcust) & "'"
23      namacustomer = ""
24      Set rsCek = objHT.FetchDataBengkel(sSQL)
25      If Not rsCek.EOF Then
26          namacustomer = Left(NulSID(rsCek!namacust1), 1) & NulSID(rsCek!namacust2) & Right(NulSID(rsCek!namacust1), 1)
27      End If
        
28      If NulSID(rsDuplikat!idcust) = idTempKolom Then
29          GoTo lanjutLooping
30      End If
        
31      sSQL = "INSERT INTO [SingleIDOri].[dbo].[UseTemp]([IDCust],[UserId],[Nama],[IDMerge],[TglEntry],[Valid],[TglValid],[UnMerge],[TglUnMerge]"
32      sSQL = sSQL & " ,[SyncUnit],[TglSync],[userpc],[userapp],[userdealer],[pcunmerge],[appunmerge],[dealerunmerge]) VALUES ( "
33      sSQL = sSQL & " '" & NulSID(rsDuplikat!idcust) & "','" & gSA & "','" & namacustomer & "','" & idTempKolom & "',getdate(),1,getdate(),0,'1900-01-01'"
34      sSQL = sSQL & " ,1,getdate(),'" & ComptName & "','SVC','" & gDb & "','','','')"
35      Call objHT.InsertData_Bengkel(sSQL)
        
        ''bridging temp
36      sSQL = "INSERT INTO [SingleIDOri].[dbo].[BridgingTemp] (KodeDealer,kodeLgn,idcust,TglEntry,UserEntry,IDMerge) "
37      sSQL = sSQL & " SELECT KodeDealer,kodeLgn,idcust,getdate(),'" & gSA & "','" & idTempKolom & "' FROM [SingleIDOri].[dbo].[Bridging] WITH (NOLOCK) WHERE idcust = '" & NulSID(rsDuplikat!idcust) & "'"
38      Call objHT.InsertData_Bengkel(sSQL)
        
        ''bridging
39      sSQL = "UPDATE [SingleIDOri].[dbo].BRIDGING SET idcust = '" & idTempKolom & "' WHERE idcust = '" & NulSID(rsDuplikat!idcust) & "'"
40      Call objHT.InsertData_Bengkel(sSQL)

        ''catatan babah cek atas kodelgn atas kode idcust jika idcust belum ada di kodelgn akan di insert revkodeuser tgl
41      sSQL = "SELECT idcust FROM [SingleIDOri]..BRIDGING WITH (NOLOCK) WHERE kodeLgn = '" & NulSID(rsDuplikat!idcust) & "'"
42      Set rsCek = objHT.FetchDataBengkel(sSQL)
43      If rsCek.EOF Then
44          sSQL = "INSERT INTO SINGLEIDORI..Bridging(KodeDealer,kodeLgn,IDCust,TglEntry,UserEntry) VALUES ('" & Right(gDb, 2) & "','" & NulSID(rsDuplikat!idcust) & "','" & NulSID(idTempKolom) & "',getdate(),'REV" & gSA & "')"
45          Call objHT.InsertData_Bengkel(sSQL)
46      End If
        ''hapusmerge
47      sSQL = "INSERT INTO [SingleIDOri].[dbo].[HapusMerge] ([IDCust],[TypeCustomer],[NoKTP],[NoNPWP],[Title],[NamaCust1],[NamaCust2],[NamaCust3],[TglLahir],[Agama],[Pekerjaan]"
48      sSQL = sSQL & " ,[Propinsi],[Kota],[Kecamatan],[Kelurahan],[KodePos],[Alamat],[RT],[RW],[IdTelp],[Telp],[HP1],[HP2],[Email],[Hobi]"
49      sSQL = sSQL & " ,[Olahraga],[Suku],[Preferensi],[FPGab],[NonPPn],[TopVhc],[BatasVhc],[TopGrp],[BatasGrp],[TopBp],[BatasBp]"
50      sSQL = sSQL & " ,[TopPart],[BatasPart],[TglEntry],[UserEntry],[Affiliasi],[JenisUsaha],[JenisKelamin],[TempatLahir],[NoFax]"
51      sSQL = sSQL & " ,[PinBB],[TitlesDM],[NamaDM1],[NamaDM2],[NamaDM3],[NoKTPDM],[JenisKelaminDM],[TempatLahirDM],[TglLahirDM]"
52      sSQL = sSQL & " ,[AlamatDM],[PropinsiDM],[KotaDM],[KecamatanDM],[KelurahanDM],[CabangDM],[KodePosDM],[noHP1DM],[noHP2DM]"
53      sSQL = sSQL & " ,[idTelpDM],[telpDM],[NPWPDM],[PekerjaanDM],[HobiDM],[OlahragaDM],[AgamaDM],[NoFaxDM],[EmailDM],[PinBBDM]"
54      sSQL = sSQL & " ,[titlesOwner],[Nama1Own],[Nama2Own],[Nama3Own],[NoKTPOWn],[JenisKelaminOwn],[TempatLahirOwn],[AlamatOwn]"
55      sSQL = sSQL & " ,[PropinsiOwn],[KotaOwn],[KecamatanOwn],[CabangOwn],[KelurahanOwn],[KodePosOwn],[noHP1Own],[noHP2Own]"
56      sSQL = sSQL & " ,[idTelpOwn],[telpOwn],[NPWPOwn],[PekerjaanOwn],[HobiOwn],[AgamaOwn],[NoFaxOwn],[EmailOwn],[PinBBOwn]"
57      sSQL = sSQL & " ,[NamaPIC],[TelpPIC],[Kewarganegaraan],[UserUpdate],[TglUpdate],[DealerUpdate],[AppUpdate],[IDMerge])"
58      sSQL = sSQL & " SELECT [IDCust],[TypeCustomer],[NoKTP],[NoNPWP],[Title],[NamaCust1],[NamaCust2],[NamaCust3],[TglLahir],[Agama],[Pekerjaan]"
59      sSQL = sSQL & " ,[Propinsi],[Kota],[Kecamatan],[Kelurahan],[KodePos],[Alamat],[RT],[RW],[IdTelp],[Telp],[HP1],[HP2],[Email],[Hobi]"
60      sSQL = sSQL & " ,[Olahraga],[Suku],[Preferensi],[FPGab],[NonPPn],[TopVhc],[BatasVhc],[TopGrp],[BatasGrp],[TopBp],[BatasBp]"
61      sSQL = sSQL & " ,[TopPart],[BatasPart],[TglEntry],[UserEntry],[Affiliasi],[JenisUsaha],[JenisKelamin],[TempatLahir],[NoFax]"
62      sSQL = sSQL & " ,[PinBB],[TitlesDM],[NamaDM1],[NamaDM2],[NamaDM3],[NoKTPDM],[JenisKelaminDM],[TempatLahirDM],[TglLahirDM]"
63      sSQL = sSQL & " ,[AlamatDM],[PropinsiDM],[KotaDM],[KecamatanDM],[KelurahanDM],[CabangDM],[KodePosDM],[noHP1DM],[noHP2DM]"
64      sSQL = sSQL & " ,[idTelpDM],[telpDM],[NPWPDM],[PekerjaanDM],[HobiDM],[OlahragaDM],[AgamaDM],[NoFaxDM],[EmailDM],[PinBBDM]"
65      sSQL = sSQL & " ,[titlesOwner],[Nama1Own],[Nama2Own],[Nama3Own],[NoKTPOWn],[JenisKelaminOwn],[TempatLahirOwn],[AlamatOwn]"
66      sSQL = sSQL & " ,[PropinsiOwn],[KotaOwn],[KecamatanOwn],[CabangOwn],[KelurahanOwn],[KodePosOwn],[noHP1Own],[noHP2Own]"
67      sSQL = sSQL & " ,[idTelpOwn],[telpOwn],[NPWPOwn],[PekerjaanOwn],[HobiOwn],[AgamaOwn],[NoFaxOwn],[EmailOwn],[PinBBOwn]"
68      sSQL = sSQL & " ,[NamaPIC],[TelpPIC],[Kewarganegaraan],[UserUpdate],[TglUpdate],[DealerUpdate],[AppUpdate],'" & idTempKolom & "' "
69      sSQL = sSQL & " FROM [SingleIDOri].[dbo].[NewCustomer] WITH (NOLOCK) WHERE idcust = '" & NulSID(rsDuplikat!idcust) & "'"
70      Call objHT.InsertData_Bengkel(sSQL)
        
        ''-- unitmerge
71      sSQL = "INSERT INTO [SingleIDOri].[dbo].[UnitMerge] ([IDMerge],[NoRangka],[IDCustSales]"
72      sSQL = sSQL & " ,[IDCustService],[IDCustSTNK],[NoKTPSTNK],[NoNPWPSTNK],[TitlesSTNK],[NamaCust1STNK],[NamaCust2STNK],[NamaCust3STNK]"
73      sSQL = sSQL & " ,[TglLahirSTNK],[AgamaSTNK],[PekerjaanSTNK],[PropinsiSTNK],[KotaSTNK],[KecamatanSTNK],[KelurahanSTNK],[KodePosSTNK]"
74      sSQL = sSQL & " ,[AlamatSTNK],[RTSTNK],[RWSTNK],[IdTelpSTNK],[TelpSTNK],[HP1STNK],[HP2STNK],[EmailSTNK],[HobiSTNK],[OlahragaSTNK]"
75      sSQL = sSQL & " ,[SukuSTNK],[PreferensiSTNK],[IDCustSL],[NoKTPSL],[NoNPWPSL],[TitlesSL],[NamaCust1SL],[NamaCust2SL],[NamaCust3SL]"
76      sSQL = sSQL & " ,[TglLahirSL],[AgamaSL],[PekerjaanSL],[PropinsiSL],[KotaSL],[KecamatanSL],[KelurahanSL],[KodePosSL],[AlamatSL],[RTSL]"
77      sSQL = sSQL & " ,[RWSL],[IdTelpSL],[TelpSL],[HP1SL],[HP2SL],[EmailSL],[HobiSL],[OlahragaSL],[SukuSL],[PreferensiSL],[IDCustSV]"
78      sSQL = sSQL & " ,[NoKTPSV],[NoNPWPSV],[TitlesSV],[NamaCust1SV],[NamaCust2SV],[NamaCust3SV],[TglLahirSV],[AgamaSV],[PekerjaanSV]"
79      sSQL = sSQL & " ,[PropinsiSV],[KotaSV],[KecamatanSV],[KelurahanSV],[KodePosSV],[AlamatSV],[RTSV],[RWSV],[IdTelpSV],[TelpSV]"
80      sSQL = sSQL & " ,[HP1SV],[HP2SV],[EmailSV],[HobiSV],[OlahragaSV],[SukuSV],[PreferensiSV],[FaxSV],[TglEntry],[UserEntry],[TglMerge]"
81      sSQL = sSQL & " ,[UserMerge],[UnMerge],[TglUnMerge],[pcmerge],[appmerge],[dealermerge],[pcunmerge],[appunmerge],[dealerunmerge])"
82      sSQL = sSQL & " SELECT '" & idTempKolom & "',[NoRangka],[IDCustSales]"
83      sSQL = sSQL & " ,[IDCustService],[IDCustSTNK],[NoKTPSTNK],[NoNPWPSTNK],[TitlesSTNK],[NamaCust1STNK],[NamaCust2STNK],[NamaCust3STNK]"
84      sSQL = sSQL & " ,[TglLahirSTNK],[AgamaSTNK],[PekerjaanSTNK],[PropinsiSTNK],[KotaSTNK],[KecamatanSTNK],[KelurahanSTNK],[KodePosSTNK]"
85      sSQL = sSQL & " ,[AlamatSTNK],[RTSTNK],[RWSTNK],[IdTelpSTNK],[TelpSTNK],[HP1STNK],[HP2STNK],[EmailSTNK],[HobiSTNK],[OlahragaSTNK]"
86      sSQL = sSQL & " ,[SukuSTNK],[PreferensiSTNK],[IDCustSL],[NoKTPSL],[NoNPWPSL],[TitlesSL],[NamaCust1SL],[NamaCust2SL],[NamaCust3SL]"
87      sSQL = sSQL & " ,[TglLahirSL],[AgamaSL],[PekerjaanSL],[PropinsiSL],[KotaSL],[KecamatanSL],[KelurahanSL],[KodePosSL],[AlamatSL],[RTSL]"
88      sSQL = sSQL & " ,[RWSL],[IdTelpSL],[TelpSL],[HP1SL],[HP2SL],[EmailSL],[HobiSL],[OlahragaSL],[SukuSL],[PreferensiSL],[IDCustSV]"
89      sSQL = sSQL & " ,[NoKTPSV],[NoNPWPSV],[TitlesSV],[NamaCust1SV],[NamaCust2SV],[NamaCust3SV],[TglLahirSV],[AgamaSV],[PekerjaanSV]"
90      sSQL = sSQL & " ,[PropinsiSV],[KotaSV],[KecamatanSV],[KelurahanSV],[KodePosSV],[AlamatSV],[RTSV],[RWSV],[IdTelpSV],[TelpSV]"
91      sSQL = sSQL & " ,[HP1SV],[HP2SV],[EmailSV],[HobiSV],[OlahragaSV],[SukuSV],[PreferensiSV],[FaxSV],[TglEntry],[UserEntry],GETDATE()"
92      sSQL = sSQL & " ,'" & gSA & "',0,'1900-01-01','" & ComptName & "','SVC','" & gDb & "','','',''"
93      sSQL = sSQL & " FROM [SingleIDOri].[dbo].[UnitCustomer] WITH (NOLOCK) WHERE IDCustSales = '" & NulSID(rsDuplikat!idcust) & "' "
94      sSQL = sSQL & " UNION "
95      sSQL = sSQL & " SELECT '" & idTempKolom & "',[NoRangka],[IDCustSales]"
96      sSQL = sSQL & " ,[IDCustService],[IDCustSTNK],[NoKTPSTNK],[NoNPWPSTNK],[TitlesSTNK],[NamaCust1STNK],[NamaCust2STNK],[NamaCust3STNK]"
97      sSQL = sSQL & " ,[TglLahirSTNK],[AgamaSTNK],[PekerjaanSTNK],[PropinsiSTNK],[KotaSTNK],[KecamatanSTNK],[KelurahanSTNK],[KodePosSTNK]"
98      sSQL = sSQL & " ,[AlamatSTNK],[RTSTNK],[RWSTNK],[IdTelpSTNK],[TelpSTNK],[HP1STNK],[HP2STNK],[EmailSTNK],[HobiSTNK],[OlahragaSTNK]"
99      sSQL = sSQL & " ,[SukuSTNK],[PreferensiSTNK],[IDCustSL],[NoKTPSL],[NoNPWPSL],[TitlesSL],[NamaCust1SL],[NamaCust2SL],[NamaCust3SL]"
100     sSQL = sSQL & " ,[TglLahirSL],[AgamaSL],[PekerjaanSL],[PropinsiSL],[KotaSL],[KecamatanSL],[KelurahanSL],[KodePosSL],[AlamatSL],[RTSL]"
101     sSQL = sSQL & " ,[RWSL],[IdTelpSL],[TelpSL],[HP1SL],[HP2SL],[EmailSL],[HobiSL],[OlahragaSL],[SukuSL],[PreferensiSL],[IDCustSV]"
102     sSQL = sSQL & " ,[NoKTPSV],[NoNPWPSV],[TitlesSV],[NamaCust1SV],[NamaCust2SV],[NamaCust3SV],[TglLahirSV],[AgamaSV],[PekerjaanSV]"
103     sSQL = sSQL & " ,[PropinsiSV],[KotaSV],[KecamatanSV],[KelurahanSV],[KodePosSV],[AlamatSV],[RTSV],[RWSV],[IdTelpSV],[TelpSV]"
104     sSQL = sSQL & " ,[HP1SV],[HP2SV],[EmailSV],[HobiSV],[OlahragaSV],[SukuSV],[PreferensiSV],[FaxSV],[TglEntry],[UserEntry],GETDATE()"
105     sSQL = sSQL & " ,'" & gSA & "',0,'1900-01-01','" & ComptName & "','SVC','" & gDb & "','','',''"
106     sSQL = sSQL & " FROM [SingleIDOri].[dbo].[UnitCustomer] WITH (NOLOCK) WHERE IDCustService = '" & NulSID(rsDuplikat!idcust) & "' "
107     sSQL = sSQL & " UNION "
108     sSQL = sSQL & " SELECT '" & idTempKolom & "',[NoRangka],[IDCustSales]"
109     sSQL = sSQL & " ,[IDCustService],[IDCustSTNK],[NoKTPSTNK],[NoNPWPSTNK],[TitlesSTNK],[NamaCust1STNK],[NamaCust2STNK],[NamaCust3STNK]"
110     sSQL = sSQL & " ,[TglLahirSTNK],[AgamaSTNK],[PekerjaanSTNK],[PropinsiSTNK],[KotaSTNK],[KecamatanSTNK],[KelurahanSTNK],[KodePosSTNK]"
111     sSQL = sSQL & " ,[AlamatSTNK],[RTSTNK],[RWSTNK],[IdTelpSTNK],[TelpSTNK],[HP1STNK],[HP2STNK],[EmailSTNK],[HobiSTNK],[OlahragaSTNK]"
112     sSQL = sSQL & " ,[SukuSTNK],[PreferensiSTNK],[IDCustSL],[NoKTPSL],[NoNPWPSL],[TitlesSL],[NamaCust1SL],[NamaCust2SL],[NamaCust3SL]"
113     sSQL = sSQL & " ,[TglLahirSL],[AgamaSL],[PekerjaanSL],[PropinsiSL],[KotaSL],[KecamatanSL],[KelurahanSL],[KodePosSL],[AlamatSL],[RTSL]"
114     sSQL = sSQL & " ,[RWSL],[IdTelpSL],[TelpSL],[HP1SL],[HP2SL],[EmailSL],[HobiSL],[OlahragaSL],[SukuSL],[PreferensiSL],[IDCustSV]"
115     sSQL = sSQL & " ,[NoKTPSV],[NoNPWPSV],[TitlesSV],[NamaCust1SV],[NamaCust2SV],[NamaCust3SV],[TglLahirSV],[AgamaSV],[PekerjaanSV]"
116     sSQL = sSQL & " ,[PropinsiSV],[KotaSV],[KecamatanSV],[KelurahanSV],[KodePosSV],[AlamatSV],[RTSV],[RWSV],[IdTelpSV],[TelpSV]"
117     sSQL = sSQL & " ,[HP1SV],[HP2SV],[EmailSV],[HobiSV],[OlahragaSV],[SukuSV],[PreferensiSV],[FaxSV],[TglEntry],[UserEntry],GETDATE()"
118     sSQL = sSQL & " ,'" & gSA & "',0,'1900-01-01','" & ComptName & "','SVC','" & gDb & "','','',''"
119     sSQL = sSQL & " FROM [SingleIDOri].[dbo].[UnitCustomer] WITH (NOLOCK) WHERE IDCustSTNK = '" & NulSID(rsDuplikat!idcust) & "' "
120     Call objHT.InsertData_Bengkel(sSQL)
        
        '' update unitcustomer
121     sSQL = "UPDATE [SingleIDOri].[dbo].[UnitCustomer] SET IDCustSales = '" & idTempKolom & "' WHERE IDCustSales = '" & NulSID(rsDuplikat!idcust) & "' "
122     Call objHT.InsertData_Bengkel(sSQL)
123     sSQL = "UPDATE [SingleIDOri].[dbo].[UnitCustomer] SET IDCustService = '" & idTempKolom & "' WHERE IDCustService = '" & NulSID(rsDuplikat!idcust) & "' "
124     Call objHT.InsertData_Bengkel(sSQL)
125     sSQL = "UPDATE [SingleIDOri].[dbo].[UnitCustomer] SET IDCustSTNK = '" & idTempKolom & "' WHERE IDCustSTNK = '" & NulSID(rsDuplikat!idcust) & "' "
126     Call objHT.InsertData_Bengkel(sSQL)
        
        ''hapus newcustomer
127     sSQL = "DELETE FROM [SingleIDOri].[dbo].[NewCustomer] WHERE idcust = '" & NulSID(rsDuplikat!idcust) & "' "
128     Call objHT.InsertData_Bengkel(sSQL)
        
lanjutLooping:
129     rsDuplikat.MoveNext
130       Loop
    
    ''update id terpilih dengan data yang dipilih user
    ''----------------------------------------------------------
131       sSQL = "UPDATE [SingleIDOri].[dbo].[NewCustomer] SET NamaCust1 = '" & NulSID(dataDuplikat(0)) & "',NamaCust2 = '" & NulSID(dataDuplikat(1)) & "',NamaCust3 = '" & NulSID(dataDuplikat(2)) & "',tglLahir = '" & NulSID(dataDuplikat(4)) & "',Agama = '" & dataDuplikat(5) & "',"
132       sSQL = sSQL & " pekerjaan = '" & NulSID(dataDuplikat(6)) & "',Propinsi = '" & NulSID(dataDuplikat(11)) & "',Kota = '" & NulSID(dataDuplikat(10)) & "',Kecamatan = '" & NulSID(dataDuplikat(12)) & "',Kelurahan = '" & NulSID(dataDuplikat(13)) & "',Alamat = '" & NulSID(dataDuplikat(9)) & "',"
133       sSQL = sSQL & " HP1 = '" & NulSID(dataDuplikat(14)) & "',noKTP = '" & NulSID(dataDuplikat(16)) & "',NoNPWP = '" & NulSID(dataDuplikat(15)) & "',Hobi = '" & NulSID(dataDuplikat(7)) & "',Olahraga = '" & NulSID(dataDuplikat(8)) & "' WHERE idcust = '" & idTempKolom & "'"
134       Call objHT.InsertData_Bengkel(sSQL)
    
135       Call objHT.CloseTransactionBengkel
136       MsgBox "Data berhasil disimpan", vbInformation
    
137       cmdDuplikasi.Visible = False
138       frameDuplikasi.Top = 36000
139       frameUploadKTPMerge.Left = 30000
140       Call btnCari_Click
141      On Error GoTo 0
142      Exit Sub

cmdSimpanMerge_Click_Error:
143       If isBegin Then
144     Call objHT.CloseTransactionBengkel_err
145       End If
146       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpanMerge_Click of Form frmSIDSearchCust"


End Sub

Private Sub cmdSplit_Click()
    cmdSplit.Visible = False
    cmdBuatIDBaru.Visible = True
End Sub

Private Sub cmdTutup_Click()
    frameDijual.Visible = False
    
End Sub

Private Sub cmdTutupDuplikasi_Click()
    frameDuplikasi.Visible = False
    frameUploadKTPMerge.Visible = False
End Sub

Private Sub cmdUploadKTP_Click()
    Dim PictBmp As String
    CommonDialog1.FileName = ""
    CommonDialog1.filter = "Pictures (*.jpg;*.jpeg;)|*.jpg;*.jpeg;"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.ShowOpen
    PictBmp = CommonDialog1.FileName
    txtUploadKTP.Text = PictBmp
    If PictBmp = "" Then Exit Sub
    If FileLen(PictBmp) > 500000 Then
        txtUploadKTP.Text = ""
        MsgBox "Ukuran foto terlalu besar, Ukuran foto maksimal 500 KB"
        Close #1
        Exit Sub
    End If
    If CommonDialog1.FileName = "" Then Exit Sub
     
    Image1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Form_Activate()
    If NulSID(gBantuan) <> "" Then
        If NulSID(gBantuan) = "" Or NulSID(gBantuan) = "BOOKING" Then
            Text(4) = ""
        Else
            Text(4) = NulSID(gBantuan)
        End If
    End If
End Sub

Private Sub Form_Load()
    isLoad = False
    clearForm
    
End Sub

Private Sub clearForm()
    Dim i As Integer
    For i = 0 To 4
        Text(i).Text = ""
    Next
    Grid.Columns(42).Visible = False
    cmdDuplikasi.Visible = False
    frameDuplikasi.Top = 36000
    frameUploadKTPMerge.Left = 30000
     Text(9).Text = ""
     Text(10).Text = ""
        Text(8).Text = ""
        Text(7).Text = ""
        Text(6).Text = ""
        tglLahir.value = "" '"__/__/____"
        Text(5).Text = ""
    tglLahirx.value = ""
    btnTambah.Enabled = False
    ARProgressBar1.value = 0
    ARProgressBar1.Visible = False
    cmdDijual.Visible = False
    Call createRsGrid
End Sub

Private Sub btnCari_Click()
    Dim boleh1 As Boolean, boleh2 As Boolean
    Dim isCariID    As Boolean
    isCariID = False
    boleh1 = True: boleh2 = True
    
    If NulSID(Text(0)) = "" And NulSID(Text(1)) = "" Then
        boleh1 = False
    End If
    If NulSID(Text(2)) = "" And NulSID(Text(3)) = "" And NulSID(Text(4)) = "" And NulSID(tglLahir.value) = "" And NulSID(Text(10)) = "" Then
        boleh2 = False
    End If
    
    If NulSID(Text(4)) <> "" Then
        isCariID = True
    End If
    
    If boleh1 = False And boleh2 = False Then
        MsgBox "Data filter kosong, tidak dapat proses data", vbInformation
        Exit Sub
    Else
        btnCari.Enabled = False
        frameLoading.Visible = True
        progBar.Max = 5
        progBar.value = 0
        DoEvents
        daftarIDCust = ""
        totalData = 0
        Call PencarianData
        If Log.monitoringScoring = 1 Then
            
            If totalData > 500 Then
                If MsgBox("Poses pencarian data duplikat dari " & Format(totalData, "#,#") & " membutuhkan waktu yg lama." & vbCrLf & "Apakah tetap melanjutkan?", vbYesNo) = vbNo Then
                    GoTo lanjutSetelahDuplikat
                End If
            End If
            Call PencarianDataDuplikat
lanjutSetelahDuplikat:
            If isCariID Then
                If PencarianDataSplit Then
                    Grid.Columns(42).Visible = True
                End If
            End If
        End If
        frameLoading.Visible = False
        Grid.SetFocus
        btnCari.Enabled = True
    End If
End Sub

Private Function PencarianDataSplit() As Boolean
     
    Dim hasil           As Boolean
    Dim Nama            As String
    hasil = False
    
    If rsGrid.RecordCount > 1 Then
        If NulSID(rsGrid!idcust) = NulSID(Text(4)) Then
            rsGrid.MoveFirst
            Do While Not rsGrid.EOF
                If Nama = "" Then
                    Nama = NulSID(rsGrid!namaCust)
                End If
                If NulSID(rsGrid!namaCust) <> Nama Then
                    hasil = True
                End If
                
                rsGrid.MoveNext
            Loop
            rsGrid.MoveFirst
        End If
    End If
    
    PencarianDataSplit = hasil
    Exit Function
    
End Function

Private Sub PencarianDataDuplikat()
          Dim totalDuplikat   As Integer
          Dim rsClone         As New ADODB.Recordset
          Dim isNama          As Boolean
          Dim isTglLahir      As Boolean
          Dim isAlamat        As Boolean
          Dim isNoHP          As Boolean
          Dim isKTP           As Boolean
1         If daftarIDCust <> "" Then
2             daftarIDCust = Left(daftarIDCust, Len(daftarIDCust) - 1)
3         End If
          
4         Set rsTemp = Nothing
5         frameLoading.Visible = True
6         progBar.Max = totalData + 1
7         progBar.value = 0
8         lblKeteranganLoading.Caption = "PENCARIAN DATA SINGLE ID DUPLIKASI......."
9         DoEvents
10        Grid.Enabled = False
11        If daftarIDCust <> "" Then
12            sSQL = "select DISTINCT IDCust,NamaCust1, NamaCust2,ltrim(rtrim(NamaCust1+ ' ' + NamaCust2 + ' ' + NamaCust3)) as nama,TglLahir,Agama,Pekerjaan,Hobi,Olahraga,Alamat,Kota,Propinsi,Kecamatan,Kelurahan,HP1,NoKTP,NoNPWP from SingleIDOri..NewCustomer WITH (NOLOCK) WHERE IDCUST IN (" & daftarIDCust & ") order by nama,TglLahir,Kota,Propinsi,Kecamatan,Kelurahan,HP1,NoKTP,NoNPWP "
13            Set rsTemp = objSID.FetchDataSID(sSQL)
14            Call createRsDuplikat
15            If Not rsTemp.EOF Then
16                rsTemp.MoveFirst
17                Set rsClone = rsTemp.Clone
18                progBar.Max = rsTemp.RecordCount
19                progBar.value = 0
20                Do While Not rsTemp.EOF
21                    totalDuplikat = 0
22                    isNama = False
23                    isTglLahir = False
24                    isAlamat = False
25                    isNoHP = False
26                    isKTP = False
27                    rsClone.MoveFirst
28                    Do While Not rsClone.EOF
                          ''If NulSID(rsTemp!idcust) <> NulSID(rsClone!idcust) And NulSID(rsTemp!Nama) = NulSID(rsClone!Nama) And isNama = False Then
29                        If NulSID(rsTemp!idcust) <> NulSID(rsClone!idcust) And (NulSID(rsTemp!namacust1) = NulSID(rsClone!namacust1) Or NulSID(rsTemp!namacust2) = NulSID(rsClone!namacust2)) And isNama = False Then
30                            totalDuplikat = totalDuplikat + 1
31                            isNama = True
32                        End If
33                        If NulSID(rsTemp!idcust) <> NulSID(rsClone!idcust) And NulSID(rsTemp!tglLahir) = NulSID(rsClone!tglLahir) And isTglLahir = False Then
34                            totalDuplikat = totalDuplikat + 1
35                            isTglLahir = True
36                        End If
37                        If NulSID(rsTemp!idcust) <> NulSID(rsClone!idcust) And NulSID(rsTemp!Kota) = NulSID(rsClone!Kota) And NulSID(rsTemp!Propinsi) = NulSID(rsClone!Propinsi) And NulSID(rsTemp!kecamatan) = NulSID(rsClone!kecamatan) And NulSID(rsTemp!kelurahan) = NulSID(rsClone!kelurahan) And isAlamat = False Then
38                            totalDuplikat = totalDuplikat + 1
39                            isAlamat = True
40                        End If
41                        If NulSID(rsTemp!idcust) <> NulSID(rsClone!idcust) And NulSID(rsTemp!hp1) = NulSID(rsClone!hp1) And isNoHP = False Then
42                            totalDuplikat = totalDuplikat + 1
43                            isNoHP = True
44                        End If
                          ''If NulSID(rsTemp!idcust) <> NulSID(rsClone!idcust) And NulSID(rsTemp!noKTP) = NulSID(rsClone!noKTP) And isKTP = False Then
45                        If NulSID(rsTemp!idcust) <> NulSID(rsClone!idcust) And (NulSID(rsTemp!noKTP) = NulSID(rsClone!noKTP) Or NulSID(rsTemp!noNPWP) = NulSID(rsClone!noNPWP)) And isKTP = False Then
46                            totalDuplikat = totalDuplikat + 1
47                            isKTP = True
48                        End If
49                        rsClone.MoveNext
50                    Loop
51                    progBar.value = progBar.value + 1
52                    DoEvents
53                    If totalDuplikat >= 2 Then
54                        GoSub InsertData
55                    End If
56                    rsTemp.MoveNext
57                Loop
58            End If
59            If rsDuplikat.RecordCount > 0 Then
60                MsgBox "Ada " & rsDuplikat.RecordCount & " potensi data duplikasi", vbInformation
61                cmdDuplikasi.Visible = True
62            End If
63        End If
          
64        Grid.Enabled = True
65        Exit Sub
InsertData:
66        With rsDuplikat
67            .AddNew
68            !idcust = NulSID(rsTemp!idcust)
69            !Nama = NulSID(rsTemp!Nama)
70            !tglLahir = Format((rsTemp!tglLahir), "DD/MM/YYYY")
71            !Agama = NulSID(rsTemp!Agama)
72            !pekerjaan = NulSID(rsTemp!pekerjaan)
73            !hobi = NulSID(rsTemp!hobi)
74            !Olahraga = NulSID(rsTemp!Olahraga)
75            !Alamat = NulSID(rsTemp!Alamat)
76            !Kota = NulSID(rsTemp!Kota)
77            !Propinsi = NulSID(rsTemp!Propinsi)
78            !kecamatan = NulSID(rsTemp!kecamatan)
79            !kelurahan = NulSID(rsTemp!kelurahan)
80            !HP = Left(NulSID(rsTemp!hp1), 15)
81            !noKTP = NulSID(rsTemp!noKTP)
82            !noNPWP = NulSID(rsTemp!noNPWP)
              
83            !IDMerge = False
84            !pilihMerge = False
85            !pilihnama = False
86            !pilihTglLahir = False
87            !pilihAgama = False
88            !pilihPekerjaan = False
89            !pilihHobi = False
90            !pilihOlahraga = False
91            !pilihAlamat = False
92            !pilihKota = False
93            !pilihPropinsi = False
94            !pilihKecamatan = False
95            !pilihKeLurahan = False
96            !pilihHP = False
97            !pilihNPWP = False
98            !pilihKTP = False
              
99            .Update
100       End With
101       Return
End Sub

Private Sub PencarianData()
1         cmdDuplikasi.Visible = False
          Dim i As Integer
          Dim txtcari As String
          Dim outPut() As String, param As String, paramX As String, paramBP As String, paramGR As String, paramSingle As String, paramSPK As String, paramGRLanjut As String
2         strTemp = "": strTemp2 = "": strTemp3 = "": strTemp4 = ""
3         cari = "": cari2 = "": cari3 = "": cari4 = "": cari5 = "": cari6 = ""
4         carix = "": cari2x = "": cari3x = "": cari4x = "": cari5x = "": cari6x = ""
5         param = "": paramX = "": paramBP = "": paramGR = ""
6         paramSingle = ""
7         paramSPK = ""
8         daftarIDCust = ""
          Dim statusUnit As String
          
      '=============================================================================================
      '   cari    -   PENCARIAN DI UNIT CUSTOMER
      '=============================================================================================
9         If NulSID(Text(1)) <> "" Then
10            cari = cari & " a.NoRangka Like '%" & NulSID(Text(1)) & "%'"
11            carix = carix & " zNoRangka Like '%" & NulSID(Text(1)) & "%'"
12        End If
13        If NulSID(Text(0)) <> "" Then
14            If NulSID(cari) = "" Then
15                cari = cari & "a.NoPolisi Like '%" & NulSID(Text(0)) & "%'"
16                carix = carix & "zNoPolisi Like '%" & NulSID(Text(0)) & "%'"
17            Else
18                cari = cari & " OR a.NoPolisi Like '%" & NulSID(Text(0)) & "%'"
19                carix = carix & " OR zNoPolisi Like '%" & NulSID(Text(0)) & "%'"
20            End If
21        End If
22        If cari <> "" Then
23            cari = "(" & cari & ")"
24            carix = "(" & carix & ")"
25            paramSingle = paramSingle & " AND (NORANGKA LIKE '%" & NulSID(Text(1)) & "%' OR NOPOLISI LIKE '%" & NulSID(Text(0)) & "%')"
26        End If
          
      '=============================================================================================
      '   cari2   -   NO KTP / NPWP
      '=============================================================================================
27        If NulSID(Text(2)) <> "" Then
28            cari2 = "((a.NoKTPSTNK Like '%" & NulSID(Text(2)) & "%' OR a.NoNPWPSTNK Like '%" & NulSID(Text(2)) & "%') OR (a.NoKTPSL Like '%" & NulSID(Text(2)) & "%' OR a.NoNPWPSL Like '%" & NulSID(Text(2)) & "%') OR (a.NoKTPSV Like '%" & NulSID(Text(2)) & "%' OR a.NoNPWPSV Like '%" & NulSID(Text(2)) & "%') OR (b.NoKTP Like '%" & NulSID(Text(2)) & "%' OR b.NoNPWP Like '%" & NulSID(Text(2)) & "%'))"
29            cari2x = "((zNoKTP Like '%" & NulSID(Text(2)) & "%' OR zNoNPWP Like '%" & NulSID(Text(2)) & "%'))"
30            paramSingle = paramSingle & " AND (NOKTP LIKE '%" & NulSID(Text(2)) & "%' OR NONPWP LIKE '%" & NulSID(Text(2)) & "%')"
31        End If
          
      '=============================================================================================
      '   cari3   -   NAMA CUSTOMER
      '=============================================================================================
32        If NulSID(Text(3)) <> "" Then
      '        outPut = Split(NulSID(text(3)), " ")
              
33            txtcari = Replace(NulSID(Text(3)), " ", "%")
34            strTemp = "": strTemp2 = "": strTemp3 = "": strTemp4 = ""
      '        For i = 0 To UBound(outPut)
      '            If NulSID(outPut(i)) <> "" Then
      ''                If NulSID(strTemp) = "" Then
      ''                    cari3x = "(zNama Like '%" & NulSID(outPut(i)) & "%')"
      ''                    strTemp = "(a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp2 = "(a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp3 = "(a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp4 = "(b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & NulSID(outPut(i)) & "%'"
      ''                Else
      ''                    cari3x = cari3x & " AND (" & "zNama Like '%" & NulSID(outPut(i)) & "%')"
      ''                    strTemp = strTemp & " AND (a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp2 = strTemp2 & " AND (a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp3 = strTemp3 & " AND (a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp4 = strTemp4 & " AND (b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & NulSID(outPut(i)) & "%'"
      ''                End If
      '            End If
      '        Next
35            If NulSID(strTemp) = "" Then
36                cari3x = "(zNama Like '%" & NulSID(txtcari) & "%')"
37                strTemp = "(a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & NulSID(txtcari) & "%'"
38                strTemp2 = "(a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & NulSID(txtcari) & "%'"
39                strTemp3 = "(a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & NulSID(txtcari) & "%'"
40                strTemp4 = "(b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & NulSID(txtcari) & "%'"
41            Else
42                cari3x = cari3x & " AND (" & "zNama Like '%" & NulSID(txtcari) & "%')"
43                strTemp = strTemp & " AND (a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & NulSID(txtcari) & "%'"
44                strTemp2 = strTemp2 & " AND (a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & NulSID(txtcari) & "%'"
45                strTemp3 = strTemp3 & " AND (a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & NulSID(txtcari) & "%'"
46                strTemp4 = strTemp4 & " AND (b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & NulSID(txtcari) & "%'"
47            End If
      '        strTemp = strTemp & " AND (a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & nulsid(Text(3)) & "%'"
      '        strTemp2 = strTemp2 & " AND (a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & nulsid(Text(3)) & "%'"
      '        strTemp3 = strTemp3 & " AND (a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & nulsid(Text(3)) & "%'"
      '        strTemp4 = strTemp4 & " AND (b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & nulsid(Text(3)) & "%'"
              
48            cari3 = " ((" & strTemp & ") OR (" & strTemp2 & ") OR (" & strTemp3 & " ) OR (" & strTemp4 & " ))"
49            If cari3x <> "" Then cari3x = "(" & cari3x & ")"
50            paramSingle = paramSingle & " AND (NAMA1 + ' ' + NAMA2 + '' + NAMA3 LIKE '%" & NulSID(txtcari) & "%')"
51        End If
          
      '=============================================================================================
      '   cari4   -   TANGGAL LAHIR
      '=============================================================================================
52        If NulSID(tglLahir.value) <> "" Then
53            cari4 = "(a.tglLahirSTNK = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "' OR a.tglLahirSL = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "' OR a.tglLahirSV = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "' OR b.tglLahir = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "')"
54            cari4x = "(zTgLahir = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "')"
55            paramSingle = paramSingle & " AND (TGLLAHIR = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "')"
56        End If
          
      '=============================================================================================
      '   cari5   -   KODE CUSTOMER
      '=============================================================================================
57        If NulSID(Text(4)) <> "" Then
58            outPut = Split(NulSID(Text(4)), " ")
59            strTemp = "": strTemp2 = "": strTemp3 = "": strTemp4 = "": strTemp5 = ""
60            For i = 0 To UBound(outPut)
61                If NulSID(outPut(i)) <> "" Then
62                    If NulSID(cari5) = "" Then
63                        cari5 = "(a.IDCustSales Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustService Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustAss Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSTNK Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSL Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSV Like '%" & NulSID(outPut(i)) & "%') OR (b.IDCust Like '%" & NulSID(outPut(i)) & "%')"
64                        cari5x = "(zKodeLg Like '%" & NulSID(outPut(i)) & "%')"
65                    Else
66                        cari5 = cari5 & " OR " & "(a.IDCustSales Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustService Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustAss Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSTNK Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSL Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSV Like '%" & NulSID(outPut(i)) & "%') OR (b.IDCust Like '%" & NulSID(outPut(i)) & "%')"
67                        cari5x = cari5x & " OR " & "(zKodeLg Like '%" & NulSID(outPut(i)) & "%')"
68                    End If
69                End If
70            Next
71            cari5 = "(" & cari5 & ")"
72            cari5x = "(" & cari5x & ")"
73            paramSingle = paramSingle & " AND (IDCUST = '" & NulSID(Text(4)) & "')"
74        End If
      '=============================================================================================
      '   cari6   -   NO TELP
      '=============================================================================================
75        If NulSID(Text(10)) <> "" Then
76            txtcari = Replace(NulSID(Text(10)), " ", "%")
77            strTemp = "": strTemp2 = "": strTemp3 = "": strTemp4 = "": strTemp5 = ""
78            cari6x = "(zTelp Like '%" & NulSID(txtcari) & "%')"
79            strTemp = "(a.IdTelpSTNK + a.TelpSTNK + ' ' + a.HP1STNK + ' ' + a.HP2STNK) Like '%" & NulSID(txtcari) & "%'"
80            strTemp2 = "(a.IdTelpSL + a.TelpSL + ' ' + a.HP1SL + ' ' + a.HP2SL) Like '%" & NulSID(txtcari) & "%'"
81            strTemp3 = "(a.IdTelpSV + a.TelpSV + ' ' + a.HP1SV + ' ' + a.HP2SV) Like '%" & NulSID(txtcari) & "%'"
82            strTemp4 = "(b.IdTelp + b.Telp + ' ' + b.HP1 + ' ' + b.HP2) Like '%" & NulSID(txtcari) & "%'"
83            cari6 = " ((" & strTemp & ") OR (" & strTemp2 & ") OR (" & strTemp3 & " ) OR (" & strTemp4 & " ))"
84            cari6x = "(" & cari6x & ")"
85            paramSingle = paramSingle & " AND (hp1 like '%" & NulSID(Text(10)) & "%')"
86        End If
              
      '=============================================================================================
      '   SETTING POSISI PENCARIAN
      '=============================================================================================
87        If cari = "" Then
88            If cari2 = "" Then
89                If cari3 = "" Then
90                    If cari4 = "" Then
91                        If cari5 = "" Then
92                            If cari6 = "" Then
93                                Exit Sub
94                            Else
95                                param = cari6
96                                paramX = cari6x
97                            End If
98                        Else
99                            If cari6 = "" Then
100                               param = cari5
101                               paramX = cari5x
102                           Else
103                               param = "(" & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
104                               paramX = "(" & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
105                           End If
106                       End If
107                   Else
108                       If cari5 = "" Then
109                           If cari6 = "" Then
110                               param = cari4
111                               paramX = cari4x
112                           Else
113                               param = "(" & cari4 & " AND " & cari6 & ")"
114                               paramX = "(" & cari4x & " AND " & cari6x & ")"
115                           End If
116                       Else
117                           If cari6 = "" Then
118                               param = "(" & cari4 & " AND " & cari5 & ")"
119                               paramX = "(" & cari4x & " AND " & cari5x & ")"
120                           Else
121                               param = "(" & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
122                               paramX = "(" & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
123                           End If
124                       End If
125                   End If
126               Else
127                   If cari4 = "" Then
128                       If cari5 = "" Then
129                           If cari6 = "" Then
130                               param = cari3
131                               paramX = cari3x
132                           Else
133                               param = "(" & cari3 & " AND " & cari6 & ")"
134                               paramX = "(" & cari3x & " AND " & cari6x & ")"
135                           End If
136                       Else
137                           If cari6 = "" Then
138                               param = "(" & cari3 & " AND " & cari5 & ")"
139                               paramX = "(" & cari3x & " AND " & cari5x & ")"
140                           Else
141                               param = "(" & cari3 & " AND " & cari5 & " AND " & cari6 & ")"
142                               paramX = "(" & cari3x & " AND " & cari5x & " AND " & cari6x & ")"
143                           End If
144                       End If
145                   Else
146                       If cari5 = "" Then
147                           If cari6 = "" Then
148                               param = "(" & cari3 & " AND " & cari4 & ")"
149                               paramX = "(" & cari3x & " AND " & cari4x & ")"
150                           Else
151                               param = "(" & cari3 & " AND " & cari4 & " AND " & cari6 & ")"
152                               paramX = "(" & cari3x & " AND " & cari4x & " AND " & cari6x & ")"
153                           End If
154                       Else
155                           If cari6 = "" Then
156                               param = "(" & cari3 & " AND " & cari4 & " AND " & cari5 & ")"
157                               paramX = "(" & cari3x & " AND " & cari4x & " AND " & cari5x & ")"
158                           Else
159                               param = "(" & cari3 & " AND " & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
160                               paramX = "(" & cari3x & " AND " & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
161                           End If
162                       End If
163                   End If
164               End If
165           Else
166               If cari3 = "" Then
167                   If cari4 = "" Then
168                       If cari5 = "" Then
169                           If cari6 = "" Then
170                               param = cari2
171                               paramX = cari2x
172                           Else
173                               param = "(" & cari2 & " AND " & cari6 & ")"
174                               paramX = "(" & cari2x & " AND " & cari6x & ")"
175                           End If
176                       Else
177                           If cari6 = "" Then
178                               param = "(" & cari2 & " AND " & cari5 & ")"
179                               paramX = "(" & cari2x & " AND " & cari5x & ")"
180                           Else
181                               param = "(" & cari2 & " AND " & cari5 & " AND " & cari6 & ")"
182                               paramX = "(" & cari2x & " AND " & cari5x & " AND " & cari6x & ")"
183                           End If
                              
184                       End If
185                   Else
186                       If cari5 = "" Then
187                           If cari6 = "" Then
188                               param = "(" & cari2 & " AND " & cari4 & ")"
189                               paramX = "(" & cari2x & " AND " & cari4x & ")"
190                           Else
191                               param = "(" & cari2 & " AND " & cari4 & " AND " & cari6 & ")"
192                               paramX = "(" & cari2x & " AND " & cari4x & " AND " & cari6x & ")"
193                           End If
194                       Else
195                           If cari6 = "" Then
196                               param = "(" & cari2 & " AND " & cari4 & " AND " & cari5 & ")"
197                               paramX = "(" & cari2x & " AND " & cari4x & " AND " & cari5x & ")"
198                           Else
199                               param = "(" & cari2 & " AND " & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
200                               paramX = "(" & cari2x & " AND " & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
201                           End If
202                       End If
203                   End If
204               Else
205                   If cari4 = "" Then
206                       If cari5 = "" Then
207                           If cari6 = "" Then
208                               param = "(" & cari2 & " AND " & cari3 & ")"
209                               paramX = "(" & cari2x & " AND " & cari3x & ")"
210                           Else
211                               param = "(" & cari2 & " AND " & cari3 & " AND " & cari6 & ")"
212                               paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari6x & ")"
213                           End If
214                       Else
215                           If cari6 = "" Then
216                               param = "(" & cari2 & " AND " & cari3 & " AND " & cari5 & ")"
217                               paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari5x & ")"
218                           Else
219                               param = "(" & cari2 & " AND " & cari3 & " AND " & cari5 & " AND " & cari6 & ")"
220                               paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari5x & " AND " & cari6x & ")"
221                           End If
222                       End If
223                   Else
224                       If cari5 = "" Then
225                           If cari6 = "" Then
226                               param = "(" & cari2 & " AND " & cari3 & " AND " & cari4 & ")"
                                  
227                               paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari4x & ")"
228                           Else
229                               param = "(" & cari2 & " AND " & cari3 & " AND " & cari4 & " AND " & cari6 & ")"
230                               paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari4x & " AND " & cari6x & ")"
231                           End If
232                       Else
233                           If cari6 = "" Then
234                               param = "(" & cari2 & " AND " & cari3 & " AND " & cari4 & " AND " & cari5 & ")"
235                               paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari4x & " AND " & cari5x & ")"
236                           Else
237                               param = "(" & cari2 & " AND " & cari3 & " AND " & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
238                               paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
239                           End If
240                       End If
241                   End If
242               End If
243           End If
244       Else
245           If cari2 = "" Then
246               If cari3 = "" Then
247                   If cari4 = "" Then
248                       If cari5 = "" Then
249                           If cari6 = "" Then
250                               param = cari
251                               paramX = carix
252                           Else
253                               param = "(" & cari & ") AND (" & cari6 & ")"
254                               paramX = "(" & carix & ") AND (" & cari6x & ")"
255                           End If
256                       Else
257                           If cari6 = "" Then
258                               param = "(" & cari & ") AND (" & cari5 & ")"
259                               paramX = "(" & carix & ") AND (" & cari5x & ")"
260                           Else
261                               param = "(" & cari & ") AND (" & cari5 & " AND " & cari6 & ")"
262                               paramX = "(" & carix & ") AND (" & cari5x & " AND " & cari6x & ")"
263                           End If
264                       End If
265                   Else
266                       If cari5 = "" Then
267                           If cari6 = "" Then
268                               param = "(" & cari & ") AND (" & cari4 & ")"
269                               paramX = "(" & carix & ") AND (" & cari4x & ")"
270                           Else
271                               param = "(" & cari & ") AND (" & cari4 & " AND " & cari6 & ")"
272                               paramX = "(" & carix & ") AND (" & cari4x & " AND " & cari6x & ")"
273                           End If
274                       Else
275                           If cari6 = "" Then
276                               param = "(" & cari & ") AND (" & cari4 & " AND " & cari5 & ")"
277                               paramX = "(" & carix & ") AND (" & cari4x & " AND " & cari5x & ")"
278                           Else
279                               param = "(" & cari & ") AND (" & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
280                               paramX = "(" & carix & ") AND (" & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
281                           End If
282                       End If
283                   End If
284               Else
285                   If cari4 = "" Then
286                       If cari5 = "" Then
287                           If cari6 = "" Then
288                               param = "(" & cari & ") AND (" & cari3 & ")"
289                               paramX = "(" & carix & ") AND (" & cari3x & ")"
290                           Else
291                               param = "(" & cari & ") AND (" & cari3 & " AND " & cari6 & ")"
292                               paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari6x & ")"
293                           End If
294                       Else
295                           If cari6 = "" Then
296                               param = "(" & cari & ") AND (" & cari3 & " AND " & cari5 & ")"
297                               paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari5x & ")"
298                           Else
299                               param = "(" & cari & ") AND (" & cari3 & " AND " & cari5 & " AND " & cari6 & ")"
300                               paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari5x & " AND " & cari6x & ")"
301                           End If
302                       End If
303                   Else
304                       If cari5 = "" Then
305                           If cari6 = "" Then
306                               param = "(" & cari & ") AND (" & cari3 & " AND " & cari4 & ")"
307                               paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari4x & ")"
308                           Else
309                               param = "(" & cari & ") AND (" & cari3 & " AND " & cari4 & " AND " & cari6 & ")"
310                               paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari4x & " AND " & cari6x & ")"
311                           End If
312                       Else
313                           If cari6 = "" Then
314                               param = "(" & cari & ") AND (" & cari3 & " AND " & cari4 & " AND " & cari5 & ")"
315                               paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari4x & " AND " & cari5x & ")"
316                           Else
317                               param = "(" & cari & ") AND (" & cari3 & " AND " & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
318                               paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
319                           End If
320                       End If
321                   End If
322               End If
323           Else
324               If cari3 = "" Then
325                   If cari4 = "" Then
326                       If cari5 = "" Then
327                           If cari6 = "" Then
328                               param = "(" & cari & ") AND (" & cari2 & ")"
329                               paramX = "(" & carix & ") AND (" & cari2x & ")"
330                           Else
331                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari6 & ")"
332                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari6x & ")"
333                           End If
334                       Else
335                           If cari6 = "" Then
336                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari5 & ")"
337                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari5x & ")"
338                           Else
339                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari5 & " AND " & cari6 & ")"
340                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari5x & " AND " & cari6x & ")"
341                           End If
342                       End If
343                   Else
344                       If cari5 = "" Then
345                           If cari6 = "" Then
346                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari4 & ")"
347                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari4x & ")"
348                           Else
349                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari4 & " AND " & cari6 & ")"
350                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari4x & " AND " & cari6x & ")"
351                           End If
352                       Else
353                           If cari6 = "" Then
354                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari4 & " AND " & cari5 & ")"
355                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari4x & " AND " & cari5x & ")"
356                           Else
357                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
358                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
359                           End If
360                       End If
361                   End If
362               Else
363                   If cari4 = "" Then
364                       If cari5 = "" Then
365                           If cari6 = "" Then
366                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & ")"
367                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & ")"
368                           Else
369                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari6 & ")"
370                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari6x & ")"
371                           End If
372                       Else
373                           If cari6 = "" Then
374                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari5 & ")"
375                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari5x & ")"
376                           Else
377                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari5 & ")"
378                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari5x & ")"
379                           End If
380                       End If
381                   Else
382                       If cari5 = "" Then
383                           If cari6 = "" Then
384                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari4 & ")"
385                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari4x & ")"
386                           Else
387                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari4 & " AND " & cari6 & ")"
388                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari4x & " AND " & cari6x & ")"
389                           End If
390                       Else
391                           If cari6 = "" Then
392                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari4 & " AND " & cari5 & ")"
393                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari4x & " AND " & cari5x & ")"
394                           Else
395                               param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari4 & " AND " & cari5 & " AND " & cari6 & ")"
396                               paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari4x & " AND " & cari5x & " AND " & cari6x & ")"
397                           End If
398                       End If
399                   End If
400               End If
401           End If
402       End If
          
      '    sSQL = "Select * From ( Select a.NoRangka, a.IDCustSTNK as IDCust, a.NoKTPSTNK as NoKTP, a.NoNPWPSTNK as NoNPWP, a.TitlesSTNK as Title, a.NamaCust1STNK as Nama1, a.NamaCust2STNK as Nama2, a.NamaCust3STNK as Nama3, a.TglLahirSTNK as TglLahir, a.AgamaSTNK as Agama, " & _
      '           "a.PekerjaanSTNK as Pekerjaan, a.PropinsiSTNK as Propinsi, a.KotaSTNK as Kota, a.KecamatanSTNK as Kecamatan, a.KelurahanSTNK as Kelurahan, a.KodePosSTNK as KodePos, a.AlamatSTNK as Alamat, a.RTSTNK as RT, a.RWSTNK as RW, a.IdTelpSTNK as KodeTelp, " & _
      '           "a.TelpSTNK as Telp, a.HP1STNK as HP1, a.HP2STNK as HP2, a.EmailSTNK as Email, a.HobiSTNK as Hobi, a.OlahragaSTNK as Olahraga, a.SukuSTNK as Suku, a.PreferensiSTNK as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, " & _
      '           "a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a Left Join NewCustomer b On a.IDCustSTNK = b.IDCust Where " & param & " Union Select a.NoRangka, a.IDCustSL as IDCust, a.NoKTPSL as " & _
      '           "NoKTP, a.NoNPWPSL as NoNPWP, a.TitlesSL as Title, a.NamaCust1SL as Nama1, a.NamaCust2SL as Nama2, a.NamaCust3SL as Nama3, a.TglLahirSL as tglLahir, a.AgamaSL as Agama, a.PekerjaanSL as Pekerjaan, a.PropinsiSL as Propinsi, a.KotaSL as Kota, " & _
      '           "a.KecamatanSL as Kecamatan, a.KelurahanSL as Kelurahan, a.KodePosSL as KodePos, a.AlamatSL as Alamat, a.RTSL as RT, a.RWSL as RW, a.IdTelpSL as KodeTelp, a.TelpSL as Telp, a.HP1SL as HP1, a.HP2SL as HP2, a.EmailSL as Email, a.HobiSL as Hobi, " & _
      '           "a.OlahragaSL as Olahraga, a.SukuSL as Suku, a.PreferensiSL as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil " & _
      '           "From UnitCustomer a Left Join NewCustomer b On a.IDCustSL = b.IDCust Where " & param & " Union Select a.NoRangka, a.IDCustSV as IDCust, a.NoKTPSV as NoKTP, a.NoNPWPSV as NoNPWP, a.TitlesSV as Title, a.NamaCust1SV as Nama1, a.NamaCust2SV as Nama2, " & _
      '           "a.NamaCust3SV as Nama3, a.TglLahirSV as TglLahir, a.AgamaSV as Agama, a.PekerjaanSV as Pekerjaan, a.PropinsiSV as Propinsi, a.KotaSV as Kota, a.KecamatanSV as Kecamatan, a.KelurahanSV as Kelurahan, a.KodePosSV as KodePos, a.AlamatSV as Alamat, " & _
      '           "a.RTSV as RT, a.RWSV as RW, a.IdTelpSV as KodeTelp, a.TelpSV as Telp, a.HP1SV as HP1, a.HP2SV as HP2, a.EmailSV as Email, a.HobiSV as Hobi, a.OlahragaSV as Olahraga, a.SukuSV as Suku, a.PreferensiSV as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, " & _
      '           "a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a Left Join NewCustomer b On a.IDCustSV = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, " & _
      '           "b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, " & _
      '           "b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a INNER Join NewCustomer b On " & _
      '           "a.IDCustService = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, " & _
      '           "b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, " & _
      '           "a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a INNER Join NewCustomer b On a.IDCustSales = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, " & _
      '           "b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, " & _
      '           "a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a INNER Join NewCustomer b On a.IDCustAss = b.IDCust Where " & param
          
      '    sSQL = "Select * From ( Select a.NoRangka, a.IDCustSTNK as IDCust, a.NoKTPSTNK as NoKTP, a.NoNPWPSTNK as NoNPWP, a.NamaCust1STNK as Nama1, a.NamaCust2STNK as Nama2, a.NamaCust3STNK as Nama3, a.TglLahirSTNK as TglLahir, " & _
      '           " a.AlamatSTNK as Alamat, " & _
      '           " a.HP1STNK as HP1, a.NoPolisi, " & _
      '           "  From NewCustomer b Left Join UnitCustomer a On a.IDCustSTNK = b.IDCust Where " & param & " Union Select a.NoRangka, a.IDCustSL as IDCust, a.NoKTPSL as " & _
      '           "NoKTP, a.NoNPWPSL as NoNPWP, a.NamaCust1SL as Nama1, a.NamaCust2SL as Nama2, a.NamaCust3SL as Nama3, a.TglLahirSL as tglLahir, " & _
      '           "a.KecamatanSL as Kecamatan, a.KelurahanSL as Kelurahan, a.KodePosSL as KodePos, a.AlamatSL as Alamat, a.RTSL as RT, a.RWSL as RW, a.IdTelpSL as KodeTelp, a.TelpSL as Telp, a.HP1SL as HP1, a.HP2SL as HP2, a.EmailSL as Email, a.HobiSL as Hobi, " & _
      '           "a.OlahragaSL as Olahraga, a.SukuSL as Suku, a.PreferensiSL as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn  " & _
      '           "From NewCustomer b Left Join UnitCustomer a On a.IDCustSL = b.IDCust Where " & param & " Union Select a.NoRangka, a.IDCustSV as IDCust, a.NoKTPSV as NoKTP, a.NoNPWPSV as NoNPWP, a.TitlesSV as Title, a.NamaCust1SV as Nama1, a.NamaCust2SV as Nama2, " & _
      '           "a.NamaCust3SV as Nama3, a.TglLahirSV as TglLahir, a.AgamaSV as Agama, a.PekerjaanSV as Pekerjaan, a.PropinsiSV as Propinsi, a.KotaSV as Kota, a.KecamatanSV as Kecamatan, a.KelurahanSV as Kelurahan, a.KodePosSV as KodePos, a.AlamatSV as Alamat, " & _
      '           "a.RTSV as RT, a.RWSV as RW, a.IdTelpSV as KodeTelp, a.TelpSV as Telp, a.HP1SV as HP1, a.HP2SV as HP2, a.EmailSV as Email, a.HobiSV as Hobi, a.OlahragaSV as Olahraga, a.SukuSV as Suku, a.PreferensiSV as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, " & _
      '           "a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn  From NewCustomer b Left Join UnitCustomer a On a.IDCustService = b.IDCust AND b.IDCust <> a.idcustsv Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, " & _
      '           "b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, " & _
      '           "b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn From NewCustomer b left join UnitCustomer a On " & _
      '           "a.IDCustService = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, " & _
      '           "b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, " & _
      '           "a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn From NewCustomer b left Join UnitCustomer a On a.IDCustSales = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, " & _
      '           "b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, " & _
      ''           "a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn From NewCustomer b left Join UnitCustomer a On a.IDCustAss = b.IDCust Where " & param
403               sSQL = " Select * From ("
404               sSQL = sSQL & " Select a.NoRangka, a.IDCustSTNK as IDCust, a.NoKTPSTNK as NoKTP, a.NoNPWPSTNK as NoNPWP, a.NamaCust1STNK as Nama1, a.NamaCust2STNK as Nama2, a.NamaCust3STNK as Nama3, a.TglLahirSTNK as TglLahir,  a.AlamatSTNK as Alamat, a.IdTelpSTNK + a.TelpSTNK + ',' + a.HP1STNK + ',' + a.HP2STNK as HP1, a.NoPolisi, a.KotaSTNK as Kota,'' AS KodeLgn"
405               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
406               sSQL = sSQL & " Left Join UnitCustomer a WITH (NOLOCK) On a.IDCustSTNK = b.IDCust"
407               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
408               sSQL = sSQL & " Union"
409               sSQL = sSQL & " Select a.NoRangka, a.IDCustSL as IDCust, a.NoKTPSL as NoKTP, a.NoNPWPSL as NoNPWP, a.NamaCust1SL as Nama1, a.NamaCust2SL as Nama2, a.NamaCust3SL as Nama3, a.TglLahirSL as tglLahir, a.AlamatSL as Alamat, a.IdTelpSL + a.TelpSL + ',' + a.HP1SL + ',' + a.HP2SL as HP1, a.NoPolisi, a.KotaSl as Kota,'' AS KodeLgn"
410               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
411               sSQL = sSQL & " Left Join UnitCustomer a WITH (NOLOCK)  On a.IDCustSL = b.IDCust"
412               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
413               sSQL = sSQL & " Union"
414               sSQL = sSQL & " Select a.NoRangka, a.IDCustSV as IDCust, a.NoKTPSV as NoKTP, a.NoNPWPSV as NoNPWP, a.NamaCust1SV as Nama1, a.NamaCust2SV as Nama2, a.NamaCust3SV as Nama3, a.TglLahirSV as TglLahir,a.AlamatSV as Alamat,a.IdTelpSv + a.TelpSv + ',' + a.HP1SV + ',' + a.HP2SV as HP1,a.NoPolisi, a.KotaSv as Kota,'' AS KodeLgn"
415               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
416               sSQL = sSQL & " Left Join UnitCustomer a WITH (NOLOCK)  On a.IDCustService = b.IDCust AND b.IDCust <> a.idcustsv"
417               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
418               sSQL = sSQL & " Union"
419               sSQL = sSQL & " Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir,b.Alamat,b.IdTelp + b.Telp + ',' + b.HP1 + ',' + b.HP2 as HP1, a.NoPolisi, b.Kota as Kota,'' AS KodeLgn"
420               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
421               sSQL = sSQL & " left join UnitCustomer a WITH (NOLOCK)  On a.IDCustService = b.IDCust"
422               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
423               sSQL = sSQL & " Union"
424               sSQL = sSQL & " Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP,b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir,b.Alamat,b.IdTelp + b.Telp + ',' + b.HP1 + ',' + b.HP2 as HP1,a.NoPolisi, b.kota as Kota,'' AS KodeLgn"
425               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
426               sSQL = sSQL & " left Join UnitCustomer a WITH (NOLOCK)  On a.IDCustSales = b.IDCust"
427               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
428               sSQL = sSQL & " Union"
429               sSQL = sSQL & " Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir,b.Alamat, b.IdTelp + b.Telp + ',' + b.HP1 + ',' + b.HP2 as HP1,a.NoPolisi, b.kota as Kota,'' AS KodeLgn"
430               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
431               sSQL = sSQL & " left Join UnitCustomer a WITH (NOLOCK)  On a.IDCustAss = b.IDCust"
432               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
433               sSQL = sSQL & " ) Z Where Nama1 + Nama2 + Nama3 <> ''" & paramSingle
          'sSQL = sSQL & " ) Z Where Nama1 + Nama2 + Nama3 <> '' " & paramSingle
434       Set rsTemp = Nothing
435       lblKeteranganLoading.Caption = "PENCARIAN DATA SINGLE ID ......."
436       progBar.value = 1
437       DoEvents
438       Set rsTemp = objSID.FetchDataSID(sSQL)
439       Call createRsGrid
          
440       GoSub InsertData
          
441       paramBP = paramX: paramGR = paramX: paramSPK = paramX
          '======================================================================================================================
          '   CARI DI CUSTDAT BP
          '======================================================================================================================
442       paramBP = Replace(paramBP, "zNoRangka", "a.NoRangka")
443       paramBP = Replace(paramBP, "zNoPolisi", "a.NoPolisi")
444       paramBP = Replace(paramBP, "zNoKTP", "a.NoKTP")
445       paramBP = Replace(paramBP, "zNoNPWP", "a.NoNPWP")
446       paramBP = Replace(paramBP, "zNama", "(a.Nama1 + a.nama2 + a.nama3)")
447       paramBP = Replace(paramBP, "zTgLahir", "a.TglLahir")
448       paramBP = Replace(paramBP, "zKodeLg", "a.IDCust")
449       paramBP = Replace(paramBP, "zTelp", "a.HP1")
450       Set rsTemp = New ADODB.Recordset
          
451       lblKeteranganLoading.Caption = "PENCARIAN DATA BENGKEL ......."
452       progBar.value = 2
453       DoEvents
454       If cekKoneksiBP = True Then
455           Set rsTemp = objSID.FetchListAllPelanggan("Where " & paramBP)
456           GoSub InsertData
457       End If
          '======================================================================================================================
          '   CARI DI CUSTDAT GRP
          '======================================================================================================================
          
458       paramGRLanjut = paramGR
          
459       paramGRLanjut = Replace(paramGRLanjut, "zNoRangka", "c.NO_CHASSIS")
460       paramGRLanjut = Replace(paramGRLanjut, "zNoPolisi", "c.No_Polisi")
461       paramGRLanjut = Replace(paramGRLanjut, "zNama", "(c.Customer)")
462       paramGRLanjut = Replace(paramGRLanjut, "zKodeLg", "(Isnull(e.IDCust,''))")
463       paramGRLanjut = Replace(paramGRLanjut, "zTelp", "c.UseGSM")
464       paramGRLanjut = Replace(paramGRLanjut, "zNoKTP", "''")
465       paramGRLanjut = Replace(paramGRLanjut, "zNoNPWP", "''")
466       paramGRLanjut = Replace(paramGRLanjut, "zTgLahir", "''")
          
467       paramGR = Replace(paramGR, "zNoRangka", "c.norangka")
468       paramGR = Replace(paramGR, "zNoPolisi", "c.NoPolisi")
469       paramGR = Replace(paramGR, "zNoKTP", "c.NoKTP")
470       paramGR = Replace(paramGR, "zNoNPWP", "c.NoNPWP")
471       paramGR = Replace(paramGR, "zNama", "(c.nama1 + c.nama2 + c.nama3)")
472       paramGR = Replace(paramGR, "zTgLahir", "c.TglLahir")
473       paramGR = Replace(paramGR, "zKodeLg", "c.IDCust")
474       paramGR = Replace(paramGR, "zTelp", "c.HP1")
475       Set rsTemp = New ADODB.Recordset
          
476       lblKeteranganLoading.Caption = "PENCARIAN DATA BENGKEL ......."
477       progBar.value = 3
478       DoEvents
          
479       Set rsTemp = objSID.FetchPelangganGRP("Where " & paramGR, "Where " & paramGRLanjut)
480       If Not rsTemp.EOF Then
481       GoSub InsertData
482       End If
          
          '======================================================================================================================
          '   CARI DI CUSTDAT SPK
          '======================================================================================================================
483       paramSPK = Replace(paramSPK, "zNoRangka", "c.norangka")
484       paramSPK = Replace(paramSPK, "zNoPolisi", "c.NoPolisi")
485       paramSPK = Replace(paramSPK, "zNoKTP", "c.NoKTP")
486       paramSPK = Replace(paramSPK, "zNoNPWP", "c.NoNPWP")
487       paramSPK = Replace(paramSPK, "zNama", "c.NamaFull")
488       paramSPK = Replace(paramSPK, "zTgLahir", "c.TglLahir")
489       paramSPK = Replace(paramSPK, "zKodeLg", "c.IDCust")
490       paramSPK = Replace(paramSPK, "zTelp", "c.HP1")
491       Set rsTemp = New ADODB.Recordset
          
492       lblKeteranganLoading.Caption = "PENCARIAN DATA SPK ......."
493       progBar.value = 4
494       DoEvents
          
495       Set rsTemp = objSID.fetchPelangganSPK("Where " & paramSPK)
496       GoSub InsertData
          
497       Set Grid.DataSource = rsGrid
498       btnTambah.Enabled = False
499       frameLoading.Visible = False
500       If rsGrid.EOF Then
501           btnTambah.Enabled = True
502           cmdDijual.Visible = False
503       Else
504           If Log.monitoringScoring = 1 Then
505               cmdDijual.Visible = True
506           End If
507       End If
508       Exit Sub
          
InsertData:
509       With rsGrid
510           If Not rsTemp.EOF Then 'If rsTemp.RecordCount <> 0 Then
511               rsTemp.MoveFirst
512               Do Until rsTemp.EOF
513                   .AddNew
514                   statusUnit = ""
515                   !Detail = "Ubah" 'False
516                   !Proses = "Proses" 'False
517                   !idcust = NulSID(rsTemp!idcust)
518                   If Log.monitoringScoring = 1 And NulSID(rsTemp!idcust) <> "" Then
519                       daftarIDCust = daftarIDCust & "'" & NulSID(rsTemp!idcust) & "' ,"
520                       totalData = totalData + 1
521                   End If
522                   !namaCust = NulSID(rsTemp!nama1) & " " & NulSID(rsTemp!nama2) & " " & NulSID(rsTemp!nama3)
523                   !tglLahir = Format((rsTemp!tglLahir), "DD/MM/YYYY")
524                   !NoPolisi = NulSID(rsTemp!NoPolisi)
525                   !Alamat = NulSID(rsTemp!Alamat)
526                   !Kota = NulSID(rsTemp!Kota)
527                   !HP = Left(NulSID(rsTemp!hp1), 15)
528                   !kodeLgn = NulSID(rsTemp!kodeLgn)
529                   If Len(NulSID(rsTemp!noKTP)) > 0 Then
530                       If Len(NulSID(rsTemp!noNPWP)) > 0 Then
531                           !NoID = NulSID(rsTemp!noKTP) & " - " & NulSID(rsTemp!noNPWP)
532                       Else
533                           !NoID = NulSID(rsTemp!noKTP)
534                       End If
535                   Else
536                       If Len(NulSID(rsTemp!noNPWP)) > 0 Then
537                           !NoID = NulSID(rsTemp!noKTP)
538                       Else
539                           !NoID = ""
540                       End If
541                   End If
542                   !noRangka = NulSID(rsTemp!noRangka)
543                   !noKTP = NulSID(rsTemp!noKTP)
544                   !noNPWP = NulSID(rsTemp!noNPWP)
545                   If Log.monitoringScoring = 1 Then
546                       !Split = "Split"
547                       !kodeLgn = NulSID(rsGrid!kodeLgn)
548                       sSQL = "SELECT dijual FROM [singleidori]..DataDijual WITH (NOLOCK) WHERE idcust = '" & NulSID(rsTemp!idcust) & "' AND norangka = '" & NulSID(rsTemp!noRangka) & "'"
549                       statusUnit = objO.FetchData_String(sSQL)
550                       If statusUnit = "" Or statusUnit = "0" Then
551                           !statusUnit = "Active"
552                       ElseIf statusUnit = "1" Then
553                           !statusUnit = "DIJUAL / NO ID"
554                       ElseIf statusUnit = "" Then
                          
555                       End If
556                   End If
557                   .Update
558                   rsTemp.MoveNext
559               Loop
560           End If
561       End With
562       Return
End Sub

Private Sub pencarianDataLanjut()
    Dim bolehTambah As Boolean, kodeLgn As String, isWhere As String
    On Error GoTo pencarianDataLanjut_Error
    
    Call createCustomer
    Call createTempCustomer
    
'===============================================================================================
'   DATA BP
'===============================================================================================
    strWhere = "": isWhere = ""
    If NulSID(Text(0)) <> "" Then strWhere = " c.NoPolisi LIKE '%" & NulSID(Text(0)) & "%'"
    If NulSID(Text(1)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NoRangka LIKE '%" & NulSID(Text(1)) & "%'"
        Else
            strWhere = strWhere & " c.NoRangka LIKE '%" & NulSID(Text(1)) & "%'"
        End If
    End If
    
    If NulSID(Text(3)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.Nama LIKE '%" & NulSID(Text(3)) & "%'"
        Else
            strWhere = strWhere & " c.Nama LIKE '%" & NulSID(Text(3)) & "%'"
        End If
    End If
       
    Set rsPel = objSID.FetchListAllPelanggan("Where " & strWhere)
    If rsPel.RecordCount > 0 Then
        rsPel.MoveFirst
        Do Until rsPel.EOF
            bolehTambah = False
            If NulSID(kodeLgn) = "" Then
                kodeLgn = NulSID(rsPel!KodeLg)
                bolehTambah = True
            ElseIf NulSID(kodeLgn) <> NulSID(rsPel!KodeLg) Then
                kodeLgn = NulSID(rsPel!KodeLg)
                bolehTambah = True
            End If
            
            If bolehTambah Then
                With rsTempCust
                    .AddNew
                    !kodeLgn = NulSID(rsPel!KodeLg)
                    !KodeDealer = Right(gDb, 2)
                    !noRangka = NulSID(rsPel!noRangka)
                    !Tanggal = Format(NulSID(rsPel!tglentry), "dd/mm/yyyy")
                    .Update
                End With
            End If
            rsPel.MoveNext
        Loop
    End If
    
'===============================================================================================
'   DATA GRP
'===============================================================================================
    strWhere = "": isWhere = ""
    If NulSID(Text(0)) <> "" Then strWhere = " c.NO_POLISI LIKE '%" & NulSID(Text(0)) & "%'"
    If NulSID(Text(1)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NO_CHASSIS LIKE '%" & NulSID(Text(1)) & "%'"
        Else
            strWhere = strWhere & " c.NO_CHASSIS LIKE '%" & NulSID(Text(1)) & "%'"
        End If
    End If
    
    If NulSID(Text(2)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NPWP LIKE '%" & NulSID(Text(2)) & "%'"
        Else
            strWhere = strWhere & " c.NPWP LIKE '%" & NulSID(Text(2)) & "%'"
        End If
    End If
    
    If NulSID(Text(3)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.CUSTOMER LIKE '%" & NulSID(Text(3)) & "%'"
        Else
            strWhere = strWhere & " c.CUSTOMER LIKE '%" & NulSID(Text(3)) & "%'"
        End If
    End If
    
    If NulSID(tglLahir.value) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.USETGLLAHIR = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        Else
            strWhere = strWhere & " AND c.USETGLLAHIR = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        End If
    End If
         
    Set rsPel = objSID.FetchPelangganGRP("Where " & strWhere, "Where " & strWhere)
    If rsPel.RecordCount > 0 Then
        rsPel.MoveFirst
        Do Until rsPel.EOF
            bolehTambah = False
            If NulSID(kodeLgn) = "" Then
                kodeLgn = NulSID(rsPel!kodeLgn)
                bolehTambah = True
            ElseIf NulSID(kodeLgn) <> NulSID(rsPel!kodeLgn) Then
                kodeLgn = NulSID(rsPel!kodeLgn)
                bolehTambah = True
            End If
            
            If bolehTambah Then
                With rsTempCust
                    .AddNew
                    !kodeLgn = NulSID(rsPel!kodeLgn)
                    !KodeDealer = Right(gDb, 2)
                    !noRangka = NulSID(rsPel!NO_CHASSIS)
                    !Tanggal = Format(NulSID(isnow), "dd/mm/yyyy")
                    .Update
                End With
            End If
            rsPel.MoveNext
        Loop
    End If
    
'===============================================================================================
'   DATA SPK
'===============================================================================================
    strWhere = "": isWhere = ""
    If NulSID(Text(0)) <> "" Then strWhere = " a.NoPol LIKE '%" & NulSID(Text(0)) & "%'"
    If NulSID(Text(1)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NoRangka LIKE '%" & NulSID(Text(1)) & "%'"
        Else
            strWhere = strWhere & " c.NoRangka LIKE '%" & NulSID(Text(1)) & "%'"
        End If
    End If
    
    If NulSID(Text(2)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NPWP LIKE '%" & NulSID(Text(2)) & "%'"
        Else
            strWhere = strWhere & " c.NPWP LIKE '%" & NulSID(Text(2)) & "%'"
        End If
    End If
    
    If NulSID(Text(3)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.Nama_Pemesan LIKE '%" & NulSID(Text(3)) & "%'"
        Else
            strWhere = strWhere & " c.Nama_Pemesan LIKE '%" & NulSID(Text(3)) & "%'"
        End If
    End If
    
    If NulSID(tglLahir.value) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.TglLahir = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        Else
            strWhere = strWhere & " AND c.TglLahir = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        End If
    End If
    
    Set rsPel = objSID.fetchPelangganSPK("Where " & strWhere)
    If rsPel.RecordCount > 0 Then
        rsPel.MoveFirst
        Do Until rsPel.EOF
            bolehTambah = False
            If NulSID(kodeLgn) = "" Then
                kodeLgn = NulSID(rsPel!kodeLgn)
                bolehTambah = True
            ElseIf NulSID(kodeLgn) <> NulSID(rsPel!kodeLgn) Then
                kodeLgn = NulSID(rsPel!kodeLgn)
                bolehTambah = True
            End If
            
            If bolehTambah Then
                With rsTempCust
                    .AddNew
                    !kodeLgn = NulSID(rsPel!kodeLgn)
                    !KodeDealer = Right(gDb, 2)
                    !noRangka = NulSID(rsPel!noRangka)
                    !Tanggal = Format(NulSID(rsPel!tglentry), "dd/mm/yyyy")
                    .Update
                End With
            End If
            rsPel.MoveNext
        Loop
    End If
       
'===============================================================================================
'   DATA SPK HO
'===============================================================================================
    strWhere = "": isWhere = ""
    If NulSID(Text(0)) <> "" Then strWhere = " a.NoPol LIKE '%" & NulSID(Text(0)) & "%'"
    If NulSID(Text(1)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND a.NoRangka LIKE '%" & NulSID(Text(1)) & "%'"
        Else
            strWhere = strWhere & " a.NoRangka LIKE '%" & NulSID(Text(1)) & "%'"
        End If
    End If
    
    If NulSID(Text(2)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NPWP LIKE '%" & NulSID(Text(2)) & "%'"
        Else
            strWhere = strWhere & " c.NPWP LIKE '%" & NulSID(Text(2)) & "%'"
        End If
    End If
    
    If NulSID(Text(3)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.Nama_Pemesan LIKE '%" & NulSID(Text(3)) & "%'"
        Else
            strWhere = strWhere & " c.Nama_Pemesan LIKE '%" & NulSID(Text(3)) & "%'"
        End If
    End If
    
    If NulSID(tglLahir.value) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.TglLahir = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        Else
            strWhere = strWhere & " AND c.TglLahir = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        End If
    End If
        
    Set rsPel = objSID.FetchPelangganSPK00("Where " & strWhere)
    If rsPel.RecordCount > 0 Then
        rsPel.MoveFirst
        Do Until rsPel.EOF
            bolehTambah = False
            If NulSID(kodeLgn) = "" Then
                kodeLgn = NulSID(rsPel!KodeLg)
                bolehTambah = True
            ElseIf NulSID(kodeLgn) <> NulSID(rsPel!KodeLg) Then
                kodeLgn = NulSID(rsPel!KodeLg)
                bolehTambah = True
            End If
            
            If bolehTambah Then
                With rsTempCust
                    .AddNew
                    !kodeLgn = NulSID(rsPel!KodeLg)
                    !KodeDealer = Right(rsPel!DBname, 2)
                    !noRangka = NulSID(rsPel!noRangka)
                    !Tanggal = Format(NulSID(isnow), "dd/mm/yyyy")
                    .Update
                End With
            End If
            rsPel.MoveNext
        Loop
    End If
          
'===============================================================================================
'   proses filter data dari kodelgn terbaru
'===============================================================================================
    If Not rsTempCust.EOF Then
        rsTempCust.MoveFirst
        Do While Not rsTempCust.EOF
            rsCust.filter = " KODELGN = '" & NulSID(rsTempCust!kodeLgn) & "'"
            If rsCust.EOF Then
                With rsCust
                    .AddNew
                    !kodeLgn = NulSID(rsTempCust!kodeLgn)
                    !KodeDealer = NulSID(rsTempCust!KodeDealer)
                    !noRangka = NulSID(rsTempCust!noRangka)
                    !Tanggal = Format(NulSID(rsTempCust!Tanggal), "dd/mm/yyyy")
                    .Update
                End With
                rsCust.MoveFirst
            Else
                rsCust.MoveFirst
                If DateDiff("d", Format(NulSID(rsCust!Tanggal), "dd/mm/yyyy"), Format(NulSID(rsTempCust!Tanggal), "dd/mm/yyyy")) > 0 Then
                    With rsCust
                        !kodeLgn = NulSID(rsTempCust!kodeLgn)
                        !KodeDealer = NulSID(rsTempCust!KodeDealer)
                        !noRangka = NulSID(rsTempCust!noRangka)
                        !Tanggal = Format(NulSID(rsTempCust!Tanggal), "dd/mm/yyyy")
                        .Update
                    End With
                End If
            End If
            rsCust.filter = ""
            rsTempCust.MoveNext
        Loop
    End If
       
    '===============================================================================================
    '   MENCOCOKAN DATA KODELGN DENGAN DATA BRIDGING
    '===============================================================================================
    If Not rsCust.EOF Then
        rsCust.MoveFirst
        Do While Not rsCust.EOF
            
            If rsTemp.State Then rsTemp.Close
            sSQL = "select DISTINCT b.[IDCustSales],b.[IDCustService]"
            sSQL = sSQL & " ,b.[NamaCust1STNK],b.[NamaCust2STNK],b.[NamaCust3STNK],b.[tglLahirSTNK],b.[AlamatSTNK],b.[kotaSTNK],b.[HP1STNK],b.[noKTPSTNK],b.[noNPWPSTNK],b.[NoRangka]"
            sSQL = sSQL & " ,b.[NamaCust1SL],b.[NamaCust2SL],b.[NamaCust3SL],b.[tglLahirSL],b.[AlamatSL],b.[kotaSL],b.[HP1SL],b.[noKTPSL],b.[noNPWPSL]"
            sSQL = sSQL & " ,b.[NamaCust1SV],b.[NamaCust2SV],b.[NamaCust3SV],b.[tglLahirSV],b.[AlamatSV],b.[kotaSV],b.[HP1SV],b.[noKTPSV],b.[noNPWPSV]"
            sSQL = sSQL & " ,b.[NOPOLISI]"
            sSQL = sSQL & " from Bridging a"
            sSQL = sSQL & " inner join UnitCustomer b on a.IDCust = b.IDCustSales OR a.IDCust = b.IDCustService"
            sSQL = sSQL & " where a.kodeLgn = '" & NulSID(rsCust!kodeLgn) & "' and a.kodeDealer = '" & Right(rsCust!KodeDealer, 2) & "'"
            rsTemp.Open sSQL, SQLConSID, adOpenStatic, adLockBatchOptimistic, adCmdText
            
            If Not rsTemp.EOF Then
                With rsGrid
                    .AddNew
                    !KodeCustomer = NulSID(rsTemp!idcustsales)
                    !namacustomer = NulSID(rsTemp!NamaCust1Stnk) & " " & NulSID(rsTemp!NamaCust2STNK) & " " & NulSID(rsTemp!NamaCust3STNK)
                    !tglLahir = TglNul(NulSID(rsTemp!tglLahirSTNK))
                    !nopol = NulSID(rsTemp!NoPolisi)
                    !Alamat = NulSID(rsTemp!AlamatSTNK)
                    !Kota = NulSID(rsTemp!Kotastnk)
                    !HP = NulSID(rsTemp!HP1Stnk)
                    !Identitas = NulSID(rsTemp!noKTPSTNK) & " / " & NulSID(rsTemp!noNPWPSTNK)
                    !noRangka = NulSID(rsTemp!noRangka)
                    !Action = "Detail"
                    !Action2 = "Proses"
                    .Update
                End With
            End If
            rsCust.MoveNext
        Loop
    End If
    On Error GoTo 0
    Exit Sub
pencarianDataLanjut_Error:
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure pencarianDataLanjut of Form frmSIDSearchCust"
End Sub

Private Sub Grid_ButtonClick(ByVal ColIndex As Integer)
    If Log.monitoringScoring = 1 And ColIndex = 42 Then
        kodeLgnLamaSplit = NulSID(rsGrid!kodeLgn)
        namaCustomerSplit = NulSID(rsGrid!namaCust)
        alamatSplit = NulSID(rsGrid!Alamat)
        noKTPSplit = NulSID(rsGrid!noKTP)
        noNPWPSplit = NulSID(rsGrid!noNPWP)
        NoHPSplit = NulSID(rsGrid!HP)
        cmdBuatIDBaru.Visible = True
    Else
        If Log.monitoringScoring = 1 Then
            If NulSID(rsGrid!statusUnit) <> NulSID("Active") And ColIndex = 1 Then
                MsgBox "Unit telah di jual / pindah tangan." & vbCrLf & "Tidak dapat diproses", vbInformation
                Exit Sub
            End If
        End If
        Call doClick(ColIndex)
    End If
End Sub

Private Sub doClick(Index As Integer)
          Dim kodeLgnLama     As String
          Dim rsDataLama      As ADODB.Recordset
          Dim rsDataDoClick   As ADODB.Recordset
          Dim Nama()          As String
          Dim i As Integer
1         If Index = 0 Then
lanjutEditData:
2             If NulSID(rsGrid.Fields("IdCust").value) = "" Then
3                 MsgBox "Data masih Id Customer kosong, tidak dapat melihat data", vbInformation
4                 Exit Sub
5             Else
6                 Set rsDataLama = objO.FetchDataACC("SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & NulSID(rsGrid.Fields("IdCust").value) & "' AND AFFILIASI = 1")
7                 If Not rsDataLama.EOF Then
8                   MsgBox "Tipe customer affiliasi tidak dapat diubah", vbInformation
9                   Exit Sub
10                End If
11                kodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
12                sSQL = " select KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2,Alamat3,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry from Armst where KodeLgn = '" & NulSID(kodeLgnLama) & "'"
13                Set rsDataLama = objO.FetchDataACC(sSQL)
                  
14                If Log.monitoringScoring = 0 Then
15                    With frmSIDMasterPelangganNonMerge
16                        .setNoRangka = NulSID(rsGrid.Fields("norangka").value)
17                        .txt(1) = NulSID(rsGrid.Fields("IdCust").value)
18                            .isEdit (True)
19                        .setKodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
20                        If Not rsDataLama.EOF Then
21                            Nama() = Split(NulSID(rsDataLama!Namalgn), " ")
22                            For i = 0 To UBound(Nama)
23                                If Nama(i) <> "" And i = 0 Then
24                                    .txt(4) = NulSID(Nama(i))
25                                ElseIf Nama(i) <> "" And i = 1 Then
26                                    .txt(5) = NulSID(Nama(i))
27                                Else
28                                    .txt(6) = .txt(6) & NulSID(Nama(i))
29                                End If
30                            Next
31                            .txt(3) = NulSID(rsDataLama!npwp)
                              
                              
                              ''pecah notelepon
32                            Nama() = Split(NulSID(rsDataLama!telepon), ",")
33                            For i = 0 To UBound(Nama)
34                                If Nama(i) <> "" And i = 0 Then
35                                    .txt(13) = NulSID(Nama(i))
36                                ElseIf Nama(i) <> "" And i = 1 Then
37                                    .txt(14) = NulSID(Nama(i))
38                                Else
39                                    .txt(15) = NulSID(Nama(i))
40                                End If
41                            Next
                              
42                            .txt(8) = NulSID(rsDataLama!alamat1) & " " & NulSID(rsDataLama!alamat2) & " " & NulSID(rsDataLama!alamat3)
                          
43                        End If
44                        .show vbModal
45                    End With
46                Else
                  
47                    With frmSIDMasterPelanggan
48                        .setNoRangka = NulSID(rsGrid.Fields("norangka").value)
49                        .txt(1) = NulSID(rsGrid.Fields("IdCust").value)
50                            .isEdit (True)
51                        .setKodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
52                        If Not rsDataLama.EOF Then
53                            Nama() = Split(NulSID(rsDataLama!Namalgn), " ")
54                            For i = 0 To UBound(Nama)
55                                If Nama(i) <> "" And i = 0 Then
56                                    .txt(4) = NulSID(Nama(i))
57                                ElseIf Nama(i) <> "" And i = 1 Then
58                                    .txt(5) = NulSID(Nama(i))
59                                Else
60                                    .txt(6) = .txt(6) & NulSID(Nama(i))
61                                End If
62                            Next
63                            .txt(3) = NulSID(rsDataLama!npwp)
                              
                              
                              ''pecah notelepon
64                            Nama() = Split(NulSID(rsDataLama!telepon), ",")
65                            For i = 0 To UBound(Nama)
66                                If Nama(i) <> "" And i = 0 Then
67                                    .txt(13) = NulSID(Nama(i))
68                                ElseIf Nama(i) <> "" And i = 1 Then
69                                    .txt(14) = NulSID(Nama(i))
70                                Else
71                                    .txt(15) = NulSID(Nama(i))
72                                End If
73                            Next
                              
74                            .txt(8) = NulSID(rsDataLama!alamat1) & " " & NulSID(rsDataLama!alamat2) & " " & NulSID(rsDataLama!alamat3)
                          
75                        End If
76                        .show vbModal
77                    End With
78                End If
79            End If
              
80            Exit Sub
81        End If
82        If rsGrid.Fields("IdCust").value <> "" Then
              ''jika idcust belum ada
              ''cek ke data bridging apakah ada kodelgn lama
              
83            If Log.monitoringScoring = 1 Then
                  ''proses cek data kosong untuk id yang terpilih
                  Dim isBoleh As Boolean
84                isBoleh = True
85                sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & rsGrid.Fields("IdCust").value & "' "
86                sSQL = sSQL & " and (isnull(replace(NoKTP,'00000000000000',''),'') = '' or (isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '' or isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '1900-01-01' or isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '3000-01-01') or isnull(replace(NoNPWP,'00.000.000.0-000.000',''),'') = ''"
87                sSQL = sSQL & " or ISNULL(Agama,'') = '' or ISNULL(Pekerjaan,'') = '' or ISNULL(hobi,'') = '' or ISNULL(Olahraga,'') = '' or ISNULL(Propinsi,'') = ''"
88                sSQL = sSQL & " or ISNULL(kota,'') = '' or ISNULL(Kecamatan,'') = '' or ISNULL(Kelurahan,'') = '' or ISNULL(Alamat,'') = '' or ISNULL(email,'') = ''"
89                sSQL = sSQL & " or ISNULL(telp+HP1+hp2,'') = '')"
90                sSQL = sSQL & " and TypeCustomer <> 'PERUSAHAAN UMUM'"
91                Set rsDataDoClick = objO.FetchData(sSQL)
92                If Not rsDataDoClick.EOF Then
93                    isBoleh = False
94                End If
                      
95                If isBoleh = False Then
96                    If MsgBox("Data Customer Tidak Lengkap" & vbCrLf & "Pilih Yes Untuk Melengkapi Data" & vbCrLf & "Pilih No Untuk Skip", vbYesNo, "") = vbYes Then
97                        GoTo lanjutEditData
98                    Else
99                        MsgBox "Data harus dilengkapi sebelum validasi order.", vbInformation
100                   End If
101               End If
102           End If
103           sSQL = "SELECT TypeCustomer FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & rsGrid.Fields("IdCust").value & "' "
104           If NulSID(objO.FetchData_String(sSQL)) = "PERORANGAN" Then
105               If Len(NulSID(Text(9).Text)) < 16 Then
106                   MsgBox "NIK tidak valid. Silahkan hubungi PIC TAX Cabang.", vbInformation
107                   Exit Sub
108               End If
109           End If
              
              
              
110           If gSearchOption = "frmSIDCustomerSurat" Then
111               frmSIDCustomer.Text(66) = NulSID(rsGrid.Fields("IdCust").value)
112               gMainKeyCari = ""
113               gBantuan = ""
114               Unload Me
115               Exit Sub
116           End If
117           If NulSID(rsGrid.Fields("norangka").value) = "" Then
                  ''jika norangka kosong maka di arahkan ke penambahan norangka
118               If frmSIDCustomer.Visible = True Then
                      ''
119                   Nama() = Split(NulSID(rsGrid.Fields("namaCust").value), " ")
                      '.isDariSearch = True
120                   With frmSIDCustomer
121                       .Text(2) = NulSID(rsGrid.Fields("IdCust").value)
122                       .Text(65) = NulSID(rsGrid.Fields("KodeLgn").value)
123                       For i = 0 To UBound(Nama)
124                           If Nama(i) <> "" And i = 0 Then
125                               .Text(4) = NulSID(Nama(i))
126                           ElseIf Nama(i) <> "" And i = 1 Then
127                               .Text(60) = NulSID(Nama(i))
128                           Else
129                               .Text(61) = .txt(6) & NulSID(Nama(i))
130                           End If
131                       Next
                          '.Text(0).SetFocus
132                   End With
133                   Unload Me
134               End If
135               gMainKeyCari = NulSID(rsGrid.Fields("IdCust").value)
136               gBantuan = NulSID(rsGrid.Fields("norangka").value)
137               Unload Me
138               Exit Sub
139           Else
              ''jika norangka sudah ada isinya dikembalikan ke form pemanggil
                  
140               If frmSIDCustomer.Visible = True Then
                      ''
141                   Nama() = Split(NulSID(rsGrid.Fields("namaCust").value), " ")
                      
142                   With frmSIDCustomer
143                       .isDariSearch = True
                          '.text(0) = Right(nulsid(rsGrid.Fields("norangka").value), 8)
144                       If NulSID(Replace(Trim(.Text(50)), ">", "") & Replace(Replace(.Text(1), ">", ""), " ", "") & Replace(Replace(.Text(0), ">", ""), " ", "")) <> NulSID(rsGrid.Fields("norangka").value) And NulSID(Replace(Trim(.Text(50)), ">", "") & Replace(Replace(.Text(1), ">", ""), " ", "") & Replace(Replace(.Text(0), ">", ""), " ", "")) <> "" Then
145                           If MsgBox("No rangka yang dipilih berbeda dengan data customer." & vbCrLf & " Apakah ingin mengikuti perubahan?", vbYesNo) = vbYes Then
146                               .Text(0) = Right(NulSID(rsGrid.Fields("norangka").value), 8)
147                               .Text(1) = Mid(NulSID(rsGrid.Fields("norangka").value), 4, Len(NulSID(rsGrid.Fields("norangka").value)) - 11)
148                               .Text(50) = Left(NulSID(rsGrid.Fields("norangka").value), 3)
149                           End If
150                       Else
151                           .Text(0) = Right(NulSID(rsGrid.Fields("norangka").value), 8)
152                           .Text(1) = Mid(NulSID(rsGrid.Fields("norangka").value), 4, Len(NulSID(rsGrid.Fields("norangka").value)) - 11)
153                           .Text(50) = Left(NulSID(rsGrid.Fields("norangka").value), 3)
154                       End If
155                       .Text(2) = NulSID(rsGrid.Fields("IdCust").value)
156                       .Text(65) = NulSID(rsGrid.Fields("KodeLgn").value)
157                       For i = 0 To UBound(Nama)
158                           If Nama(i) <> "" And i = 0 Then
159                               .Text(4) = NulSID(Nama(i))
160                           ElseIf Nama(i) <> "" And i = 1 Then
161                               .Text(60) = NulSID(Nama(i))
162                           Else
163                               .Text(61) = .txt(6) & NulSID(Nama(i))
164                           End If
165                       Next
                          '.Text(0).SetFocus
166                   End With
167                   Unload Me
168               End If
169               gMainKeyCari = NulSID(rsGrid.Fields("IdCust").value)
170               gBantuan = NulSID(rsGrid.Fields("norangka").value)
171               Unload Me
172               Exit Sub
173           End If
              
174       Else
              ''jika belum ada idcust di arahkan ke tambah master
175           MsgBox "Data belum ada di Customer ID" & vbCrLf & "Silahkan untuk menambahkan data", vbInformation
              
176           kodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
177           sSQL = " select KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2,Alamat3,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry from Armst where KodeLgn = '" & NulSID(kodeLgnLama) & "'"
178           Set rsDataLama = objO.FetchDataACC(sSQL)
              
179           sSQL = "select namacust1sv,namacust2sv,namacust3sv,tgllahirsv,agamasv,pekerjaansv,propinsisv,kotasv,kecamatansv,kelurahansv,kodepossv,alamatsv,rtsv,rwsv,hp1sv"
180           sSQL = sSQL & " From [SINGLEIDORI]..UnitCustomer"
181           sSQL = sSQL & " where NoRangka = '" & NulSID(rsGrid.Fields("norangka").value) & "'"
182           sSQL = sSQL & " and ltrim(rtrim(NamaCust1SV + ' ' + NamaCust2SV + ' ' + NamaCust3SV)) = '" & NulSID(rsGrid.Fields("namaCust").value) & "'"
183           sSQL = sSQL & " and HP1SV = '" & NulSID(rsGrid.Fields("hp").value) & "'"
184           sSQL = sSQL & " Union"
185           sSQL = sSQL & " select NamaCust1SL,namacust2sl,namacust3sl,tgllahirsl,agamasl,pekerjaansl,propinsisl,kotasl,kecamatansl,kelurahansl,kodepossl,alamatsl,rtsl,rwsl,hp1sl"
186           sSQL = sSQL & " From [SINGLEIDORI]..UnitCustomer"
187           sSQL = sSQL & " where NoRangka = '" & NulSID(rsGrid.Fields("norangka").value) & "'"
188           sSQL = sSQL & " and ltrim(rtrim(NamaCust1Sl + ' ' + NamaCust2Sl + ' ' + NamaCust3Sl)) = '" & NulSID(rsGrid.Fields("namaCust").value) & "'"
189           sSQL = sSQL & " and HP1Sl = '" & NulSID(rsGrid.Fields("hp").value) & "'"
190           sSQL = sSQL & " Union"
191           sSQL = sSQL & " select NamaCust1Stnk,namacust2stnk,namacust3stnk,tgllahirstnk,agamastnk,pekerjaanstnk,propinsistnk,kotastnk,kecamatanstnk,kelurahanstnk,kodeposstnk,alamatstnk,rtstnk,rwstnk,hp1stnk"
192           sSQL = sSQL & " From [SINGLEIDORI]..UnitCustomer"
193           sSQL = sSQL & " where NoRangka = '" & NulSID(rsGrid.Fields("norangka").value) & "'"
194           sSQL = sSQL & " and ltrim(rtrim(NamaCust1Stnk + ' ' + NamaCust2Stnk + ' ' + NamaCust3Stnk)) = '" & NulSID(rsGrid.Fields("namaCust").value) & "'"
195           sSQL = sSQL & " and HP1Stnk = '" & NulSID(rsGrid.Fields("hp").value) & "'"
196           Set rsDataLama = objO.FetchData(sSQL)
              
197           If Log.monitoringScoring = 0 Then
198               With frmSIDMasterPelangganNonMerge
199                   .setKodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
200                     .isEdit (False)
201                   If Not rsDataLama.EOF Then
202                       rsDataLama.MoveFirst
          '                Nama() = Split(nulsid(rsDataLama!Namalgn), " ")
          '                For i = 0 To UBound(Nama)
          '                    If Nama(i) <> "" And i = 0 Then
          '                        .Txt(4) = nulsid(Nama(i))
          '                    ElseIf Nama(i) <> "" And i = 1 Then
          '                        .Txt(5) = nulsid(Nama(i))
          '                    Else
          '                        .Txt(6) = .Txt(6) & nulsid(Nama(i))
          '                    End If
          '                Next
          '                .Txt(3) = nulsid(rsDataLama!npwp)
                          
203                       .txt(4) = NulSID(rsDataLama!NamaCust1SV)
204                       .txt(5) = NulSID(rsDataLama!NamaCust2SV)
205                       .txt(6) = NulSID(rsDataLama!NamaCust3SV)
                          
206                       .txt(14) = NulSID(rsDataLama!HP1SV)
                          
207                       .txt(11) = NulSID(rsDataLama!HP1SV)
                          
208                       .txt(8) = NulSID(rsDataLama!AlamatSV)
209                       .DTLahir.value = Format(NulSID(rsDataLama!tglLahirsv), "dd/mm/yyyy")
210                       .txt(11) = NulSID(rsDataLama!KodePosSV)
211                       .txt(63) = NulSID(rsDataLama!PropinsiSV)
212                       .txt(64) = NulSID(rsDataLama!KotaSV)
213                       .txt(65) = NulSID(rsDataLama!KecamatanSV)
214                       .txt(66) = NulSID(rsDataLama!KelurahanSV)
                          
215                       .txt(9) = NulSID(rsDataLama!RTSV)
216                       .txt(10) = NulSID(rsDataLama!RWSV)
                          
217                       .txt(20) = NulSID(rsDataLama!Agamasv)
                          
                          ''pecah notelepon
          '                Nama() = Split(nulsid(rsDataLama!telepon), ",")
          '                For i = 0 To UBound(Nama)
          '                    If Nama(i) <> "" And i = 0 Then
          '                        .Txt(13) = nulsid(Nama(i))
          '                    ElseIf Nama(i) <> "" And i = 1 Then
          '                        .Txt(14) = nulsid(Nama(i))
          '                    Else
          '                        .Txt(15) = nulsid(Nama(i))
          '                    End If
          '                Next
          '
          '                .Txt(8) = nulsid(rsDataLama!alamat1) & " " & nulsid(rsDataLama!alamat2) & " " & nulsid(rsDataLama!alamat3)
218                   Else
219                       Nama() = Split(NulSID(rsGrid.Fields("namaCust").value), " ")
220                       For i = 0 To UBound(Nama)
221                           If Nama(i) <> "" And i = 0 Then
222                               .txt(4) = NulSID(Nama(i))
223                           ElseIf Nama(i) <> "" And i = 1 Then
224                               .txt(5) = NulSID(Nama(i))
225                           Else
226                               .txt(6) = .txt(6) & NulSID(Nama(i))
227                           End If
228                       Next
                          '.txt(3) = nulsid(rsGrid.Fields("namaCust").value)
229                       .txt(8) = NulSID(rsGrid.Fields("Alamat").value)
230                       .txt(64) = NulSID(rsGrid.Fields("kota").value)
231                   End If
232                   .show vbModal
233               End With
              
234           Else
              
235               With frmSIDMasterPelanggan
236                   .setKodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
237                     .isEdit (False)
238                   If Not rsDataLama.EOF Then
239                       rsDataLama.MoveFirst
          '                Nama() = Split(nulsid(rsDataLama!Namalgn), " ")
          '                For i = 0 To UBound(Nama)
          '                    If Nama(i) <> "" And i = 0 Then
          '                        .Txt(4) = nulsid(Nama(i))
          '                    ElseIf Nama(i) <> "" And i = 1 Then
          '                        .Txt(5) = nulsid(Nama(i))
          '                    Else
          '                        .Txt(6) = .Txt(6) & nulsid(Nama(i))
          '                    End If
          '                Next
          '                .Txt(3) = nulsid(rsDataLama!npwp)
                          
240                       .txt(4) = NulSID(rsDataLama!NamaCust1SV)
241                       .txt(5) = NulSID(rsDataLama!NamaCust2SV)
242                       .txt(6) = NulSID(rsDataLama!NamaCust3SV)
                          
243                       .txt(14) = NulSID(rsDataLama!HP1SV)
                          
244                       .txt(11) = NulSID(rsDataLama!HP1SV)
                          
245                       .txt(8) = NulSID(rsDataLama!AlamatSV)
246                       .DTLahir.value = Format(NulSID(rsDataLama!tglLahirsv), "dd/mm/yyyy")
247                       .txt(11) = NulSID(rsDataLama!KodePosSV)
248                       .txt(63) = NulSID(rsDataLama!PropinsiSV)
249                       .txt(64) = NulSID(rsDataLama!KotaSV)
250                       .txt(65) = NulSID(rsDataLama!KecamatanSV)
251                       .txt(66) = NulSID(rsDataLama!KelurahanSV)
                          
252                       .txt(9) = NulSID(rsDataLama!RTSV)
253                       .txt(10) = NulSID(rsDataLama!RWSV)
                          
254                       .txt(20) = NulSID(rsDataLama!Agamasv)
                          
                          ''pecah notelepon
          '                Nama() = Split(nulsid(rsDataLama!telepon), ",")
          '                For i = 0 To UBound(Nama)
          '                    If Nama(i) <> "" And i = 0 Then
          '                        .Txt(13) = nulsid(Nama(i))
          '                    ElseIf Nama(i) <> "" And i = 1 Then
          '                        .Txt(14) = nulsid(Nama(i))
          '                    Else
          '                        .Txt(15) = nulsid(Nama(i))
          '                    End If
          '                Next
          '
          '                .Txt(8) = nulsid(rsDataLama!alamat1) & " " & nulsid(rsDataLama!alamat2) & " " & nulsid(rsDataLama!alamat3)
255                   Else
256                       Nama() = Split(NulSID(rsGrid.Fields("namaCust").value), " ")
257                       For i = 0 To UBound(Nama)
258                           If Nama(i) <> "" And i = 0 Then
259                               .txt(4) = NulSID(Nama(i))
260                           ElseIf Nama(i) <> "" And i = 1 Then
261                               .txt(5) = NulSID(Nama(i))
262                           Else
263                               .txt(6) = .txt(6) & NulSID(Nama(i))
264                           End If
265                       Next
                          '.txt(3) = nulsid(rsGrid.Fields("namaCust").value)
266                       .txt(8) = NulSID(rsGrid.Fields("Alamat").value)
267                       .txt(64) = NulSID(rsGrid.Fields("kota").value)
268                   End If
269                   .show vbModal
270               End With
271           End If
272           Call PencarianData
              'Grid.SetFocus
273       End If
          'gMainKeyCari = ""
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

Private Sub grid_FilterChange()
        On Error GoTo errHandler
    Set cols = Grid.Columns
    Dim C As Integer
    
    C = Grid.col
    Grid.HoldFields
    rsGrid.filter = getFilter()
    If rsGrid.EOF Then
        For Each col In Grid.Columns
            col.FilterText = ""
        Next col
        rsGrid.filter = ""
    End If
    Grid.col = C
    Grid.EditActive = True
    
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    End If
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If rsGrid.RecordCount > 0 Then
        Text(9).Text = NulSID(rsGrid!noKTP) 'Mid(NulSID(rsGrid("NoID")), 1, InStr(NulSID(rsGrid("NoID")), " - "))
        Text(8).Text = NulSID(rsGrid!noNPWP) ' Mid(NulSID(rsGrid("NoID")), InStr(NulSID(rsGrid("NoID")), " - ") + 3)
        Text(7).Text = NulSID(rsGrid("IDCust"))
        Text(6).Text = NulSID(rsGrid("NamaCust"))
        tglLahirx.value = rsGrid("TglLahir")
        Text(5).Text = NulSID(rsGrid("Alamat"))
        Text(11).Text = NulSID(rsGrid("hp"))
    Else
        Text(9).Text = ""
        Text(8).Text = ""
        Text(7).Text = ""
        Text(6).Text = ""
        tglLahirx.value = "__/__/____"
        Text(5).Text = ""
        Text(11).Text = ""
    End If
End Sub
 

Private Function cekKoneksiBP() As Boolean
    Dim hasil       As Boolean
    Dim ip          As String
    Dim database    As String
   On Error GoTo cekKoneksiBP_Error

    hasil = False
    ip = ""
    database = ""
    If rsTemp.State Then rsTemp.Close
    sSQL = "SELECT SiDbBP, isnull(BPServerIP,'') as IP,SiBklDual FROM profilacc..profil"
    rsTemp.Open sSQL, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        If NulSID(rsTemp!SiBklDual) = 1 Then
            ip = NulSID(rsTemp!ip)
        Else
            ip = NulSID(Log.LogServ)
        End If
        
        database = NulSID(rsTemp!SiDbBP)

        If UCase(database) = "BP" Or NulSID(database) = "" Then
            hasil = False
        Else
            SQLConBP = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog=" & database & " ;server=" & ip
            If rsTemp.State Then rsTemp.Close
            sSQL = "SELECT TOP 1 * FROM custdat"
            rsTemp.Open sSQL, SQLConBP, adOpenStatic, adLockBatchOptimistic, adCmdText
            hasil = True
            Log.logServBP = ip
            Log.logDbBP = database
            
        End If
    End If
    cekKoneksiBP = hasil

   On Error GoTo 0
   Exit Function

cekKoneksiBP_Error:
    cekKoneksiBP = False
End Function

Private Sub gridDuplikat_AfterColUpdate(ByVal ColIndex As Integer)
1         Select Case ColIndex
              Case 0
2                 If rsDuplikat!IDMerge Then
3                     idCustTerpilih = NulSID(rsDuplikat!idcust)
                      ''rsDuplikat.filter = "idMerge = '" & True & "' and idcust <> " & idCustTerpilih
4                     rsDuplikat.filter = "idcust <> " & idCustTerpilih
5                     kolomTerpilih = "idMerge"
6                     GoSub gantiKolomLain
7                     rsDuplikat.filter = "" '
8                     gridDuplikat.ReBind
9                 End If
                  
10            Case 3
11                If rsDuplikat!pilihnama Then
12                    If rsDuplikat!pilihMerge = True Then
13                        idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
14                        kolomTerpilih = "pilihnama"
15                        GoSub gantiKolomLainFind 'gantiKolomLain
16                    Else
17                        rsDuplikat!pilihnama = False
18                        gridDuplikat.Columns(ColIndex).value = False
19                        MsgBox "Pilih id merge terlebih dahulu", vbInformation
20                    End If
                      'rsDuplikat.filter = "" '
21                    gridDuplikat.Refresh
22                End If
              
23            Case 5
24                If rsDuplikat!pilihTglLahir Then
25                    If rsDuplikat!pilihMerge = True Then
26                        idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
27                        kolomTerpilih = "pilihTglLahir"
28                        GoSub gantiKolomLainFind 'gantiKolomLain
29                    Else
30                        rsDuplikat!pilihTglLahir = False
31                        gridDuplikat.Columns(ColIndex).value = False
32                        MsgBox "Pilih id merge terlebih dahulu", vbInformation
33                    End If
                      'rsDuplikat.filter = "" '
34                    gridDuplikat.Refresh
35                End If
              
36            Case 7
37                If rsDuplikat!pilihAgama Then
38                    If rsDuplikat!pilihMerge = True Then
39                        idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
40                        kolomTerpilih = "pilihAgama"
41                        GoSub gantiKolomLainFind 'gantiKolomLain
42                    Else
43                        rsDuplikat!pilihAgama = False
44                        gridDuplikat.Columns(ColIndex).value = False
45                        MsgBox "Pilih id merge terlebih dahulu", vbInformation
46                    End If
                      'rsDuplikat.filter = "" '
47                    gridDuplikat.Refresh
48                End If
49            Case 9
50                If rsDuplikat!pilihPekerjaan Then
51                    If rsDuplikat!pilihMerge = True Then
52                        idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
53                        kolomTerpilih = "pilihPekerjaan"
54                        GoSub gantiKolomLainFind 'gantiKolomLain
55                    Else
56                        rsDuplikat!pilihPekerjaan = False
57                        gridDuplikat.Columns(ColIndex).value = False
58                        MsgBox "Pilih id merge terlebih dahulu", vbInformation
59                    End If
                      'rsDuplikat.filter = "" '
60                    gridDuplikat.Refresh
61                End If
62            Case 11
63                If rsDuplikat!pilihHobi Then
64                    If rsDuplikat!pilihMerge = True Then
65                        idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
66                        kolomTerpilih = "pilihHobi"
67                        GoSub gantiKolomLainFind 'gantiKolomLain
68                    Else
69                        rsDuplikat!pilihHobi = False
70                        gridDuplikat.Columns(ColIndex).value = False
71                        MsgBox "Pilih id merge terlebih dahulu", vbInformation
72                    End If
73                    rsDuplikat.filter = "" '
                      'gridDuplikat.Refresh
74                End If
75            Case 13
76                If rsDuplikat!pilihOlahraga Then
77                    If rsDuplikat!pilihMerge = True Then
78                        idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
79                        kolomTerpilih = "pilihOlahraga"
80                        GoSub gantiKolomLainFind 'gantiKolomLain
81                    Else
82                        rsDuplikat!pilihOlahraga = False
83                        gridDuplikat.Columns(ColIndex).value = False
84                        MsgBox "Pilih id merge terlebih dahulu", vbInformation
85                    End If
                      'rsDuplikat.filter = "" '
86                    gridDuplikat.Refresh
87                End If
88            Case 15
89                If rsDuplikat!pilihPropinsi Then
90                    If rsDuplikat!pilihMerge = True Then
91                        idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
92                        kolomTerpilih = "pilihPropinsi"
93                        GoSub gantiKolomLainFind 'gantiKolomLain
94                    Else
95                        rsDuplikat!pilihPropinsi = False
96                        gridDuplikat.Columns(ColIndex).value = False
97                        MsgBox "Pilih id merge terlebih dahulu", vbInformation
98                    End If
                      'rsDuplikat.filter = "" '
99                    gridDuplikat.Refresh
100               End If
101           Case 17
102               If rsDuplikat!pilihKota Then
103                   If rsDuplikat!pilihMerge = True Then
104                       idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
105                       kolomTerpilih = "pilihKota"
106                       GoSub gantiKolomLainFind 'gantiKolomLain
107                   Else
108                       rsDuplikat!pilihKota = False
109                       gridDuplikat.Columns(ColIndex).value = False
110                       MsgBox "Pilih id merge terlebih dahulu", vbInformation
111                   End If
                      'rsDuplikat.filter = "" '
112                   gridDuplikat.Refresh
113               End If
114           Case 19
115               If rsDuplikat!pilihKecamatan Then
116                   If rsDuplikat!pilihMerge = True Then
117                       idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
118                       kolomTerpilih = "pilihKecamatan"
119                       GoSub gantiKolomLainFind 'gantiKolomLain
120                   Else
121                       rsDuplikat!pilihKecamatan = False
122                       gridDuplikat.Columns(ColIndex).value = False
123                       MsgBox "Pilih id merge terlebih dahulu", vbInformation
124                   End If
                      'rsDuplikat.filter = "" '
125                   gridDuplikat.Refresh
126               End If
127           Case 21
128               If rsDuplikat!pilihKeLurahan Then
129                   If rsDuplikat!pilihMerge = True Then
130                       idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
131                       kolomTerpilih = "pilihKelurahan"
132                       GoSub gantiKolomLainFind 'gantiKolomLain
133                   Else
134                       rsDuplikat!pilihKecamatan = False
135                       gridDuplikat.Columns(ColIndex).value = False
136                       MsgBox "Pilih id merge terlebih dahulu", vbInformation
137                   End If
                      'rsDuplikat.filter = "" '
138                   gridDuplikat.Refresh
139               End If
140           Case 23
141               If rsDuplikat!pilihHP Then
142                   If rsDuplikat!pilihMerge = True Then
143                       idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.Find "idcust <> " & idCustTerpilih, , adSearchForward, 1  'rsDuplikat.filter = "idcust <> " & idCustTerpilih
144                       kolomTerpilih = "pilihHP"
145                       GoSub gantiKolomLainFind
146                   Else
147                       rsDuplikat!pilihHP = False
148                       gridDuplikat.Columns(ColIndex).value = False
149                       MsgBox "Pilih id merge terlebih dahulu", vbInformation
150                   End If
                      'rsDuplikat.filter = "" '
151                   gridDuplikat.Refresh
152               End If
153           Case 25
154               If rsDuplikat!pilihKTP Then
155                   If rsDuplikat!pilihMerge = True Then
156                       idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
157                       kolomTerpilih = "pilihKTP"
158                       GoSub gantiKolomLainFind 'gantiKolomLain
159                   Else
160                       rsDuplikat!pilihKTP = False
161                       gridDuplikat.Columns(ColIndex).value = False
162                       MsgBox "Pilih id merge terlebih dahulu", vbInformation
163                   End If
                      'rsDuplikat.filter = "" '
164                   gridDuplikat.Refresh
165               End If
166           Case 27
167               If rsDuplikat!pilihNPWP Then
168                   If rsDuplikat!pilihMerge = True Then
169                       idCustTerpilih = NulSID(rsDuplikat!idcust)
                          'rsDuplikat.filter = "idcust <> " & idCustTerpilih
170                       kolomTerpilih = "pilihNPWP"
171                       GoSub gantiKolomLainFind 'gantiKolomLain
172                   Else
173                       rsDuplikat!pilihNPWP = False
174                       gridDuplikat.Columns(ColIndex).value = False
175                       MsgBox "Pilih id merge terlebih dahulu", vbInformation
176                   End If
                      'rsDuplikat.filter = "" '
177                   gridDuplikat.Refresh
178               End If
179       End Select
180       Exit Sub
gantiKolomLainFind:
181       If Not rsDuplikat.EOF Then
182           rsDuplikat.MoveFirst
183           Do While Not rsDuplikat.EOF
184               If rsDuplikat!idcust <> idCustTerpilih Then
185                   rsDuplikat(kolomTerpilih) = False
186               End If
187               rsDuplikat.Update
188               rsDuplikat.MoveNext
189           Loop
190       End If
191       Return
gantiKolomLain:
192       If Not rsDuplikat.EOF Then
193           Do While Not rsDuplikat.EOF
194               rsDuplikat(kolomTerpilih) = False
195               rsDuplikat.Update
196               rsDuplikat.MoveNext
197           Loop
198       End If
199       Return
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If Index = 3 Then
'            tglLahir.SetFocus
'        ElseIf Index = 4 Then
'            Call btnCari_Click
'        Else
'            If Text(Index).BackColor = fColor Then
'                Text(Index).Locked = True
'            Else
'                Text(Index).Locked = False
'                SendKeys ("{TAB}")
'            End If
'        End If
        Call btnCari_Click
    End If
    KeyAscii = UpCase(KeyAscii)
    
End Sub
 

Private Sub createTempCustomer()
1       On Error GoTo createTempCustomer_Error
2       If rsTempCust.State Then rsTempCust.Close
3       With rsTempCust
4           .Fields.Append "KodeLgn", adVarChar, 13
5           .Fields.Append "KodeDealer", adVarChar, 13
6           .Fields.Append "NoRangka", adVarChar, 20
7           .Fields.Append "tanggal", adDate
8           .Open
9       End With
10      On Error GoTo 0
11      Exit Sub
createTempCustomer_Error:
12      MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createTempCustomer of Form frmSIDSearchCust"
End Sub

Private Sub createCustomer()
1       On Error GoTo createCustomer_Error
2       If rsCust.State Then rsCust.Close
3       Set rsCust = Nothing
4       With rsCust
5           .Fields.Append "KodeLgn", adVarChar, 13
6           .Fields.Append "KodeDealer", adVarChar, 13
7           .Fields.Append "NoRangka", adVarChar, 20
8           .Fields.Append "tanggal", adDate
9           .Open
10      End With
11      On Error GoTo 0
12      Exit Sub
createCustomer_Error:
13    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createCustomer of Form frmSIDSearchCust"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : createRsGrid
' DateTime  : 09/09/2020 10:56
' Author    : Win7
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createRsGrid()
1        On Error GoTo createRsGrid_Error

2       On Error GoTo createRsGrid_Error
3       If rsGrid.State Then rsGrid.Close
4       Set rsGrid = Nothing
5       With rsGrid
6           .Fields.Append "Detail", adVarChar, 10 'adBoolean
7           .Fields.Append "Proses", adVarChar, 10
8           .Fields.Append "IDCust", adVarChar, 13
9           .Fields.Append "NamaCust", adVarChar, 210
10          .Fields.Append "tglLahir", adVarChar, 10
11          .Fields.Append "NoPolisi", adVarChar, 12
12          .Fields.Append "Alamat", adVarChar, 350
13          .Fields.Append "Kota", adVarChar, 50
14          .Fields.Append "HP", adVarChar, 20
15          .Fields.Append "NoID", adVarChar, 75
16          .Fields.Append "NoRangka", adVarChar, 20
17          .Fields.Append "kodeLgn", adVarChar, 40
18          .Fields.Append "noNPWP", adVarChar, 75
19          .Fields.Append "NoKTP", adVarChar, 75
      '      .Fields.Append "Alamat", adVarChar, 150
      '      .Fields.Append "Alamat", adVarChar, 150
      '      .Fields.Append "Alamat", adVarChar, 150
      '      .Fields.Append "Alamat", adVarChar, 150
20          If Log.monitoringScoring = 1 Then
              '.Fields.Append "kodelgn", adVarChar, 13
21            .Fields.Append "Split", adVarChar, 10
22            .Fields.Append "StatusUnit", adVarChar, 100
23          End If
24          .Open
25      End With
26      On Error GoTo 0
27      Exit Sub
createRsGrid_Error:
28      MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRsGrid of Form frmSIDSearchCust")

29       On Error GoTo 0
30       Exit Sub

End Sub

Private Sub createHeaderMergeFix2()
    With gridCustomerFix
        .Columns(0).Caption = "ID Customer"
        .Columns(0).Locked = True
        
        .Columns(1).Caption = "Tgl Lahir"
        .Columns(1).Locked = True
        
        .Columns(2).Caption = "Agama"
        .Columns(2).Locked = True
        
        .Columns(3).Caption = "Pekerjaan"
        .Columns(3).Locked = True
        
        .Columns(4).Caption = "Hobi"
        .Columns(4).Locked = True
        
        .Columns(5).Caption = "Olahraga"
        .Columns(5).Locked = True
        
        .Columns(6).Caption = "Alamat"
        .Columns(6).Locked = True
        
        .Columns(7).Caption = "Kota"
        .Columns(7).Locked = True
        
        .Columns(8).Caption = "Propinsi"
        .Columns(8).Locked = True
        
        .Columns(9).Caption = "Kecamatan"
        .Columns(9).Locked = True
        
        .Columns(10).Caption = "Kelurahan"
        .Columns(10).Locked = True
        
        .Columns(11).Caption = "No HP"
        .Columns(11).Locked = True
        
        .Columns(12).Caption = "No NPWP"
        .Columns(12).Locked = True
        
        .Columns(13).Caption = "No KTP"
        .Columns(13).Locked = True
        
    End With
End Sub

Private Sub createRsMergeFix()

   On Error GoTo createRsMergeFix_Error
    
  If rsMergeFix.State Then rsMergeFix.Close
  Set rsMergeFix = Nothing
  With rsMergeFix
      .Fields.Append "IDCust", adVarChar, 13
      .Fields.Append "Nama", adVarChar, 210
      .Fields.Append "tglLahir", adVarChar, 10
      .Fields.Append "Agama", adVarChar, 12
      .Fields.Append "pekerjaan", adVarChar, 50
      .Fields.Append "Hobi", adVarChar, 50
      .Fields.Append "Olahraga", adVarChar, 50
      .Fields.Append "Alamat", adVarChar, 350
      .Fields.Append "Kota", adVarChar, 50
      .Fields.Append "Propinsi", adVarChar, 50
      .Fields.Append "Kecamatan", adVarChar, 50
      .Fields.Append "Kelurahan", adVarChar, 50
      .Fields.Append "HP", adVarChar, 20
      .Fields.Append "noNPWP", adVarChar, 75
      .Fields.Append "NoKTP", adVarChar, 75
      .Open
  End With
   On Error GoTo 0
   Exit Sub

createRsMergeFix_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure createRsMergeFix of Form frmSIDSearchCust"
End Sub

Private Sub createHeaderMergeFix()
    With gridCustomerFix
        .Columns(0).Caption = "ID Customer"
        .Columns(0).DataField = "idcust"
        .Columns(0).Locked = True
        .Columns(0).Width = 1000
        
        .Columns(1).Caption = "Nama Customer"
        .Columns(1).DataField = "Nama"
        .Columns(1).Locked = True
        .Columns(1).Width = 1000
        
        .Columns(2).Caption = "Tgl Lahir"
        .Columns(2).DataField = "tgllahir"
        .Columns(2).Locked = True
        .Columns(2).Width = 1000
         
        .Columns(3).Caption = "Agama"
        .Columns(3).DataField = "Agama"
        .Columns(3).Locked = True
        .Columns(3).Width = 1000
        
        .Columns(4).Caption = "Pekerjaan"
        .Columns(4).DataField = "pekerjaan"
        .Columns(4).Locked = True
        .Columns(4).Width = 1000
        
        .Columns(5).Caption = "Hobi"
        .Columns(5).DataField = "hobi"
        .Columns(5).Locked = True
        .Columns(5).Width = 1000
        
        .Columns(6).Caption = "Olahraga"
        .Columns(6).DataField = "Olahraga"
        .Columns(6).Locked = True
        .Columns(6).Width = 1000
        
        .Columns(7).Caption = "Alamat"
        .Columns(7).DataField = "Alamat"
        .Columns(7).Locked = True
        .Columns(7).Width = 1000
        
        .Columns(8).Caption = "Kota"
        .Columns(8).DataField = "Kota"
        .Columns(8).Locked = True
        .Columns(8).Width = 1000
        
        .Columns(9).Caption = "Propinsi"
        .Columns(9).DataField = "Propinsi"
        .Columns(9).Locked = True
        .Columns(9).Width = 1000
        
        .Columns(10).Caption = "Kota"
        .Columns(10).DataField = "Kota"
        .Columns(10).Locked = True
        .Columns(10).Width = 1000
        
        .Columns(11).Caption = "Kecamatan"
        .Columns(11).DataField = "kecamatan"
        .Columns(11).Locked = True
        .Columns(11).Width = 1000
        
        .Columns(12).Caption = "Kelurahan"
        .Columns(12).DataField = "kelurahan"
        .Columns(12).Locked = True
        .Columns(12).Width = 1000
        
        .Columns(13).Caption = "HP"
        .Columns(13).DataField = "HP"
        .Columns(13).Locked = True
        .Columns(13).Width = 1000
        
        .Columns(14).Caption = "No NPWP"
        .Columns(14).DataField = "noNPWP"
        .Columns(14).Locked = True
        .Columns(14).Width = 1000
        
        .Columns(15).Caption = "No KTP"
        .Columns(15).DataField = "noKTP"
        .Columns(15).Locked = True
        .Columns(15).Width = 1000
                
    End With
End Sub


Private Sub createRsDuplikat()
   On Error GoTo createRsDuplikat_Error

  If rsDuplikat.State Then rsDuplikat.Close
  Set rsDuplikat = Nothing
  With rsDuplikat
      .Fields.Append "idMerge", adBoolean
      .Fields.Append "pilihMerge", adBoolean
      .Fields.Append "IDCust", adVarChar, 13
      .Fields.Append "Nama", adVarChar, 210
      .Fields.Append "pilihNama", adBoolean
      .Fields.Append "tglLahir", adVarChar, 10
      .Fields.Append "pilihTglLahir", adBoolean
      .Fields.Append "Agama", adVarChar, 12
      .Fields.Append "pilihAgama", adBoolean
      .Fields.Append "pekerjaan", adVarChar, 50
      .Fields.Append "pilihPekerjaan", adBoolean
      .Fields.Append "Hobi", adVarChar, 50
      .Fields.Append "pilihHobi", adBoolean
      .Fields.Append "Olahraga", adVarChar, 50
      .Fields.Append "pilihOlahraga", adBoolean
      .Fields.Append "Alamat", adVarChar, 350
      .Fields.Append "pilihAlamat", adBoolean
      .Fields.Append "Kota", adVarChar, 50
      .Fields.Append "pilihKota", adBoolean
      .Fields.Append "Propinsi", adVarChar, 50
      .Fields.Append "pilihPropinsi", adBoolean
      .Fields.Append "Kecamatan", adVarChar, 50
      .Fields.Append "pilihKecamatan", adBoolean
      .Fields.Append "Kelurahan", adVarChar, 50
      .Fields.Append "pilihKelurahan", adBoolean
      .Fields.Append "HP", adVarChar, 20
      .Fields.Append "pilihHP", adBoolean
      .Fields.Append "noNPWP", adVarChar, 75
      .Fields.Append "pilihNPWP", adBoolean
      .Fields.Append "NoKTP", adVarChar, 75
      .Fields.Append "pilihKTP", adBoolean
      .Open
  End With

   On Error GoTo 0
   Exit Sub

createRsDuplikat_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure createRsDuplikat of Form frmSIDSearchCust"


End Sub

Private Sub btnKeluar_Click()
    IsAct = "CLOSE"
    isPassing = False
    gMainKeyCari = ""
    gBantuan = ""
    Unload Me
End Sub

Private Sub btnTambah_Click()
    Dim namaSplit() As String
    namaSplit = Split(NulSID(Text(3)), " ")
    Call clearForm
    If Log.monitoringScoring = 0 Then
        With frmSIDMasterPelangganNonMerge
            .setKodeLgnLama = ""
            For i = 0 To UBound(namaSplit)
                If i = 0 Then
                    .txt(4) = NulSID(namaSplit(i))
                ElseIf i = 1 Then
                    .txt(5) = NulSID(namaSplit(i))
                Else
                    .txt(6) = .txt(6) & NulSID(namaSplit(i)) & " "
                End If
            Next
            .isEdit False
            .show vbModal
        End With
    
    Else
        With frmSIDMasterPelanggan
            .setKodeLgnLama = ""
            For i = 0 To UBound(namaSplit)
                If i = 0 Then
                    .txt(4) = NulSID(namaSplit(i))
                ElseIf i = 1 Then
                    .txt(5) = NulSID(namaSplit(i))
                Else
                    .txt(6) = .txt(6) & NulSID(namaSplit(i)) & " "
                End If
            Next
            .isEdit False
            .show vbModal
        End With
    End If
    If NulSID(gBantuan) <> "" Then
        If NulSID(gBantuan) = "" Or NulSID(gBantuan) = "BOOKING" Then
            Text(4) = ""
        Else
            Text(4) = NulSID(gBantuan)
        End If
        
    End If
    Call PencarianData
    If Me.Visible Then
        Grid.SetFocus
    End If
End Sub

Private Sub tglLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text(4).SetFocus
    End If
End Sub



'=====================
''20171027 - Param
'=====================
''    param = cari                                'NOPOL & NOKA
''
''    If cari2 <> "" Then                         'KTP
''        If cari = "" Then
''            param = "(" & cari2
''        Else
''            param = param & " AND (" & cari2
''        End If
''    End If
''
''    If cari3 <> "" Then                         'NAMA
''        If cari2 = "" Then
''            If cari = "" Then
''                param = "(" & cari3
''            Else
''                param = param & " AND (" & cari3
''            End If
''        Else
''            If cari = "" Then
''                param = param & " AND " & cari3
''            Else
''                param = param & " AND " & cari3
''            End If
''        End If
''    End If
''
''    If cari4 <> "" Then
''        If cari = "" Then
''            If cari2 = "" Then
''                If cari3 = "" Then
''                    param = "(" & cari4
''                Else
''                    param = param & " AND " & cari4
''                End If
''            Else
''                param = param & " AND " & cari4
''            End If
''        Else
''            If cari2 = "" Then
''                If cari3 = "" Then
''                    param = param & " AND (" & cari4
''                Else
''                    param = param & " AND " & cari4
''                End If
''            Else
''                If cari3 = "" Then
''                    param = param & " AND " & cari4
''                Else
''                    param = param & " AND " & cari4
''                End If
''            End If
''        End If
''    End If
''
''    If cari5 <> "" Then
''        If cari = "" Then
''            If cari2 = "" Then
''                If cari3 = "" Then
''                    If cari4 = "" Then
''                        param = "(" & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                Else
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                End If
''            Else
''                If cari3 = "" Then
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                Else
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                End If
''            End If
''        Else
''            If cari2 = "" Then
''                If cari3 = "" Then
''                    If cari4 = "" Then
''                        param = param & " AND (" & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                Else
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                End If
''            Else
''                If cari3 = "" Then
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                Else
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                End If
''            End If
''        End If
''    End If


Private Function simpanImageKTP(idcust As String) As Boolean
    Dim fotoStream      As New ADODB.Stream
    Dim readStream      As ADODB.Stream
    Dim rsUploadKTP     As New ADODB.Recordset
   On Error GoTo simpanImageKTP_Error
   
    sSQL = "SELECT TOP 1 * FROM UploadKTP WHERE IDCUST = '" & NulSID(idcust) & "' ORDER BY TGLENTRY DESC "
    If rsUploadKTP.State Then rsUploadKTP.Close
    If fotoStream.State Then fotoStream.Close
    rsUploadKTP.Open sSQL, SQLConSID, adOpenKeyset, adLockOptimistic
     
        With fotoStream
            .Type = adTypeBinary
            .Open
            If CommonDialog1.FileName <> "" Then
            .LoadFromFile CommonDialog1.FileName
             
            End If
            If rsUploadKTP.EOF Then
                rsUploadKTP.AddNew
                rsUploadKTP!idcust = NulSID(idcust)
                rsUploadKTP!delIndex = NulSID(gKode & Left(McDelIndex(7), 8))
                rsUploadKTP!App = NulSID("SV")
                rsUploadKTP!dealer = NulSID(Right(gDb, 2))
                rsUploadKTP!userentry = NulSID(gSA)
                rsUploadKTP!PIC = .Read
                rsUploadKTP!tglentry = isnow
                rsUploadKTP.Update
            Else
                rsUploadKTP!PIC = .Read
                rsUploadKTP!UpdateKe = CDbl(rsUploadKTP!UpdateKe) + 1
                rsUploadKTP!UserUpdate = NulSID(gSA)
                rsUploadKTP!TglUpdate = isnow
                rsUploadKTP.Update
            End If
        End With
        Set fotoStream = Nothing
        simpanImageKTP = True
   On Error GoTo 0
   Exit Function
simpanImageKTP_Error:
    simpanImageKTP = False
    MsgBox "Error " & Erl & " (" & err.Description & ") in procedure simpanImageKTP of Form frmSIDMasterPelanggan"
End Function

 
