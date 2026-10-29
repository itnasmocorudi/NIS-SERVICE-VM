VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmInfoValidasiPartTMSRetur 
   Caption         =   ".: Validasi Part TMS Retur :."
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmInfoValidasiPartTMSRetur.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8535
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   13335
      Begin ARProgBarCtrl.ARProgressBar ProgBar 
         Height          =   615
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   1085
         Value           =   0
         ForeColor       =   52479
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionForeColor=   128
         IniColor        =   52479
      End
      Begin TrueOleDBGrid70.TDBGrid GridSupplier 
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   10186
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Cek"
         Columns(0).DataField=   "Cek"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "ReqCode"
         Columns(1).DataField=   "ReqCode"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Kode Item"
         Columns(2).DataField=   "KodeItem"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nama Barang"
         Columns(3).DataField=   "NamaBarang"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Sat"
         Columns(4).DataField=   "KodeSatuan"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Harga Jual"
         Columns(5).DataField=   "HargaJual"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Qty Retur"
         Columns(6).DataField=   "QtyOrder"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Qty Approve"
         Columns(7).DataField=   "QtySupply"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=979"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2408"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2302"
         Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2566"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2461"
         Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=4154"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=4048"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=1508"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1402"
         Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(25)=   "Column(5).Width=3281"
         Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=8196"
         Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(30)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=8196"
         Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(35)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
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
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFF80&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&H4080FF&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
         _StyleDefs(89)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(90)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(91)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(92)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2400
         MaxLength       =   60
         TabIndex        =   17
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2400
         MaxLength       =   60
         TabIndex        =   15
         Top             =   2400
         Width           =   3615
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import Excel"
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
         Left            =   8640
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2400
         MaxLength       =   60
         TabIndex        =   9
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2400
         MaxLength       =   60
         TabIndex        =   6
         Top             =   1680
         Width           =   3615
      End
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   495
         Left            =   4920
         TabIndex        =   2
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":84E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdEdit 
         Height          =   495
         Left            =   7320
         TabIndex        =   3
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":8500
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
         Left            =   6120
         TabIndex        =   4
         Top             =   7440
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":851C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdViewKPI 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":8538
         PICN            =   "frmInfoValidasiPartTMSRetur.frx":8554
         PICH            =   "frmInfoValidasiPartTMSRetur.frx":C2B5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdRejectPart 
         Height          =   495
         Left            =   4320
         TabIndex        =   11
         Top             =   960
         Width           =   1750
         _ExtentX        =   3096
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":101EC
         PICN            =   "frmInfoValidasiPartTMSRetur.frx":10208
         PICH            =   "frmInfoValidasiPartTMSRetur.frx":11E03
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdApprove 
         Height          =   495
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":13B39
         PICN            =   "frmInfoValidasiPartTMSRetur.frx":13B55
         PICH            =   "frmInfoValidasiPartTMSRetur.frx":157A5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   12000
         TabIndex        =   19
         Top             =   7440
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":174E1
         PICN            =   "frmInfoValidasiPartTMSRetur.frx":174FD
         PICH            =   "frmInfoValidasiPartTMSRetur.frx":1DE30
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdBatal 
         Height          =   495
         Left            =   10800
         TabIndex        =   20
         Top             =   7440
         Visible         =   0   'False
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":24AE7
         PICN            =   "frmInfoValidasiPartTMSRetur.frx":24B03
         PICH            =   "frmInfoValidasiPartTMSRetur.frx":2B0FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   9600
         TabIndex        =   21
         Top             =   7440
         Visible         =   0   'False
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
         MICON           =   "frmInfoValidasiPartTMSRetur.frx":319AB
         PICN            =   "frmInfoValidasiPartTMSRetur.frx":319C7
         PICH            =   "frmInfoValidasiPartTMSRetur.frx":381CA
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
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "KETERSEDIAAN INSENTIF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2805
         Width           =   2040
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "PORSI INSENTIF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   16
         Top             =   2445
         Width           =   1320
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "OPEX LEMBUR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1065
         TabIndex        =   10
         Top             =   2085
         Width           =   1140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   12600
         X2              =   120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "DPP JASA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1425
         TabIndex        =   7
         Top             =   1725
         Width           =   810
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmInfoValidasiPartTMSRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmInfoValidasiPartTMSRetur
' DateTime  : 19/09/2018
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------


Dim rsGrid      As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsTemp      As New ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset
Dim cols        As TrueOleDBGrid70.Columns
Dim Urut        As String
Dim isTambah    As Boolean
Dim isBegin     As Boolean
Dim isLoad      As Boolean

Private dppJasa, opexLembur, ketersediaan, totalBobot, iProgress As Double
 

Private Sub cmdApprove_Click()
1        On Error GoTo cmdApprove_Click_Error
         Dim delIndex, vPersediaanD, vPersediaanK, kodeItemAsli As String
         Dim hargaPokok, hargaBeli As Double
         
         
         Dim isPart As Boolean
2         isPart = False
3         isBegin = False
          
4         If Not rsGrid.EOF Then
5             rsGrid.MoveFirst
6             Call objHT.OpenTransactionAll
7             isBegin = True
8             delIndex = "IN" & Left(McDelIndex(7), 8)
9             Do While Not rsGrid.EOF
10                If rsGrid!cek = True Then
11                   kodeItemAsli = Mid(NulSID(rsGrid!KodeItem), 6, Len(rsGrid!KodeItem)) 'Replace(rsGrid!KodeItem, "TMS1-", "")
                     'kodeItemAsli = Replace(rsGrid!KodeItem, "TMS2-", "")
12                   kodeItemTms = rsGrid!KodeTMS & "-" & rsGrid!KodeItem
                     
13                   hargaBeli = CDbl(objHT.FetchDataACC_String("select HargaBeli  FROM Inmst where KodeItem = '" & rsGrid!KodeItem & "'"))
14                   hargaPokok = CDbl(objHT.FetchDataACC_String("select HPokokSkr  FROM Inmst where KodeItem = '" & rsGrid!KodeItem & "'"))
15                   sSQL = "INSERT INTO Intrn (NoBukti,KodeItem,TglTrn,TypeTrn,Qty,"
16                   sSQL = sSQL + "Qoh,HargaPokok,HargaBeli,HargaJual,Kodejurnal,"
17                   sSQL = sSQL + "DelIndex,SbrBukti,QtyEntry,Konversi,SatuanNs,"
18                   sSQL = sSQL + "TypeSatuan,KodeUser,TglEntry,nobaris)"
19                   sSQL = sSQL + "VALUES('TMSOUT','" & rsGrid!KodeItem & "',"
                     'sSQL = sSQL + "'" & FormatTgl(isnow) & "','S',"
20                   sSQL = sSQL + "GETDATE(),'S',"
21                   sSQL = sSQL + "" & k2t(-Abs(rsGrid!QtySupply)) & "," & k2t(rsGrid!QtySupply) & ","
22                   sSQL = sSQL + "" & k2t(hargaPokok) & "," & k2t(hargaBeli) & ","
23                   sSQL = sSQL + "" & k2t(rsGrid!hargaJual) & ",'',"
24                   sSQL = sSQL + "'" & delIndex & "','" & rsGrid!ReturCode & "',"
25                   sSQL = sSQL + "" & k2t(Abs(rsGrid!QtySupply)) & ",null,"
26                   sSQL = sSQL + "null,null,"
27                   sSQL = sSQL + "'" & gSA & "',GetDate(),(select max(nobaris) from intrn)+1)"
                     
28                   Call objHT.InsertData_ACC(sSQL)
                     
                     
29                   Set rsTemp = objHT.FetchDataACC("SELECT * FROM Inhdr WHERE NoBukti = 'TMSOUT'")
30                   If rsTemp.EOF Then
31                    sSQL = "INSERT INTO [Inhdr]"
32                    sSQL = sSQL + " ([NoBukti],[TypeTrn],[TglTrn],[KodeLgn],[Keterangan],[KodeCCenter],[GdgTarget],"
33                    sSQL = sSQL + " [KodeSumber],[DelIndex],[KodeJurnal],[TglEntry],[Kodeuser],[TrfFlag],[TutupBukuFlag]) "
34                    sSQL = sSQL + " VALUES ('TMSOUT','O','" & FormatTgl(isnow) & "','','MutasiTMS','','','IN','" & delIndex & "',"
35                    sSQL = sSQL + " '',getdate(),'" & gSA & "','0','0')"
                      
36                    Call objHT.InsertData_ACC(sSQL)
37                   End If
38                   sSQL = "INSERT INTO Intrn (NoBukti,KodeItem,TglTrn,TypeTrn,Qty,"
39                   sSQL = sSQL + "Qoh,HargaPokok,HargaBeli,HargaJual,Kodejurnal,"
40                   sSQL = sSQL + "DelIndex,SbrBukti,QtyEntry,Konversi,SatuanNs,"
41                   sSQL = sSQL + "TypeSatuan,KodeUser,TglEntry,nobaris)"
42                   sSQL = sSQL + "VALUES('TMSIN','" & kodeItemAsli & "',"
                     'sSQL = sSQL + "'" & FormatTgl(isnow) & "','G',"
43                   sSQL = sSQL + "GETDATE(),'G',"
44                   sSQL = sSQL + "" & k2t(Abs(rsGrid!QtySupply)) & ",0,"
45                   sSQL = sSQL + "" & k2t(hargaPokok) & "," & k2t(hargaBeli) & ","
46                   sSQL = sSQL + "" & k2t(rsGrid!hargaJual) & ",'',"
47                   sSQL = sSQL + "'" & delIndex & "','" & rsGrid!ReturCode & "',"
48                   sSQL = sSQL + "" & k2t(Abs(rsGrid!QtySupply)) & ",null,"
49                   sSQL = sSQL + "null,null,"
50                   sSQL = sSQL + "'" & gSA & "',GetDate(),(select max(nobaris) from intrn)+1)"
                     
51                   Call objHT.InsertData_ACC(sSQL)
                     
52                   Set rsTemp = objHT.FetchDataACC("SELECT * FROM Inhdr WHERE NoBukti = 'TMSIN'")
53                   If rsTemp.EOF Then
54                    sSQL = "INSERT INTO [Inhdr]"
55                    sSQL = sSQL + " ([NoBukti],[TypeTrn],[TglTrn],[KodeLgn],[Keterangan],[KodeCCenter],[GdgTarget],"
56                    sSQL = sSQL + " [KodeSumber],[DelIndex],[KodeJurnal],[TglEntry],[Kodeuser],[TrfFlag],[TutupBukuFlag]) "
57                    sSQL = sSQL + " VALUES ('TMSIN','G','" & FormatTgl(isnow) & "','','MutasiTMS','','','IN','" & delIndex & "',"
58                    sSQL = sSQL + " '',getdate(),'" & gSA & "','0','0')"
                      
59                    Call objHT.InsertData_ACC(sSQL)
60                   End If
61                   Set rsTemp = objHT.FetchDataACC("SELECT * FROM inmst WHERE KodeItem = '" & rsGrid!KodeItem & "'")
62                   If Not rsTemp.EOF Then
63                    sSQL = "UPDATE inmst SET QtyAkhir=QtyAkhir - " & k2t(Abs(rsGrid!QtySupply)) & " WHERE KodeItem = '" & rsGrid!KodeItem & "';"
64                    sSQL = sSQL & "UPDATE inmst SET QtyAkhir=QtyAkhir + " & k2t(Abs(rsGrid!QtySupply)) & " WHERE KodeItem = '" & kodeItemAsli & "';"
65                   Else
66                       Set rsTemp = objHT.FetchDataACC("SELECT * FROM inmst WHERE KodeItem = '" & kodeItemAsli & "'")
67                       If Not rsTemp.EOF Then
68                       sSQL = "INSERT INTO InMst (kodeitem,NamaBarang,KodeKategori,KodeGdg,NoItem,"
69                       sSQL = sSQL + "CostSystem,NoRack,KodeSatuan,KodeSupplier,HargaBeli,"
70                       sSQL = sSQL + "HPokokAwal,KodeWarna,KodeMerek,KodeTahun,"
71                       sSQL = sSQL + "NoPolisi,NoMesin,HPokokSkr,HPokokAkhir,HargaJualA,"
72                       sSQL = sSQL + "HargaJualB,QtyMin,QtyMax,LeadTime,QtySlowMove,"
73                       sSQL = sSQL + "QtyFastMove,QtyStockOpn,Memo,KodeUser,AdaSatLain,"
74                       sSQL = sSQL + "TglEntry,QtyAkhir)"
75                       sSQL = sSQL + "VALUES ('" & rsGrid!KodeItem & "','" & rsTemp!NamaBarang & "',"
76                       sSQL = sSQL + "'" & rsTemp!kodeKATEGori & "','07',"
77                       sSQL = sSQL + "'" & rsGrid!KodeItem & "','" & rsTemp!CostSystem & "',"
78                       sSQL = sSQL + "'" & rsTemp!NoRack & "','" & rsTemp!KODESATUAN & "',"
79                       sSQL = sSQL + "'" & rsTemp!KodeSupplier & "'," & k2t(rsTemp!hargaBeli) & ","
80                       sSQL = sSQL + "" & k2t(rsTemp!HPokokAwal) & ",'" & rsTemp!KodeWarna & "',"
81                       sSQL = sSQL + "'" & rsTemp!KodeMerek & "','" & rsTemp!kodetahun & "',"
82                       sSQL = sSQL + "'" & rsTemp!NoPolisi & "','" & rsTemp!noMesin & "',"
83                       sSQL = sSQL + "" & k2t(rsTemp!HPokokSkr) & "," & k2t(rsTemp!HPokokAkhir) & ","
84                       sSQL = sSQL + "" & k2t(rsTemp!HargaJualA) & "," & k2t(rsTemp!HargaJualB) & ","
85                       sSQL = sSQL + "" & rsTemp!QtyMin & "," & rsTemp!QtyMax & ","
86                       sSQL = sSQL + "'" & Format(rsTemp!LeadTime, "YYYY-MM-DD") & "'," & rsTemp!QtySlowMove & ","
87                       sSQL = sSQL + "" & rsTemp!QtyFastMove & "," & rsTemp!QtyStockOpn & ","
88                       sSQL = sSQL + "'" & rsTemp!Memo & "','" & gSA & "',"
89                       sSQL = sSQL + "" & IIf(rsTemp!AdaSatLain, 1, 0) & ","
90                       sSQL = sSQL + "GetDate()," & k2t(Abs(rsGrid!QtySupply)) & ");UPDATE inmst SET QtyAkhir=QtyAkhir + " & k2t(Abs(rsGrid!QtySupply)) & " WHERE KodeItem = '" & rsGrid!KodeItem & "';"
91                       Else
                         
92                            Call objHT.CloseTransactionAll_Err
93                            MsgBox "Kode item " & rsGrid!KodeItem & " tidak terdaftar di INMST", vbInformation
94                            Exit Sub
95                       End If
96                   End If
                     
97                   Call objHT.InsertData_ACC(sSQL)
                     
98                   If Left(UCase(objO.FetchData_String_Acc("select B.Jenis from Inmst A WITH (NOLOCK) Inner Join Profilacc..SCC B WITH (NOLOCK) On A.SCC_Code = B.Scc_Code Where A.KodeItem = '" & Trim(rsGrid("KodeItem")) & "'")), 3) = "PAR" Then
99                    isPart = True
100                   vPersediaanD = "11322100"
101                   vPersediaanK = "11325100"
102                  Else
103                   isPart = False
104                   vPersediaanD = "11322200"
105                   vPersediaanK = "11325200"
106                  End If
                     
107                  sSQL = "INSERT INTO GlTrn(KodeGl,TglTrn,TypeTrn,NoBukti,Nobatch,"
108                   sSQL = sSQL + "KodeSumber,KodeRekl,keterangan,JlhDebit,JlhKredit,"
109                   sSQL = sSQL + "KodeUser,DelIndex,LastUpdated)"
110                   sSQL = sSQL + "VALUES('" & vPersediaanD & "','" & FormatTgl(isnow) & "',"
111                   sSQL = sSQL + "'O','" & rsGrid!ReturCode & "',"
112                   sSQL = sSQL + "'0','SV',"
113                   sSQL = sSQL + "'','TMS-" & rsGrid!KodeItem & "',"
114                   sSQL = sSQL + "" & k2t(rsGrid!hargaJual) * k2t(rsGrid!QtySupply) & ",0,"
115                   sSQL = sSQL + "'" & gSA & "','" & delIndex & "',"
116                   sSQL = sSQL + " GetDate())"
                     
117                  Call objHT.InsertData_ACC(sSQL)
                     
                     
118                  sSQL = "INSERT INTO GlTrn(KodeGl,TglTrn,TypeTrn,NoBukti,Nobatch,"
119                   sSQL = sSQL + "KodeSumber,KodeRekl,keterangan,JlhDebit,JlhKredit,"
120                   sSQL = sSQL + "KodeUser,DelIndex,LastUpdated)"
121                   sSQL = sSQL + "VALUES('" & vPersediaanK & "','" & FormatTgl(isnow) & "',"
122                   sSQL = sSQL + "'O','" & rsGrid!ReturCode & "',"
123                   sSQL = sSQL + "'0','SV',"
124                   sSQL = sSQL + "'','TMS-" & rsGrid!KodeItem & "',"
125                   sSQL = sSQL + "0," & k2t(rsGrid!hargaJual) * k2t(rsGrid!QtySupply) & ","
126                   sSQL = sSQL + "'" & gSA & "','" & delIndex & "',"
127                   sSQL = sSQL + " GetDate())"
                     
128                  Call objHT.InsertData_ACC(sSQL)
                     
129                  sSQL = "Update ReturPart set QtyApproved='" & k2t(rsGrid!QtySupply) & "', [Status]=2 where kodeitem = '" & rsGrid!KodeItem & "' and ReturCode='" & rsGrid!ReturCode & "' and [Status] = 1"
130                  Call objHT.InsertData_Bengkel(sSQL)
131               End If
132               rsGrid.MoveNext
133           Loop
134           objHT.CloseTransactionAll
              
135       End If
136       Call refreshGrid
137      On Error GoTo 0
138      Exit Sub

cmdApprove_Click_Error:
139       If isBegin Then
140           Call objHT.CloseTransactionAll_Err
141       End If
142       MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdApprove_Click of Form frmInfoValidasiPartTMSRetur"
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdRejectPart_Click()

1        On Error GoTo cmdRejectPart_Click_Error

          Dim delIndex As String
         
2         isBegin = False
          
3         If Not rsGrid.EOF Then
4             rsGrid.MoveFirst
5             Call objHT.OpenTransactionAll
6             isBegin = True
7             delIndex = "IN" & Left(McDelIndex(7), 8)
8             Do While Not rsGrid.EOF
9                 If rsGrid!cek = True Then
                     
10                   sSQL = "Update ReturPart set QtyApproved='0', [Status]=3 where  kodeitem = '" & rsGrid!KodeItem & "' and  ReturCode='" & rsGrid!ReturCode & "' and [Status] = 1"
11                   Call objHT.InsertData_Bengkel(sSQL)
                     
12                End If
13                rsGrid.MoveNext
14            Loop
15            objHT.CloseTransactionAll
16            MsgBox "Data berhasil disimpan", vbInformation
17        End If
18        Call refreshGrid

19       On Error GoTo 0
20       Exit Sub

cmdRejectPart_Click_Error:
          
21        If isBegin Then
22            Call objHT.CloseTransactionAll_Err
23        End If
24        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdRejectPart_Click of Form frmInfoValidasiPartTMSRetur"
End Sub

Private Sub cmdViewKPI_Click()
    Call refreshGrid
End Sub

Private Sub Form_Load()
    Call refreshGrid
End Sub


Private Sub GridSupplier_FilterChange()
    On Error GoTo errHandler
    Set cols = GridSupplier.Columns
    Dim C As Integer
    
    C = GridSupplier.col
    GridSupplier.HoldFields
    rsGrid.filter = getFilter()
    GridSupplier.col = C
    GridSupplier.EditActive = True
    Exit Sub
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub GridSupplier_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo GridSupplier_HeadClick_Error

    Urut = IIf(Urut = "DESC", "ASC", "DESC")
    rsGrid.Sort = GridSupplier.Columns(ColIndex).DataField & " " & Urut
    GridSupplier.SetFocus
    Sendkeys "{Right " & Val(ColIndex) & "}"
 
   Exit Sub

GridSupplier_HeadClick_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure GridSupplier_HeadClick of Form frmInfoValidasiPartTMSRetur")
 End Sub

Private Sub gridSupplier_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF And isLoad Then
'        Text(0).Text = NulSID(rsGrid!nik)
'        Text(1).Text = NulSID(rsGrid!Nama)
'        Text(2).Text = NulSID(rsGrid!KdKPI)
'        Text(3).Text = NulSID(rsGrid!namaKPI)
'        Text(4).Text = NulSID(rsGrid!target)
'        If NulSID(rsGrid!AllowAdj) = 0 Then
'            Text(4).Enabled = False
'        Else
'            Text(4).Enabled = True
'        End If
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
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
    'KeyCode = UCase(KeyCode)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 9
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLGAP"
                frmCariAcess.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) ORDER BY RIGHT(KDSUPP,10) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KDSUPP"), 4))
        noUrut = noUrut + 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    End If
End Sub

Private Sub isiModel()
    Set rsTemp = objO.FetchData("SELECT DISTINCT a.MODEL, a.mdl, a.[desc], ISNULL(id_program, '') as PRO FROM [" & dbReplikasiHO & "]..model a WITH (NOLOCK) LEFT JOIN ho_PROmdl b WITH (NOLOCK) ON a.MDL = b.MDL AND B.ID_PROGRAM = '" & Trim(Text(0)) & "' ORDER BY a.MODEL, a.mdl")
    If Not rsTemp.EOF Then
        Call createTbl
        Do While Not rsTemp.EOF
            rsGridMdl.AddNew
            rsGridMdl!grpMDL = Trim(rsTemp("MODEL"))
            rsGridMdl!kodeMDL = Trim(rsTemp("MDL"))
            rsGridMdl!NAMAMDL = Trim(rsTemp("DESC"))
            rsGridMdl!ikut = IIf(rsTemp("PRO") = "", False, True)
            rsGridMdl.Update
            rsTemp.MoveNext
        Loop

        rsGridMdl.MoveFirst
        GridModel.Columns(0).Width = 1200
        GridModel.Columns(1).Width = 2000
        GridModel.Columns(2).Width = 4000
        GridModel.Columns(3).Alignment = dbgCenter
        GridModel.Columns(3).Width = 800
    End If
End Sub

Private Sub createTbl()
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    With rsGridMdl
        .Fields.Append "GRPMDL", adVarChar, 15
        .Fields.Append "KODEMDL", adVarChar, 20
        .Fields.Append "NAMAMDL", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridModel.DataSource = rsGridMdl
    GridModel.AlternatingRowStyle = True
End Sub

Private Sub GridModel_FilterChange()
    On Error GoTo errHandler
    Set cols = GridModel.Columns
    Dim C As Integer
    
    C = GridModel.col
    GridModel.HoldFields
    rsGridMdl.filter = getFilter()
    GridModel.col = C
    GridModel.EditActive = True
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

 

Private Sub refreshGrid()
          ''GridSupplier.ClearFields
1         Call createRS
          Dim porsiInsentif As Double
          ''sSQL = "SELECT KDSUPP,SUPPLIER,ALAMAT1,ALAMAT2,KOTA,TELPON,CP,KETERANGAN,KDDEALER,PKP,NONJURNAL FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY KDDEALER "
         ' If NulSID(cbJenis.Text) = NulSID("Part") Then
2             sSQL = "select ReturCode,KodeItem,NamaBarang,KodeSatuan,HargaJual,QtyRetur,QtyApproved,KodeTMS from ReturPart where [Status] = 1"
          'Else
          '    sSQL = "Select KdMP,NIK,Nama,NoInv,No_WO,KdItem,NamaItem,Qty,Diskon,AllowDiskon,AlokasiIns from  INSENTIF.dbo.GRInsProdukTersedia where KategoriItem = 'Sublet' and Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_DEALER = '" & NulSID(gKdCbg) & "'"
          'End If
          
          
3         Set rsTemp = objO.FetchData(sSQL)
4         If Not rsTemp.EOF Then
5             rsTemp.MoveFirst
6             iProgress = 1
7             progBar.Visible = True
8             DoEvents
9             progBar.Min = 1
10            progBar.Max = rsTemp.RecordCount + 1
              
11            Do While Not rsTemp.EOF
12                With rsGrid
13                    .AddNew
14                    !cek = True
15                    !ReturCode = NulSID(rsTemp!ReturCode)
16                    !KodeItem = NulSID(rsTemp!KodeItem)
17                    !NamaBarang = NulSID(rsTemp!NamaBarang)
18                    !KODESATUAN = NulSID(rsTemp!KODESATUAN)
19                    !hargaJual = k2t(rsTemp!hargaJual)
20                    !QtyOrder = k2t(rsTemp!QtyRetur)
21                    !QtySupply = k2t(rsTemp!QtyApproved)
22                    !KodeTMS = NulSID(rsTemp!KodeTMS)
23                    .Update
24                End With
                  
25                progBar.value = iProgress
26                iProgress = iProgress + 1
27                DoEvents
                  
28                rsTemp.MoveNext
29            Loop
30            isLoad = True
31        End If
32        progBar.Visible = False
33        Set GridSupplier.DataSource = rsGrid
End Sub

Private Sub createRS()
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = Nothing
    With rsGrid
        .Fields.Append "Cek", adBoolean
        .Fields.Append "ReturCode", adVarChar, 100
        .Fields.Append "KodeItem", adVarChar, 100
        .Fields.Append "NamaBarang", adVarChar, 200
        .Fields.Append "KodeSatuan", adVarChar, 100
        .Fields.Append "HargaJual", adDouble
        .Fields.Append "QtyOrder", adDouble
        .Fields.Append "QtySupply", adDouble
        .Fields.Append "KodeTMS", adVarChar, 100
        .Open
    End With
End Sub
