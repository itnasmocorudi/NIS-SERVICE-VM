VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmInfoKetersediaanInsentifCabang 
   Caption         =   ".: Perhitungan Ketersediaan Insentif Cabang :."
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmInfoKetersediaanInsentifCabang.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   6
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
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   13005
         _ExtentX        =   22939
         _ExtentY        =   873
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
         TabIndex        =   19
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
         TabIndex        =   17
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
         Left            =   4920
         TabIndex        =   15
         Top             =   960
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
         TabIndex        =   11
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
         TabIndex        =   7
         Top             =   1680
         Width           =   3615
      End
      Begin TrueOleDBGrid70.TDBGrid GridSupplier 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   7435
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Posisi"
         Columns(0).DataField=   "tipe"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Bobot"
         Columns(1).DataField=   "BobotIns"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Jml Manpower"
         Columns(2).DataField=   "jmlMp"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Jml Bobot"
         Columns(3).DataField=   "jmlBobot"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Persentase Bobot"
         Columns(4).DataField=   "jmlPersen"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Porsi Per Posisi"
         Columns(5).DataField=   "porsiPosisi"
         Columns(5).NumberFormat=   "#,#0"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Porsi Manpower"
         Columns(6).DataField=   "porsiMP"
         Columns(6).NumberFormat=   "#,#0"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
         Splits(0)._ColumnProps(24)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(26)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(29)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
         Splits.Count    =   1
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&H4080FF&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
         _StyleDefs(85)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(86)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(87)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(88)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   495
         Left            =   4920
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":A2CD
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
         TabIndex        =   4
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":A2E9
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
         TabIndex        =   5
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":A305
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBDate6Ctl.TDBDate Periode 
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   300
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calendar        =   "frmInfoKetersediaanInsentifCabang.frx":A321
         Caption         =   "frmInfoKetersediaanInsentifCabang.frx":A44D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoKetersediaanInsentifCabang.frx":A4B9
         Keys            =   "frmInfoKetersediaanInsentifCabang.frx":A4D7
         Spin            =   "frmInfoKetersediaanInsentifCabang.frx":A535
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm-yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/yyyy"
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
         Text            =   "11/2024"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   45611
         CenturyMode     =   0
      End
      Begin ucXPButton.XPButton cmdViewKPI 
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   240
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":A55D
         PICN            =   "frmInfoKetersediaanInsentifCabang.frx":A579
         PICH            =   "frmInfoKetersediaanInsentifCabang.frx":E2DA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdGenerateKPIMP 
         Height          =   495
         Left            =   2520
         TabIndex        =   13
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":12211
         PICN            =   "frmInfoKetersediaanInsentifCabang.frx":1222D
         PICH            =   "frmInfoKetersediaanInsentifCabang.frx":16348
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKPIMP 
         Height          =   495
         Left            =   120
         TabIndex        =   14
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":1A98E
         PICN            =   "frmInfoKetersediaanInsentifCabang.frx":1A9AA
         PICH            =   "frmInfoKetersediaanInsentifCabang.frx":1EB2D
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
         TabIndex        =   21
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":2309F
         PICN            =   "frmInfoKetersediaanInsentifCabang.frx":230BB
         PICH            =   "frmInfoKetersediaanInsentifCabang.frx":299EE
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
         TabIndex        =   22
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":306A5
         PICN            =   "frmInfoKetersediaanInsentifCabang.frx":306C1
         PICH            =   "frmInfoKetersediaanInsentifCabang.frx":36CBA
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
         TabIndex        =   23
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
         MICON           =   "frmInfoKetersediaanInsentifCabang.frx":3D569
         PICN            =   "frmInfoKetersediaanInsentifCabang.frx":3D585
         PICH            =   "frmInfoKetersediaanInsentifCabang.frx":43D88
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
         TabIndex        =   20
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
         TabIndex        =   18
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
         TabIndex        =   12
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
         TabIndex        =   8
         Top             =   1725
         Width           =   810
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   2
         Top             =   360
         Width           =   1305
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
Attribute VB_Name = "frmInfoKetersediaanInsentifCabang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmInfoKetersediaanInsentifCabang
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
Dim iProgress   As Integer

Private dppJasa, opexLembur, ketersediaan, totalBobot As Double

Private Sub ber_sih()
    For i = 0 To 3
        Text(i) = ""
    Next
    isLoad = False
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
'    For i = 1 To 4
'        Text(i).Locked = Not Flg
'        If Flg Then
'            Text(i).BackColor = vbWhite
'        Else
'            Text(i).BackColor = &HC0E0FF
'        End If
'    Next
    cmdSimpan.Enabled = Flg
    cmdBatal.Enabled = Flg
    
   ' cmdTambah.Enabled = Not Flg
    cmdEdit.Enabled = Not Flg
    cmdHapus.Enabled = Not Flg
End Sub

Private Sub chAllDealer_Click()
     
    If chAllDealer = Unchecked Then
        frmeDealer.Visible = True
        If rsDealer.State = adStateClosed Then
            Call isiDealer
        End If
        frmeDealer.Visible = True
        GridDealer.SetFocus
    ElseIf chAllDealer = Checked Then
        frmeDealer.Visible = False
    End If
End Sub

Private Sub chModel_Click()
    If chModel.value = Checked Then
        frmeModel.Visible = False
    ElseIf chModel.value = Unchecked Then
        frmeModel.Visible = True
        Call isiModel
    End If
End Sub

Private Sub cmdBatal_Click()
    Call buka_Tutup(False)
    Call refreshGrid
End Sub

Private Sub cmdEdit_Click()
1         isTambah = False
2         Call buka_Tutup(True)
3         If NulSID(GridSupplier.Columns(8).value) = "NR 00" Then
4             For i = 1 To 8
5                 Text(i).Locked = True
6                 Text(i).BackColor = &HC0E0FF
7             Next
8             Text(10).Locked = True
9             Text(10).BackColor = &HC0E0FF
10        ElseIf NulSID(GridSupplier.Columns(8).value) <> NulSID(gKdCbg) Then
11            For i = 1 To 8
12                Text(i).Locked = True
13                Text(i).BackColor = &HC0E0FF
14            Next
15            Text(10).Locked = True
16            Text(10).BackColor = &HC0E0FF
17        End If
End Sub

Private Sub cmdGenerateKPIMP_Click()
1          If Not rsGrid.EOF Then
2             isLoad = False
3             rsGrid.MoveFirst
4             iProgress = 1
5             ProgBar.Visible = True
6             DoEvents
7             ProgBar.Min = 1
8             ProgBar.Max = rsGrid.RecordCount + 1
9             Do While Not rsGrid.EOF
                  
10                If totalBobot > 0 Then
11                  rsGrid!jmlPersen = IIf(totalBobot > 0, Round((rsGrid!jmlBobot * 100) / totalBobot, 2), 0)
12                Else
13                  rsGrid!jmlPersen = 0
14                End If
15                If rsGrid!jmlPersen > 0 Then
16                  rsGrid!porsiPosisi = Round((ketersediaan * rsGrid!jmlPersen) / 100)
17                Else
18                  rsGrid!porsiPosisi = 0
19                End If
                  
20                If rsGrid!jmlMp > 0 Then
21                  rsGrid!PorsiMP = IIf(rsGrid!jmlMp > 0, Round((rsGrid!porsiPosisi / rsGrid!jmlMp), 2), 0)
22                Else
23                  rsGrid!PorsiMP = 0
24                End If
                  

25                ProgBar.value = iProgress
26                iProgress = iProgress + 1
27                DoEvents
28                rsGrid.MoveNext
29            Loop
30            ProgBar.Visible = False
31            isLoad = True
32        End If
End Sub

Private Sub cmdHapus_Click()
1        On Error GoTo cmdHapus_Click_Error
2         isBegin = False
3         If NulSID(GridSupplier.Columns(8).value) = "NR 00" Then
4             MsgBox "Master supplier HO tidak dapat di hapus!", vbInformation
5             Exit Sub
6         ElseIf NulSID(GridSupplier.Columns(8).value) <> NulSID(gKdCbg) Then
7             MsgBox "Master supplier Dealer lain tidak dapat di hapus!", vbInformation
8             Exit Sub
9         End If
          
10        Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE ispakai = '1' and KDSUPP = '" & NulSID(Text(0)) & "'")
11        If Not rsTemp.EOF Then
12            MsgBox "Sudah Ada Transaksi Work Order untuk sublet ini" & vbCrLf & "Kode Supplier " & NulSID(Text(0)) & " tidak dapat dihapus", vbInformation
13            Exit Sub
14        End If
15        If MsgBox("Apakah yakin menghapus data kode supplier " & NulSID(Text(0)) & " ?", vbYesNo, "Konfirmasi Hapus Data") = vbYes Then
16            Set rsTemp = objO.FetchDataHO("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..mstsublet WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND KDSUPP = '" & NulSID(Text(0)) & "'")
17            If Not rsTemp.EOF Then
18                MsgBox "Masih ada sublet untuk supplier ini." & vbCrLf & "Kode Sublet " & NulSID(rsTemp!KDSUBLET) & " masih aktif tidak dapat menghapus supplier", vbInformation
19                Exit Sub
20            End If
21            Call objHT.OpenTransactionBengkel
22            isBegin = True
23            erTranAll = True
24            sSQL = "UPDATE [" & dbReplikasiHO & "]..MSTSUPPLIER SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = GETDATE() ,STATUS = 0 WHERE KDSUPP = '" & NulSID(Text(0)) & "'"
25            Call objHT.InsertData_Bengkel(sSQL)
26            sSQL = "UPDATE T_SUPPLIER SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = GETDATE() ,STATUS = 0 WHERE KDSUPP = '" & NulSID(Text(0)) & "'"
27            Call objHT.InsertData_Bengkel(sSQL)
              
28            If erTranAll Then
29                Call objHT.CloseTransactionBengkel
30                MsgBox "Data berhasil dihapus", vbInformation
31                Call refreshGrid
32                Call buka_Tutup(False)
33            Else
34                Call objHT.CloseTransactionBengkel_err
35                MsgBox "Gagal hapus data", vbInformation
36                Exit Sub
37            End If
38        End If

39       On Error GoTo 0
40       Exit Sub

cmdHapus_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
41        varErlError = Erl
42        varErrNumber = err.Number
43        If isBegin Then
44            Call objHT.CloseTransactionBengkel_err
45        End If
46        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmInfoKetersediaanInsentifCabang")
End Sub

Private Sub cmdImport_Click()
      Dim ErrBegin As Boolean, konter As Long
1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim rsTempExcel As ADODB.Recordset
         Dim Jml As Integer
         Dim itemPart As String
         Dim isBegin As Boolean
         Dim target, aktual, pencapaian As Double
         
2        isBegin = False
         
        ' Dim konter As Integer
         
3       konter = 0
4       ProgBar.Min = 0
5       CommonDialog1.FileName = ""
6       CommonDialog1.InitDir = App.Path
7       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
8       CommonDialog1.ShowOpen
9       If CommonDialog1.FileName <> "" Then
10         If CnExel.State Then CnExel.Close
11         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
12         GoSub Proses
13     Else
14        Exit Sub
15     End If
16     Exit Sub
       
Proses:
17    ErrBegin = True
         Dim VarNmFile As String
         Dim kodePosisi As String
         Dim kodeKPI, nik  As String
         Dim rsModel  As ADODB.Recordset
18       Call objHT.OpenTransactionBengkel
19       isBegin = True
20       gBantuan = "SBE"
21       If RcExcel.State Then RcExcel.Close
22       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
23       If RcExcel.RecordCount > 0 Then
24          ProgBar.Visible = True
25          ProgBar.Max = RcExcel.RecordCount
26          RcExcel.MoveFirst
27          Do While Not RcExcel.EOF
28             DoEvents
29             konter = konter + 1
30             ProgBar.value = konter
               
31             target = 0
32             aktual = 0
33             pencapaian = 0
34             kodeKPI = ""
35             nik = ""
               
              
36             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
               
37             aktual = NulSID(RcExcel.Fields(6).value)
38             kodeKPI = NulSID(RcExcel.Fields(4).value)
39             nik = NulSID(RcExcel.Fields(2).value)
40             target = objHT.FetchDataBengkel_String("select TargetKPIMP from INSENTIF..GRKPIResultMP where periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_Kd_Dealer = '" & gKdCbg & "' and NIK = '" & nik & "' and  KdKPI = '" & kodeKPI & "'")
41             If target > 0 Then
42                 pencapaian = Round((aktual * 100) / target, 2)
43             Else
44                pencapaian = 0
45             End If
               
46             sSQL = "UPDATE INSENTIF..GRKPIResultMP SET TargetKPIMP = '" & target & "',ActKPIMP = '" & aktual & "',AchvKPIMP = '" & k2t(pencapaian) & "' where Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_Kd_Dealer = '" & NulSID(gKdCbg) & "' and NIK = '" & NulSID(nik) & "' and KdKPI = '" & NulSID(kodeKPI) & "' "
47             Call objHT.InsertData_Bengkel(sSQL)
                
                       
48             RcExcel.MoveNext
49          Loop
50       End If
51       RcExcel.Close
52       CnExel.Close
      '   ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
53        Call objHT.CloseTransactionBengkel
54       MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
55       ProgBar.Visible = False
56       Call cmdViewKPI_Click
57       Exit Sub
vb_error:
58       If ErrBegin = True Then
           ' DbCon.RollbackTrans
59       End If
60       If isBegin Then
61        Call objHT.CloseTransactionBengkel_err
62       End If
63       ProgBar.Visible = False
64       MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdKPIMP_Click()
          Dim porsiInsentif As Double
1         sSQL = "select (SUM(jlhkredit)-SUM(jlhdebit)) From [acc" & Right(gDb, 2) & "-" & Format(Periode.value, "YYYYMM") & "].dbo.gltrn where kodegl "
2         sSQL = sSQL & " in( select kodegl From [acc" & Right(gDb, 2) & "-" & Format(Periode.value, "YYYYMM") & "].dbo.glmst where kodegl like '4310%' and "
3         sSQL = sSQL & " kodemaster = 'S')"
4         dppJasa = objO.FetchData_Double(sSQL)
5         sSQL = "select SUM(jlhdebit)-SUM(jlhkredit) From [acc" & Right(gDb, 2) & "-" & Format(Periode.value, "YYYYMM") & "].dbo.gltrn where kodegl "
6         sSQL = sSQL & " in( select kodegl From [acc" & Right(gDb, 2) & "-" & Format(Periode.value, "YYYYMM") & "].dbo.glmst where kodegl = '60401020')"
7         opexLembur = objO.FetchData_Double(sSQL)
          
8         sSQL = "Select PorsiInsentif from insentif..grAdtParameter where periode='" & Format(Periode.value, "MM-YYYY") & "'"
9         porsiInsentif = objO.FetchData_Double(sSQL)
10        ketersediaan = ((dppJasa - opexLembur) / 100) * porsiInsentif
11        totalBobot = 0
12        Text(0) = Format(dppJasa, "#,#0")
13        Text(1) = Format(opexLembur, "#,#0")
14        Text(2) = porsiInsentif
15        Text(3) = Format(ketersediaan, "#,#0")
          
16        Set rsTemp = objO.FetchData("Select tipe,bobotposisi from INSENTIF..GRTipeMP WHERE isnull(userdel,'') = ''")
17        If Not rsTemp.EOF Then
18            iProgress = 1
19            ProgBar.Visible = True
20            DoEvents
21            ProgBar.Min = 1
22            rsTemp.MoveFirst
23            ProgBar.Max = rsTemp.RecordCount + 1
              
24            Do While Not rsTemp.EOF
25                With rsGrid
26                    .AddNew
                      
27                    !tipe = NulSID(rsTemp!tipe)
28                    !BobotIns = k2t(rsTemp!bobotposisi)
29                    !jmlMp = objO.FetchData_Count("SELECT * FROM manpower WITH (NOLOCK) WHERE tipe = '" & NulSID(rsTemp!tipe) & "'")
30                    !jmlBobot = (!BobotIns * !jmlMp)
31                    totalBobot = totalBobot + !jmlBobot
32                    .Update
33                End With
34                ProgBar.value = iProgress
35                iProgress = iProgress + 1
36                DoEvents
37                rsTemp.MoveNext
38            Loop
              
39            Set GridSupplier.DataSource = rsGrid
40            isLoad = True
41            ProgBar.Visible = False
42        End If
43        cmdKPIMP.Enabled = False
44        cmdGenerateKPIMP.Enabled = True
          
End Sub

Private Sub cmdSimpan_Click()
   On Error GoTo cmdSimpan_Click_Error
    Dim boleh           As Boolean
    Dim pesan           As String
    Dim rsTemp          As ADODB.Recordset
    Dim namaSupplier    As String
    boleh = True
    isBegin = False
     
    If rsGrid.RecordCount > 0 Then
        rsGrid.MoveFirst
        Call objHT.OpenTransactionBengkel
        isBegin = True
        Do While Not rsGrid.EOF
            
            sSQL = "select * from INSENTIF..GRInsposisiTersedia where Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_Kd_Dealer = '" & NulSID(gKdCbg) & "' and Tipe = '" & NulSID(rsGrid!tipe) & "' "
            Set rsTemp = objHT.FetchDataBengkel(sSQL)
            If Not rsTemp.EOF Then
                sSQL = "UPDATE INSENTIF..GRInsposisiTersedia SET BobotIns = '" & k2t(rsGrid!BobotIns) & "', JmlMP = '" & k2t(rsGrid!jmlMp) & "', JmlBobot = '" & k2t(rsGrid!jmlBobot) & "', PstBobot = '" & k2t(rsGrid!jmlPersen) & "', PorsiPosisi = '" & k2t(rsGrid!porsiPosisi) & "', PorsiMP = '" & k2t(rsGrid!PorsiMP) & "' where Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_Kd_Dealer = '" & NulSID(gKdCbg) & "' and tipe = '" & NulSID(rsGrid!tipe) & "' "
            Else
            
                sSQL = "INSERT INTO INSENTIF..GRInsposisiTersedia (Periode , FS_KD_DEALER, tipe, BobotIns, jmlMp, jmlBobot, PstBobot, porsiPosisi, porsiMP, UserEntry, TglEntry) "
                sSQL = sSQL & " VALUES ('" & Format(Periode.value, "MM-YYYY") & "','" & NulSID(gKdCbg) & "','" & NulSID(rsGrid!tipe) & "','" & k2t(rsGrid!BobotIns) & "','" & k2t(rsGrid!jmlMp) & "','" & k2t(rsGrid!jmlBobot) & "','" & k2t(rsGrid!jmlPersen) & "','" & k2t(rsGrid!porsiPosisi) & "','" & k2t(rsGrid!PorsiMP) & "','" & gSA & "',getdate())"
            End If
            Call objHT.InsertData_Bengkel(sSQL)
            rsGrid.MoveNext
        Loop
        
        sSQL = "select * from INSENTIF..GRInsCbgTersedia where periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_Dealer = '" & NulSID(gKdCbg) & "'"
        Set rsTemp = objHT.FetchDataBengkel(sSQL)
        If Not rsTemp.EOF Then
            sSQL = "UPDATE INSENTIF..GRInsCbgTersedia SET dppJasa = " & k2t(Replace(Text(0), ".", "")) & ", opexLembur = " & k2t(Replace(Text(1), ".", "")) & ", InsCbgTersedia = " & k2t(Replace(Text(3), ".", "")) & " where periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_Dealer = '" & NulSID(gKdCbg) & "' "
        Else
            sSQL = "INSERT INTO INSENTIF..GRInsCbgTersedia (Periode , FS_KD_DEALER, dppJasa, opexLembur, InsCbgTersedia, userentry, tglentry) VALUES ('" & Format(Periode.value, "MM-YYYY") & "' , '" & NulSID(gKdCbg) & "', " & k2t(Replace(Text(0), ".", "")) & ", " & k2t(Replace(Text(1), ".", "")) & ", " & k2t(Replace(Text(3), ".", "")) & ", '" & gSA & "', getdate())"
        End If
        Call objHT.InsertData_Bengkel(sSQL)
        Call objHT.CloseTransactionBengkel
        Call cmdViewKPI_Click
    Else
        MsgBox "Tidak ada data yang diproses.", vbInformation
        Exit Sub
    End If
   On Error GoTo 0
   Exit Sub
cmdSimpan_Click_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    If isBegin = True Then
        Call objHT.CloseTransactionBengkel_err
    End If
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmInfoKetersediaanInsentifCabang")
End Sub

Private Sub cmdTambah_Click()
    isTambah = True
    Call ber_sih
    Call buka_Tutup(True)
    Call isiId
End Sub

Private Sub cmdTutup_Click()
    frmeModel.Visible = False
End Sub

Private Sub cmdTutupDealer_Click()

    frmeDealer.Visible = False
End Sub

Private Sub cmdUpdate_Click()
    Dim target, aktual, ach As Double
    If cmdGenerateKPIMP.Enabled = True Then
        rsGrid!target = NulSID(Text(4))
        rsGrid!actual = NulSID(Text(5))
        If rsGrid!target > 0 Then
            rsGrid!Achievement = Round((rsGrid!actual * 100) / rsGrid!target, 2)
        Else
            rsGrid!Achievement = 0
        End If
    Else
        target = k2t(Text(2))
        aktual = k2t(Text(3))
        If target > 0 Then
            ach = Round((aktual / target) * 100, 2)
        Else
            ach = 0
        End If
        Call objO.InsertData("UPDATE [INSENTIF].dbo.grKPIResultCbg SET TargetKPICbg = '" & k2t(target) & "',ActKpiCbg = '" & k2t(aktual) & "',AchvKpiCbg = '" & k2t(ach) & "' where KdKPI = '" & NulSID(Text(0)) & "' and kategori='0' and Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_DEALER = '" & NulSID(gKdCbg) & "'")
        Call cmdViewKPI_Click
    End If
    
End Sub

Private Sub cmdViewKPI_Click()
    Call ber_sih
    Call buka_Tutup(True)
    Call refreshGrid
End Sub

Private Sub Form_Load()
    Periode.value = Format(Date, "mm-yyyy")
    Call ber_sih
    Call buka_Tutup(True)
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
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure GridSupplier_HeadClick of Form frmInfoKetersediaanInsentifCabang")
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


Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
    Set rsTmp = objO.FetchDataHO("SELECT fs_kd_dealer, fs_nm_dealer, ISNULL(ID_PROGRAM,'') AS KODE FROM t_dealer a WITH (NOLOCK) " _
                            & "LEFT JOIN HO_PROCBG b ON a.fs_kd_dealer = b.kodedealer and A.FS_LOG_DEL = '' AND b.ID_PROGRAM = '" & Text(0) & "' WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
    If Not rsTmp.EOF Then
        Call createRS
        Do While Not rsTmp.EOF
            rsDealer.AddNew
            rsDealer!KodeDealer = rsTmp("FS_KD_DEALER")
            rsDealer!namaDealer = rsTmp("FS_NM_DEALER")
            If rsTmp("KODE") <> "" Then
                rsDealer!terlibat = True
            Else
                rsDealer!terlibat = False
            End If
            rsDealer.Update
            rsTmp.MoveNext
        Loop
        
        GridDealer.DataSource = rsDealer
        GridDealer.EvenRowStyle.BackColor = vbYellow
        GridDealer.OddRowStyle.BackColor = vbWhite
        GridDealer.AlternatingRowStyle = True
    End If
End Sub

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 3500
    GridDealer.Columns(2).Width = 1000
    GridDealer.Columns(2).Alignment = dbgCenter
End Sub

 

Private Sub refreshGrid()
          ''GridSupplier.ClearFields
1         Call createRS
          Dim porsiInsentif As Double
          ''sSQL = "SELECT KDSUPP,SUPPLIER,ALAMAT1,ALAMAT2,KOTA,TELPON,CP,KETERANGAN,KDDEALER,PKP,NONJURNAL FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY KDDEALER "
          
2         sSQL = "Select dppjasa,opexlembur,inscbgtersedia from  INSENTIF.dbo.GRInsCbgTersedia where Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_DEALER = '" & NulSID(gKdCbg) & "'"
3         iProgress = 1
4         ProgBar.Visible = True
5         DoEvents
6         ProgBar.Min = 1
7         Set rsTemp = objO.FetchData(sSQL)
8         If Not rsTemp.EOF Then
9             rsTemp.MoveFirst
              
10            sSQL = "Select PorsiInsentif from insentif..grAdtParameter where periode='" & Format(Periode.value, "MM-YYYY") & "'"
11            porsiInsentif = objO.FetchData_Double(sSQL)
12            Text(0) = Format(rsTemp!dppJasa, "#,#0")
13            Text(1) = Format(rsTemp!opexLembur, "#,#0")
14            Text(2) = porsiInsentif
15            Text(3) = Format(rsTemp!inscbgtersedia, "#,#0")
16        End If
          
17        sSQL = "Select tipe,bobotins,jmlMp,jmlBobot,pstbobot as jmlPersen,porsiPosisi,porsiMP from  INSENTIF.dbo.GRInsposisiTersedia where Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_DEALER = '" & NulSID(gKdCbg) & "'"
          
18        Set rsTemp = objO.FetchData(sSQL)
19        If Not rsTemp.EOF Then
20            cmdKPIMP.Enabled = False
21            cmdGenerateKPIMP.Enabled = True
22            cmdImport.Enabled = True
23            rsTemp.MoveFirst
24            ProgBar.Max = rsTemp.RecordCount + 1
25            Do While Not rsTemp.EOF
26                With rsGrid
27                    .AddNew
28                    !tipe = NulSID(rsTemp!tipe)
29                    !BobotIns = k2t(rsTemp!BobotIns)
30                    !jmlMp = NulSID(rsTemp!jmlMp)
31                    !jmlBobot = (rsTemp!jmlBobot)
32                    !jmlPersen = (rsTemp!jmlPersen)
33                    !porsiPosisi = (rsTemp!porsiPosisi)
34                    !PorsiMP = (rsTemp!PorsiMP)
35                    .Update
36                End With
37                ProgBar.value = iProgress
38                iProgress = iProgress + 1
39                DoEvents
                  
40                rsTemp.MoveNext
41            Loop
42            isLoad = True
43        Else
44            cmdImport.Enabled = False
45            cmdKPIMP.Enabled = True
46            cmdGenerateKPIMP.Enabled = False
47        End If
48        ProgBar.Visible = False
49        Set GridSupplier.DataSource = rsGrid
End Sub

Private Sub createRS()
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = Nothing
    With rsGrid
        .Fields.Append "tipe", adVarChar, 50
        .Fields.Append "BobotIns", adDouble
        .Fields.Append "jmlMp", adInteger
        .Fields.Append "jmlBobot", adDouble
        .Fields.Append "jmlPersen", adDouble
        .Fields.Append "porsiPosisi", adDouble
        .Fields.Append "porsiMP", adDouble
        .Open
    End With
End Sub
