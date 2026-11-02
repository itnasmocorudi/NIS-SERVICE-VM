VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmHOKelas 
   BackColor       =   &H80000016&
   Caption         =   "KELAS KENDARAAN"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmHOKelas.frx":0000
   ScaleHeight     =   9375
   ScaleWidth      =   15285
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   9045
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox cboWar 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   7320
         Width           =   800
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   5
         Top             =   7680
         Width           =   4815
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   1
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   2
         Top             =   6600
         Width           =   4815
      End
      Begin ucXPButton.XPButton cmdDelete 
         Height          =   495
         Left            =   5640
         TabIndex        =   9
         Top             =   8280
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&d"
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
         MICON           =   "frmHOKelas.frx":24D6
         PICN            =   "frmHOKelas.frx":24F2
         PICH            =   "frmHOKelas.frx":2A7A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSave 
         Height          =   495
         Left            =   4440
         TabIndex        =   10
         Top             =   8280
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
         MICON           =   "frmHOKelas.frx":30B6
         PICN            =   "frmHOKelas.frx":30D2
         PICH            =   "frmHOKelas.frx":35E5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdCreate 
         Height          =   495
         Left            =   4440
         TabIndex        =   11
         Top             =   8280
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&c"
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
         MICON           =   "frmHOKelas.frx":3B87
         PICN            =   "frmHOKelas.frx":3BA3
         PICH            =   "frmHOKelas.frx":40D8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   5415
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9551
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kelas"
         Columns(0).DataField=   "Kode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama"
         Columns(1).DataField=   "Kelas"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Harga"
         Columns(2).DataField=   "Harga"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Harga TAM"
         Columns(3).DataField=   "TAM"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Keterangan"
         Columns(4).DataField=   "Keterangan"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Disc NRM"
         Columns(5).DataField=   "DiscNRM"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Disc NON"
         Columns(6).DataField=   "DiscNon"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Warr."
         Columns(7).DataField=   "Warr"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1588"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1482"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(0).ButtonHeader=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4948"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4842"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(1).ButtonHeader=1"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2831"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2725"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2831"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2725"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=5292"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=5186"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(25)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1667"
         Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(30)=   "Column(6).Width=1773"
         Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1667"
         Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(35)=   "Column(7).Width=1244"
         Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=1138"
         Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         BorderStyle     =   0
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
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H80000016&,.bold=-1"
         _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2,.wraptext=0"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
      Begin ucXPButton.XPButton cmdCancel 
         Height          =   495
         Left            =   5640
         TabIndex        =   13
         Top             =   8280
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&c"
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
         MICON           =   "frmHOKelas.frx":46CB
         PICN            =   "frmHOKelas.frx":46E7
         PICH            =   "frmHOKelas.frx":4C2B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdAdd 
         Height          =   495
         Left            =   2040
         TabIndex        =   14
         Top             =   8280
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
         MICON           =   "frmHOKelas.frx":5222
         PICN            =   "frmHOKelas.frx":523E
         PICH            =   "frmHOKelas.frx":568A
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
         Left            =   3240
         TabIndex        =   15
         Top             =   8280
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&e"
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
         MICON           =   "frmHOKelas.frx":5B52
         PICN            =   "frmHOKelas.frx":5B6E
         PICH            =   "frmHOKelas.frx":6019
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdExit 
         Height          =   495
         Left            =   9960
         TabIndex        =   16
         Top             =   8280
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
         MICON           =   "frmHOKelas.frx":653E
         PICN            =   "frmHOKelas.frx":655A
         PICH            =   "frmHOKelas.frx":6A2E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBNumber6Ctl.TDBNumber txtRp 
         Height          =   315
         Index           =   10
         Left            =   2040
         TabIndex        =   3
         Top             =   6960
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   556
         Calculator      =   "frmHOKelas.frx":6F79
         Caption         =   "frmHOKelas.frx":6F99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOKelas.frx":7005
         Keys            =   "frmHOKelas.frx":7023
         Spin            =   "frmHOKelas.frx":706D
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12640511
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
      Begin TDBNumber6Ctl.TDBNumber txtRp 
         Height          =   315
         Index           =   11
         Left            =   2040
         TabIndex        =   4
         Top             =   7320
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   556
         Calculator      =   "frmHOKelas.frx":7095
         Caption         =   "frmHOKelas.frx":70B5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOKelas.frx":7121
         Keys            =   "frmHOKelas.frx":713F
         Spin            =   "frmHOKelas.frx":7189
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12640511
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
      Begin TDBNumber6Ctl.TDBNumber txtRp 
         Height          =   315
         Index           =   12
         Left            =   9240
         TabIndex        =   6
         Top             =   6600
         Width           =   765
         _Version        =   65536
         _ExtentX        =   1349
         _ExtentY        =   556
         Calculator      =   "frmHOKelas.frx":71B1
         Caption         =   "frmHOKelas.frx":71D1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOKelas.frx":723D
         Keys            =   "frmHOKelas.frx":725B
         Spin            =   "frmHOKelas.frx":72A5
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12640511
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "##,###0 %;##,###0 %;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1397882885
         MinValueVT      =   1112342533
      End
      Begin TDBNumber6Ctl.TDBNumber txtRp 
         Height          =   315
         Index           =   13
         Left            =   9240
         TabIndex        =   7
         Top             =   6960
         Width           =   765
         _Version        =   65536
         _ExtentX        =   1349
         _ExtentY        =   556
         Calculator      =   "frmHOKelas.frx":72CD
         Caption         =   "frmHOKelas.frx":72ED
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOKelas.frx":7359
         Keys            =   "frmHOKelas.frx":7377
         Spin            =   "frmHOKelas.frx":73C1
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12640511
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "##,###0 %;##,###0 %;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1397882885
         MinValueVT      =   1112342533
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   9120
         TabIndex        =   35
         Top             =   7320
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jasa Warranty"
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
         Left            =   7680
         TabIndex        =   34
         Top             =   7320
         Width           =   1410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   9120
         TabIndex        =   33
         Top             =   7005
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount NON"
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
         Left            =   7680
         TabIndex        =   32
         Top             =   6960
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   5
         Left            =   9120
         TabIndex        =   31
         Top             =   6600
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount NRM"
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
         Left            =   7680
         TabIndex        =   30
         Top             =   6600
         Width           =   1305
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   720
         TabIndex        =   29
         Top             =   7680
         Width           =   1125
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1920
         TabIndex        =   28
         Top             =   7680
         Width           =   60
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga TAM"
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
         Left            =   720
         TabIndex        =   27
         Top             =   7320
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1920
         TabIndex        =   26
         Top             =   7320
         Width           =   60
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
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
         Left            =   720
         TabIndex        =   25
         Top             =   6960
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   6960
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: LIST KELAS KENDARAAN ::."
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
         Left            =   3720
         TabIndex        =   23
         Top             =   60
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   11280
         Y1              =   8880
         Y2              =   8880
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   720
         TabIndex        =   22
         Top             =   6600
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas "
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
         Left            =   720
         TabIndex        =   21
         Top             =   6240
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: DATA KELAS KENDARAAN ::."
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
         Left            =   3840
         TabIndex        =   20
         Top             =   5820
         Width           =   3255
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   11280
         Y1              =   8160
         Y2              =   8160
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1920
         TabIndex        =   19
         Top             =   6240
         Width           =   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1920
         TabIndex        =   18
         Top             =   6600
         Width           =   60
      End
      Begin VB.Label lblJml 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   7680
         TabIndex        =   17
         Top             =   60
         Width           =   1380
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11055
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   11055
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
End
Attribute VB_Name = "frmHOKelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Job ID          :
'Tgl. Mulai      : 06 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private rsM As ADODB.Recordset
Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String
Private C As Integer
Private isOnWar As Boolean

Private Sub Form_Load()
    Call SetControlForm
    Call FillGridClass
End Sub

Private Sub SetControlForm()
    With Me
        .cmdAdd.Enabled = True
        .cmdCreate.Visible = False
        .cmdEdit.Enabled = True
        .cmdSave.Visible = True
        .cmdSave.Enabled = False
        .cmdCancel.Visible = False
        .cmdDelete.Visible = True
        .cmdDelete.Enabled = True
        .Grid.Enabled = True
        .txt(0).BackColor = fColor
        .txt(1).BackColor = fColor
        .txt(2).BackColor = fColor
        .txtRp(10).BackColor = fColor
        .txtRp(11).BackColor = fColor
        .txtRp(12).BackColor = fColor
        .txtRp(13).BackColor = fColor
        .cboWar.BackColor = fColor
        .cboWar.Clear
        .cboWar.AddItem "NO"
        .cboWar.AddItem "YES"
        .cboWar.Text = "NO"
        .cmdExit.Enabled = True
    End With
End Sub

Private Sub FillGridClass()
    Set rsM = objO.FetchDataHO("SELECT ClassCode as [KODE], ClassName as [KELAS], Price as [HARGA], PriceTAM as [TAM], Note as [KETERANGAN], " & _
                             "DiscNRM, DiscNON, War as Warr FROM CarClass WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY ClassCode")
    If rsM.RecordCount = 0 Then
        With Me
            .cmdAdd.Enabled = True
            .cmdCreate.Visible = False
            .cmdEdit.Enabled = False
            .cmdSave.Visible = True
            .cmdSave.Enabled = False
            .cmdCancel.Visible = False
            .cmdDelete.Visible = False
            .cmdDelete.Enabled = True
            .Grid.Enabled = False
            .Grid.DataSource = Nothing
            .lblJml.Caption = "0"
            .txt(0).BackColor = fColor
            .txt(1).BackColor = fColor
            .txt(2).BackColor = fColor
            .txtRp(10).BackColor = fColor
            .txtRp(11).BackColor = fColor
            .txtRp(12).BackColor = fColor
            .txtRp(13).BackColor = fColor
            .cboWar.BackColor = fColor
            .txt(0).Text = ""
            .txt(1).Text = ""
            .txt(2).Text = ""
            .txtRp(10) = 0
            .txtRp(11) = 0
            .txtRp(12) = 0
            .txtRp(13) = 0
            .cboWar.Text = "NO"
            .cmdExit.Enabled = True
        End With
        Exit Sub
    End If
    Set Grid.DataSource = rsM
    lblJml.Caption = rsM.RecordCount
 
        rsM.Find "Warr = 1", , adSearchForward, 1
        If Not rsM.EOF Then
            isOnWar = True
        Else
            isOnWar = False
        End If

End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call BindToControl(rsM)
End Sub
Private Sub BindToControl(rs As ADODB.Recordset)
On Error GoTo ErrNull
    txt(0).Text = rs.Fields(0).value
    txt(1).Text = rs.Fields(1).value
    txt(2).Text = rs.Fields(4).value
    txtRp(10) = Val(rs.Fields(2).value)
    txtRp(11) = Val(rs.Fields(3).value)
    txtRp(12) = Val(rs.Fields(5).value)
    txtRp(13) = Val(rs.Fields(6).value)
    cboWar.Text = IIf(Val(rs.Fields(7).value) = 0, "NO", "YES")
ErrNull:
End Sub

'===============================================================================================
'========= A D D ======== A D D ======== A D D =================================================
'===============================================================================================
Private Sub CmdAdd_Click()
    With Me
        .cmdAdd.Enabled = False
        .cmdCreate.Visible = True
        .cmdCreate.Enabled = True
        .cmdEdit.Enabled = False
        .cmdSave.Visible = False
        .cmdCancel.Visible = True
        .cmdCancel.Enabled = True
        .cmdDelete.Visible = False
        .Grid.Enabled = False
        .txt(0).BackColor = tColor
        .txt(1).BackColor = tColor
        .txt(2).BackColor = tColor
        .txtRp(10).BackColor = tColor
        .txtRp(11).BackColor = tColor
        .txtRp(12).BackColor = tColor
        .txtRp(13).BackColor = tColor
        .cboWar.BackColor = tColor
        .txt(0).Text = ""
        .txt(1).Text = ""
        .txt(2).Text = ""
        .txtRp(10) = 0
        .txtRp(11) = 0
        .txtRp(12) = 0
        .txtRp(13) = 0
        .cboWar.Text = "NO"
        .cmdExit.Enabled = False
        .txt(0).SetFocus
        
        If isOnWar = True Then
            .cboWar.BackColor = fColor
        Else
            .cboWar.BackColor = tColor
        End If
    End With
End Sub

Private Sub cmdCreate_Click()
    If FillNotNull = True Then
        rsM.Find "Kode = '" & Trim(txt(0).Text) & "'", , adSearchForward, 1
        If Not rsM.EOF Then
            MsgBox "Data Kelas " & txt(0).Text & " sudah terdaftar, Ganti Kode !", vbInformation + vbOKOnly, "Car Class"
            Exit Sub
        End If
        If MsgBox("Simpan Data Kelas Baru ?", vbQuestion + vbYesNo, "Car Class") = vbYes Then
            Call objO.InsertDataHO("INSERT INTO CarClass(ClassCode, ClassName, Price,priceTAM, Note, DiscNRM, DiscNON, war, FS_LOG_DEL, FD_TGL_DEL) " & _
                                   "VALUES('" & Trim(txt(0).Text) & "','" & Trim(txt(1).Text) & "'," & RepRp(txtRp(10)) & "," & RepRp(txtRp(11)) & ",'" & Trim(txt(2).Text) & "'," & RepRp(txtRp(12)) & "," & RepRp(txtRp(13)) & "," & IIf(cboWar.Text = "YES", 1, 0) & ",'','01/01/3000')")
            Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('KELAS', '" & Trim(txt(0).Text) & "','S')")
            Call SetControlForm
            Call FillGridClass
        End If
    Else
        Call NotValid_Input
    End If
End Sub

'===============================================================================================
'======= E D I T =========== E D I T =========== E D I T =======================================
'===============================================================================================
Private Sub cmdEdit_Click()
    With Me
        .cmdAdd.Enabled = False
        .cmdCreate.Visible = False
        .cmdEdit.Enabled = False
        .cmdSave.Visible = True
        .cmdSave.Enabled = True
        .cmdCancel.Visible = True
        .cmdCancel.Enabled = True
        .cmdDelete.Visible = False
        .Grid.Enabled = False
        .txt(0).BackColor = fColor
        .txt(1).BackColor = tColor
        .txt(2).BackColor = tColor
        .txtRp(10).BackColor = tColor
        .txtRp(11).BackColor = tColor
        .txtRp(12).BackColor = tColor
        .txtRp(13).BackColor = tColor
        .cboWar.BackColor = tColor
        .cmdExit.Enabled = False
        .txt(1).SetFocus
        
        If isOnWar = True Then
            If .cboWar.Text = "YES" Then
                .cboWar.BackColor = tColor
            Else
                .cboWar.BackColor = fColor
            End If
        Else
            .cboWar.BackColor = tColor
        End If
    End With
End Sub

Private Sub cmdSave_Click()
    If FillNotNull = True Then
        If MsgBox("Update Data Kelas '" & txt(0).Text & "' ?", vbQuestion + vbYesNo, "Car Class") = vbYes Then
            Call objO.InsertDataHO("Update CarClass Set  ClassName='" & Trim(txt(1).Text) & "', Price=" & RepRp(txtRp(10)) & ",priceTAM=" & RepRp(txtRp(11)) & ", Note='" & Trim(txt(2).Text) & "', DiscNRM=" & RepRp(txtRp(12)) & ", DiscNON=" & RepRp(txtRp(13)) & ", war=" & IIf(cboWar.Text = "YES", 1, 0) & " " & _
                                   "Where ClassCode ='" & Trim(txt(0).Text) & "' and  FS_LOG_DEL = ''")
            Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('KELAS', '" & Trim(txt(0).Text) & "','S')")
            Call SetControlForm
            Call FillGridClass
        End If
    Else
        Call NotValid_Input
    End If
End Sub

'===============================================================================================
'===============================================================================================
Private Function FillNotNull() As Boolean
    FillNotNull = True
    If txt(0).Text = "" Or txt(1).Text = "" Then
        FillNotNull = False
    End If
End Function
Private Sub NotValid_Input()
    MsgBox "Data Yang Anda Inputkan Belum Lengkap !", vbInformation + vbOKOnly, "Car Class"
End Sub
'===============================================================================================
'===============================================================================================

'===============================================================================================
'========= C A N C E L ========= C A N C E L ========= C A N C E L =============================
'===============================================================================================
Private Sub cmdCancel_Click()
    Call SetControlForm
    Call FillGridClass
End Sub

'===============================================================================================
'========= D E L E T E ========= D E L E T E ========= D E L E T E =============================
'===============================================================================================
Private Sub CmdDelete_Click()
    If MsgBox("Yakin, Hapus Data Kelas : " & txt(0).Text & " ?", vbQuestion + vbYesNo, "Car Class") = vbYes Then
        Call objO.UpdateDataHO("UPDATE CarClass SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE ClassCode = '" & Trim(txt(0).Text) & "' and FS_LOG_DEL = ''")
        Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('KELAS', '" & Trim(txt(0).Text) & "','D')")
        Call SetControlForm
        Call FillGridClass
    End If
End Sub

Private Sub grid_FilterChange()
On Error GoTo err
   Set cols = Grid.Columns
   
   C = Grid.col
   Grid.HoldFields
   rsM.filter = getFilter()
   If rsM.EOF = True Then
      For Each col In Grid.Columns
         col.FilterText = ""
      Next col
      rsM.filter = ""
   End If
   Grid.col = C
   Grid.EditActive = True
   lblJml.Caption = rsM.RecordCount
   Exit Sub
err:
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

Private Sub grid_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 6 Then rsM.Sort = "[" & Grid.Columns(ColIndex).DataField & "]" & " " & Urut
4        Grid.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        lblJml.Caption = rsM.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter ")
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If txt(Index).BackColor = fColor Then Sendkeys ("{TAB}")
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys ("{TAB}")
End Sub


Private Sub txtRp_GotFocus(Index As Integer)
    If txtRp(Index).BackColor = fColor Then Sendkeys ("{TAB}")
End Sub

Private Sub txtRp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys ("{TAB}")
End Sub


Private Sub cboWar_GotFocus()
    If cboWar.BackColor = fColor Then Sendkeys ("{TAB}")
End Sub

Private Sub cboWar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys ("{TAB}")
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub



'Dim rsGrid As New ADodb.Recordset
'Dim rsTemp As ADodb.Recordset
'Dim cols As TrueOleDBGrid70.Columns
'
'Private Sub ber_sih()
'    For i = 0 To 4
'        Text(i) = ""
'    Next
'    chWar = Unchecked
'End Sub
'
'Private Sub buka_Tutup(Optional Ket As Boolean = True, Optional Flg As Boolean = True)
'    Text(0).Enabled = Ket
'    Text(1).Enabled = Flg
'    Text(2).Enabled = Flg
'    Text(3).Enabled = Flg
'    Text(4).Enabled = Flg
'    chWar.Enabled = Flg
'End Sub
'
'Private Sub refreshGrid()
'    GridKelas.ClearFields
'    Set rsGrid = objO.FetchDataHO("SELECT ClassCode as [KODE], ClassName as [KELAS], Price as [HARGA], Note as [KETERANGAN], PriceTAM as [TAM] FROM CarClass WHERE FS_LOG_DEL = '' ORDER BY ClassCode")
'    If Not rsGrid.EOF Then
'        GridKelas.DataSource = rsGrid
'        GridKelas.Columns(0).Width = 1000
'        GridKelas.Columns(1).Width = 3000
'        GridKelas.Columns(2).Width = 1000
'        GridKelas.Columns(2).NumberFormat = "#,#0"
'        GridKelas.Columns(3).Width = 3500
'        GridKelas.Columns(4).Width = 1000
'        GridKelas.Columns(4).NumberFormat = "#,#0"
'    End If
'End Sub
'
'Private Sub chWar_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call CekTombol(KeyCode)
'End Sub
'
'Private Sub cmdCetak_Click()
'    frmHOKelasCtk.Show
'End Sub
'
'Private Sub cmdHapus_Click()
'    If Trim(Text(1)) <> "" Then
'        temp = MsgBox("Yakin akan menghapus Kelas " & Trim(Text(1)) & " ?", vbYesNo, "HAPUS DATA")
'        If temp = vbYes Then
'            Call objO.UpdateDataHO("UPDATE CarClass SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE ClassCode = '" & Trim(Text(0)) & "' and FS_LOG_DEL = ''")
'            Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('KELAS', '" & Text(0) & "','D')")
'        End If
'    Else
'        MsgBox "Pilih Data yang akan dihapus"
'        Exit Sub
'    End If
'    Call ber_sih
'    Call refreshGrid
'End Sub
'
'Private Sub cmdKeluar_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdSimpan_Click()
'Dim boleh As Boolean
'Dim pesan As String
'
'    If cmdSimpan.Caption = "&Tambah" Then
'        Call ber_sih
'        Call buka_Tutup
'        Text(0).SetFocus
'        cmdSimpan.Caption = "&Simpan"
'        cmdUbah.Caption = "&Batal"
'        Me.Tag = "baru"
'    ElseIf cmdSimpan.Caption = "&Simpan" Then
'
'        boleh = True
'        pesan = "Lengkapi Data berikut "
'
'        If Text(0) = "" Then
'            boleh = False
'            pesan = pesan + "Kelas. "
'        End If
'        If Text(1) = "" Then
'            boleh = False
'            pesan = pesan + "Nama. "
'        End If
'        If Text(2) = 0 Then
'            boleh = False
'            pesan = pesan + "Harga. "
'        End If
'
'        If boleh Then
'            ' PROSES SIMPAN
'            Set rsTemp = objO.FetchDataHO("SELECT * FROM carclass WHERE ClassCode = '" & Trim(Text(0)) & "' and FS_LOG_DEL = '' ORDER BY classCode")
'            If Not rsTemp.EOF Then
'                ' SUDAH ADA DATA
'                Call objO.UpdateDataHO("UPDATE CarClass SET className = '" & Trim(Text(1)) & "', price = '" & CDbl(Text(2)) & "', pricetam = '" & CDbl(Text(4)) & "', note = '" & Trim(Text(3)) & "', war = '" & IIf(chWar = Checked, 1, 0) & "' WHERE classCode ='" & Trim(Text(0)) & "' and FS_LOG_DEL = ''")
'            Else
'                ' INPUT DATA BARU
'                Call objO.InsertDataHO("INSERT INTO CarClass VALUES ('" & Trim(Text(0)) & "','" & Trim(Text(1)) & "','" & Trim(Text(2)) & "','" & Trim(Text(3)) & "', '" & IIf(chWar = Checked, 1, 0) & "','','','" & CDbl(Text(4)) & "')")
'            End If
'
'            Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('KELAS', '" & Trim(Text(0)) & "','S')")
'            Call ber_sih
'            Call buka_Tutup(False, False)
'            Call refreshGrid
'
'            cmdSimpan.Caption = "&Tambah"
'            cmdUbah.Caption = "&Ubah"
'        Else
'            MsgBox pesan, vbCritical, "KELENGKAPAN"
'            Text(1).SelLength = Len(Text(1))
'            Text(1).SetFocus
'            Exit Sub
'        End If
'    End If
'End Sub
'
'Private Sub cmdUbah_Click()
'    If cmdUbah.Caption = "&Ubah" Then
'        Call buka_Tutup(False, True)
'        cmdSimpan.Caption = "&Simpan"
'        cmdUbah.Caption = "&Batal"
'        Me.Tag = "ubah"
'    ElseIf cmdUbah.Caption = "&Batal" Then
'        Call buka_Tutup(False, False)
'        cmdSimpan.Caption = "&Tambah"
'        cmdUbah.Caption = "&Ubah"
'    End If
'End Sub
'
'Private Sub Form_Load()
'    Call ber_sih
'    Call buka_Tutup(False, False)
'    Call refreshGrid
'End Sub
'
'Private Sub Form_Resize()
'    asxPowerBanner1.Height = Me.Height - 600
'    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
'    Frame1.Width = asxBars1.Width
'    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
'    GridKelas.Width = Frame1.Width - 200
'    GridKelas.Height = Frame1.Height - Text(3).Top - Text(3).Height - 300
'    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
'    cmdHapus.Top = cmdKeluar.Top
'    cmdUbah.Top = cmdKeluar.Top
'    cmdSimpan.Top = cmdKeluar.Top
'    cmdCetak.Top = cmdKeluar.Top
'    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
'    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
'    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
'    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
'    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
'End Sub
'
'Private Sub GridKelas_FilterChange()
'    On Error GoTo errHandler
'    Set cols = GridKelas.Columns
'    Dim C As Integer
'
'    C = GridKelas.col
'    GridKelas.HoldFields
'    rsGrid.filter = getFilter()
'    GridKelas.col = C
'    GridKelas.EditActive = True
'    Exit Sub
'
'errHandler:
'    MsgBox err.Source & ":" & vbCrLf & err.Description
'End Sub
'
'Private Function getFilter() As String
'Dim tmp As String
'Dim n As Integer
'    For Each col In cols
'        If Trim(col.FilterText) <> "" Then
'            n = n + 1
'            If n > 1 Then
'                tmp = tmp & " AND "
'            End If
'            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
'
'        End If
'    Next col
'    getFilter = tmp
'End Function
'
'Private Sub GridKelas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If Not rsGrid.EOF Then
'        Text(0) = rsGrid.Fields(0)    'GridKelas.Columns(0)
'        Text(1) = rsGrid.Fields(1)    'GridKelas.Columns(1)
'        Text(2) = Format(rsGrid.Fields(2), "#,#0")   'GridKelas.Columns(2)
'        Text(3) = rsGrid.Fields(3)    'GridKelas.Columns(3)
'        Text(4) = Format(rsGrid.Fields(4), "#,#0")
'    End If
'End Sub
'
'Private Sub Text_GotFocus(Index As Integer)
'    Dim DiTab As Boolean
'    DiTab = False
'
'    If Text(Index).BackColor = &HC0E0FF Then
'        DiTab = True
'    Else
'        If gMainKeyCari <> "" Then
'            Text(Index) = gMainKeyCari
'            gMainKeyCari = ""
'            DiTab = True
'        End If
'    End If
'    If DiTab Then
'        SendKeys ("{TAB}")
'    Else
'        SendKeys ("{HOME}+{END}")
'    End If
'End Sub
'
'Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    Call CekTombol(KeyCode)
'    If KeyCode = vbKeyF12 Then
'        Select Case Index
'            Case 0
'                gMainKeyCari = Text(Index)
'                gSearchOption = "KELAS_MOBIL"
'                frmCariHO.Show (vbModal)
'        End Select
'    End If
'End Sub
'
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 39 Then
'        KeyAscii = 0
'        SendKeys ("´")
'    End If
'    If KeyAscii = 34 Then
'        KeyAscii = 0
'        SendKeys ("”")
'    End If
'End Sub
'
'Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
'    Select Case Index
'        Case 0
'            If Text(Index) <> "" Then
'                Set oRs = objO.FetchDataHO("Select ClassName from carClass where classCode ='" & Text(0) & "' and FS_LOG_DEL = ''")
'                If Not oRs.BOF Then
'                    Text(1) = oRs("ClassName")
'                End If
'                oRs.Close
'            End If
'        Case 2
'            If Text(Index) = "" Then
'                Text(Index) = 0
'                Text(Index).SetFocus
'                Exit Sub
'            ElseIf Text(Index) <= 0 Then
'                MsgBox "Harga Kelas tidak boleh kosong !!!", vbCritical, "KELENGKAPAN"
'                Text(Index).SetFocus
'                Exit Sub
'            End If
'    End Select
'End Sub
