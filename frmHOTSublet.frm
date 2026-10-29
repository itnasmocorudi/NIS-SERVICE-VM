VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form frmHOTSublet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Master Sublet"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   11850
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   5760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   16744576
      Caption         =   "Import Excell"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   5175
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   11295
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   450
         Calculator      =   "frmHOTSublet.frx":0000
         Caption         =   "frmHOTSublet.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOTSublet.frx":008C
         Keys            =   "frmHOTSublet.frx":00AA
         Spin            =   "frmHOTSublet.frx":00F4
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
         MaxValue        =   999999999999999
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
         MaxValueVT      =   1869742085
         MinValueVT      =   1280311301
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   600
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CC"
         PromptChar      =   ">"
      End
      Begin TrueOleDBGrid70.TDBGrid GridJobOrder 
         Height          =   3015
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5318
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ID"
         Columns(0).DataField=   "ID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "TYPE"
         Columns(1).DataField=   "TYPE"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "DESC"
         Columns(2).DataField=   "DESC"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "HARGA BELI"
         Columns(3).DataField=   "HARGABELI"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "HARGA JUAL"
         Columns(4).DataField=   "HARGAJUAL"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "CREATE BY"
         Columns(5).DataField=   "CREATEBY"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TGL ENTRY"
         Columns(6).DataField=   "TGLCREATE"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "UPDATE BY"
         Columns(7).DataField=   "UPDATEBY"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TGL UPDATE"
         Columns(8).DataField=   "TGLUPDATE"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "ISHAPUS"
         Columns(9).DataField=   "ISHAPUS"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
         Splits(0)._ColumnProps(40)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H77EAE1&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
         _StyleDefs(76)  =   "Named:id=33:Normal"
         _StyleDefs(77)  =   ":id=33,.parent=0"
         _StyleDefs(78)  =   "Named:id=34:Heading"
         _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   ":id=34,.wraptext=-1"
         _StyleDefs(81)  =   "Named:id=35:Footing"
         _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   "Named:id=36:Selected"
         _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=37:Caption"
         _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(87)  =   "Named:id=38:HighlightRow"
         _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(89)  =   "Named:id=39:EvenRow"
         _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(91)  =   "Named:id=40:OddRow"
         _StyleDefs(92)  =   ":id=40,.parent=33"
         _StyleDefs(93)  =   "Named:id=41:RecordSelector"
         _StyleDefs(94)  =   ":id=41,.parent=34"
         _StyleDefs(95)  =   "Named:id=42:FilterBar"
         _StyleDefs(96)  =   ":id=42,.parent=33"
         _StyleDefs(97)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(98)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(99)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(100) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1680
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   450
         Calculator      =   "frmHOTSublet.frx":011C
         Caption         =   "frmHOTSublet.frx":013C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOTSublet.frx":01A8
         Keys            =   "frmHOTSublet.frx":01C6
         Spin            =   "frmHOTSublet.frx":0210
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
         MaxValue        =   999999999999999
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
         MaxValueVT      =   1869742085
         MinValueVT      =   1280311301
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID : "
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
         Left            =   1410
         TabIndex        =   17
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HARGA JUAL : "
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
         Left            =   540
         TabIndex        =   13
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HARGA BELI : "
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
         Left            =   615
         TabIndex        =   12
         Top             =   1380
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESC : "
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
         Left            =   1185
         TabIndex        =   11
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE : "
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
         Left            =   1200
         TabIndex        =   6
         Top             =   660
         Width           =   540
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
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   7080
      TabIndex        =   1
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   65280
      BorderStyle     =   3
      Caption         =   "&Tambah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdUbah 
      Height          =   360
      Left            =   8040
      TabIndex        =   3
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   65535
      BorderStyle     =   3
      Caption         =   "&Ubah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   9000
      TabIndex        =   4
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   255
      BorderStyle     =   3
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdKeluar 
      Height          =   360
      Left            =   9960
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin ARProgBarCtrl.ARProgressBar ProgBar 
      Height          =   255
      Left            =   375
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   450
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
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   6195
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   10927
      StartColor      =   49152
      EndColor        =   12648447
      FormatString    =   "    Table Master Sublet"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   4194304
   End
End
Attribute VB_Name = "frmHOTSublet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOTSublet
' DateTime  : 27/07/2015 11:18
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------

Dim rsGrid As New ADODB.Recordset       ' Recordset Grid
Dim rsTemp As ADODB.Recordset           ' Recordset Proses
Dim cols As TrueOleDBGrid70.Columns
Dim boleh As Boolean
Dim pesan As String
Dim gAtur As Boolean
'Dim sTampung As String

Private Sub ber_sih()
    For i = 0 To 2
        Text(i) = ""
    Next
End Sub

Private Sub buka_Tutup(Optional Ket As Boolean = True, Optional stat As Boolean = True)
    Text(2).Enabled = Ket
    Text(1).Enabled = stat
    Text(0).Enabled = stat
    TDBNumber1.Enabled = stat
    TDBNumber2.Enabled = stat
End Sub

Private Sub refreshGrid()
    gridJobOrder.ClearFields
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = objO.FetchDataHO("SELECT * FROM MASTER_SUBLET WITH(NOLOCK) WHERE ISHAPUS = '0' ORDER BY ID DESC ")
    gridJobOrder.DataSource = rsGrid
End Sub

Private Sub cboKM_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub chSBE_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub cmdCetak_Click()
    frmHOTSubletCtk.Show
End Sub

Private Sub cmdHapus_Click()
    If Trim(Text(2)) <> "" Then
        temp = MsgBox("Yakin akan menghapus data ID = " & Trim(Text(2)) & " ?", vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE master_SUBLET SET ishapus = '1', tglupdate = getdate(),updateby = '" & gSA & "' WHERE id = '" & Trim(Text(2)) & "' and ishapus = '0'")
            Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('M_SUBLET','" & Trim(Text(2)) & "','S')")
            Call ber_sih
            Call refreshGrid
        End If
    Else
        MsgBox "Pilih Data yang akan dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdImport_Click()
Dim ErrBegin As Boolean, konter As Long
  On Error GoTo vb_error
   Dim CnExel As New ADODB.Connection
   Dim RcExcel As New ADODB.Recordset
   Dim Jml As Integer
  konter = 0
  ProgBar.Min = 0
  CommonDialog1.FileName = ""
  CommonDialog1.InitDir = App.Path
  CommonDialog1.filter = "Exel File (*.xls)|*.xls"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then
     If CnExel.State Then CnExel.Close
     CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
     GoSub ProsesImport
 Else
    Exit Sub
 End If
 Exit Sub
 
ProsesImport:
ErrBegin = True
   Dim VarNmFile As String
   Call objHT.OpenTransactionHO
   If RcExcel.State Then RcExcel.Close
   RcExcel.Open "Select * from [Sheet1$]", CnExel, adOpenKeyset
   If RcExcel.RecordCount > 0 Then
      ProgBar.Visible = True
      ProgBar.Max = RcExcel.RecordCount
      RcExcel.MoveFirst
      Do Until RcExcel.EOF
         DoEvents
         konter = konter + 1
         ProgBar.value = konter
         If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
'         If objO.FetchDataHO_Count("select id from master_t1 WITH(NOLOCK) WHERE ishapus = '0' and id = '" & Trim(RcExcel.Fields(0).value) & "'") > 0 Then
'            Call objHT.CloseTransactionHO_err
'            ProgBar.Visible = False
'            MsgBox "ID " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " sudah ada!" & vbCrLf & "Proses dibatalkan", vbCritical, "Error"
'            Exit Sub
'         End If
         
         If Nul(RcExcel.Fields(1).value) = "" Then
            Call objHT.CloseTransactionHO_err
            ProgBar.Visible = False
            MsgBox "TYPE baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
         End If
         
         If Nul(RcExcel.Fields(2).value) = "" Then
            Call objHT.CloseTransactionHO_err
            ProgBar.Visible = False
            MsgBox "DESC baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
         End If
         
         Set rsTemp = objO.FetchDataHO("select id from master_sublet WITH(NOLOCK) WHERE ishapus = '0' and [type] = '" & Trim(RcExcel.Fields(1).value) & "' and [desc] = '" & Trim(RcExcel.Fields(2).value) & "'")
        If Not rsTemp.EOF Then
        'UPDATE DATA
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                Call objHT.UpdateData_Ho("UPDATE MASTER_SUBLET SET [type] = '" & Nul(RcExcel.Fields(1).value) & "',[desc] = '" & Nul(RcExcel.Fields(2).value) & "',HARGABELI = " & Replace(Nul(RcExcel.Fields(7).value), ",", ".") & " ,HARGAJUAL = " & Replace(Nul(RcExcel.Fields(8).value), ",", ".") & ",tglupdate = getdate(),updateby = '" & gSA & "' WHERE ID = '" & Trim(rsTemp!id) & "' AND ISHAPUS = '0'")
                Call objHT.UpdateData_Ho("INSERT INTO TEMPKIRIM VALUES ('M_SUBLET','" & Trim(rsTemp!id) & "','S')")
                rsTemp.MoveNext
            Loop
            
        Else
        ' INPUT DATA BARU
            Call objHT.UpdateData_Ho("INSERT INTO [MASTER_SUBLET]([TYPE],[DESC],[HARGABELI],[HARGAJUAL],[CREATEBY],[TGLCREATE],[UPDATEBY],[TGLUPDATE],[ISHAPUS]) VALUES ('" & Nul(RcExcel.Fields(1).value) & "','" & Nul(RcExcel.Fields(2).value) & "'," & IIf(Nul(RcExcel.Fields(7).value) = "", 0, Nul(RcExcel.Fields(7).value)) & "," & IIf(Nul(RcExcel.Fields(8).value) = "", 0, Nul(RcExcel.Fields(8).value)) & ",'" & gSA & "',getdate(),'','','0')")
            Set rsTemp = objO.FetchDataHO("select id from master_sublet WITH(NOLOCK) WHERE ishapus = '0' and [type] = '" & Trim(RcExcel.Fields(1).value) & "' and [desc] = '" & Trim(RcExcel.Fields(2).value) & "'")
            If Not rsTemp.EOF Then
                Call objHT.UpdateData_Ho("INSERT INTO TEMPKIRIM VALUES ('M_SUBLET','" & Trim(rsTemp!id) & "','S')")
            End If
            
        End If
        '
        RcExcel.MoveNext
      Loop
   End If
   RcExcel.Close
   CnExel.Close
'   ProgBar.Visible = False
'  ' IsiGrid
'   DbCon.CommitTrans
'   ErrBegin = False
    Call objHT.CloseTransactionHO
    
   MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
   Call refreshGrid
   ProgBar.Visible = False
'   GridAFI.Refresh
   Return
   Exit Sub
vb_error:
   If ErrBegin = True Then
     ' DbCon.RollbackTrans
   End If
   ProgBar.Visible = False
   MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim jobOrd As String

    If cmdSimpan.Caption = "&Tambah" Then
        Call ber_sih
        
        Call buka_Tutup(False, True)
        Text(0).SetFocus
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
        Me.Tag = "baru"
    ElseIf cmdSimpan.Caption = "&Simpan" Then
        boleh = True
        
        ' CHECK DATA KOSONG
        pesan = ""
        
        If Trim(Text(0)) = "" Then
            boleh = False
            pesan = pesan + "ID. "
        End If
        
        If Trim(Text(1)) = "" Then
            boleh = False
            pesan = pesan + "Desc. "
        End If
        
        
        ' JIKA DATA VALID
        If boleh Then
           
            ' PROSES SIMPAN
            If Trim(Text(2)) <> "" Then
                'UPDATE DATA
                Call objO.UpdateDataHO("UPDATE MASTER_SUBLET SET [TYPE] = '" & Trim(Text(0)) & "',[DESC] = '" & Trim(Text(1)) & "',HARGABELI = " & Replace(TDBNumber1.value, ",", ".") & ",HARGAJUAL = " & Replace(TDBNumber2.value, ",", ".") & ",UPDATEBY = '" & gSA & "',TGLUPDATE = GETDATE() WHERE ID = '" & Trim(Text(2)) & "'")
            Else
                ' INPUT DATA BARU
                Call objO.UpdateDataHO("INSERT INTO [MASTER_SUBLET]([TYPE],[DESC],[HARGABELI],[HARGAJUAL],[CREATEBY],[TGLCREATE],[UPDATEBY],[TGLUPDATE],[ISHAPUS]) VALUES ('" & Trim(Text(0)) & "','" & Trim(Text(1)) & "'," & Replace(TDBNumber1.value, ",", ".") & "," & Replace(TDBNumber2.value, ",", ".") & ",'" & gSA & "',getdate(),'','','0')")
            End If
            Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('M_SUBLET','" & Trim(Text(2)) & "','S')")
            Call ber_sih
            Call buka_Tutup(False, False)
            Call refreshGrid
            cmdSimpan.Caption = "&Tambah"
            cmdUbah.Caption = "&Ubah"
        Else
            MsgBox pesan, vbExclamation, "KELENGKAPAN"
        End If
    End If
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call buka_Tutup(False, True)
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
        Me.Tag = "ubah"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdSimpan.Caption = "&Tambah"
        cmdUbah.Caption = "&Ubah"
    End If
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    gridJobOrder.Width = Frame1.Width - 200
    gridJobOrder.Height = Frame1.Height - TDBNumber2.Top - TDBNumber2.Height - 200
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    ProgBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdImport.Left = cmdSimpan.Left - cmdImport.Width - 100
    ProgBar.Left = cmdImport.Left - ProgBar.Width - 100
End Sub

Private Sub GridJobOrder_FilterChange()
    On Error GoTo errHandler
    Set cols = gridJobOrder.Columns
    Dim C As Integer
    
    C = gridJobOrder.col
    gridJobOrder.HoldFields
    rsGrid.filter = getFilter()
    gridJobOrder.col = C
    gridJobOrder.EditActive = True
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

Private Sub GridJobOrder_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsGrid.State <> adStateClosed Then
        If Not rsGrid.EOF Then
            'If GridJobOrder.row > 0 Then
            Text(0) = rsGrid.Fields(1)
            Text(2) = rsGrid.Fields(0) 'GridJobOrder.Columns(0)
            Text(1) = rsGrid.Fields(2)   'gridJobOrder.Columns(2)
            'chSBE.value = IIf(rsGrid.Fields(3) = "Y", Checked, Unchecked)   'gridJobOrder.Columns(3)
            TDBNumber1.value = rsGrid.Fields(3)
            TDBNumber2.value = rsGrid.Fields(4)
            
        End If
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCariHO.Show (vbModal)
            Case 1
                If Text(0) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = Text(0)
                    frmCariHO.Show (vbModal)
                Else
                    MsgBox "Isi Tipe Job Order", vbCritical, "KELENGKAPAN"
                    Text(0).SetFocus
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
    If Index = 0 Then
        If Me.Tag = "baru" Then
            Set rsTemp = objO.FetchDataHO("SELECT ID FROM MASTER_T1 WHERE ISHAPUS = '0' AND ID = '" & Trim(Text(0)) & "'")
            If Not rsTemp.EOF Then
                MsgBox "Id = " & Trim(Text(0)) & " sudah ada silahkan ganti dengan yang lain.", vbInformation
                Text(0) = ""
                Exit Sub
            End If
        End If
    End If
End Sub
