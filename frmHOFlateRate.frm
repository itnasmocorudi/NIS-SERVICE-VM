VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form frmHOFlateRate 
   BackColor       =   &H00FFFFFF&
   Caption         =   "FLATE RATE"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   11835
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12240
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   6615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   11295
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   7440
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   3375
         Begin VB.OptionButton opBP 
            BackColor       =   &H00FFFFFF&
            Caption         =   "REPAIR"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   26
            Top             =   960
            Width           =   2535
         End
         Begin VB.OptionButton opBP 
            BackColor       =   &H00FFFFFF&
            Caption         =   "REPLACE && REPAIR"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   25
            Top             =   600
            Width           =   2535
         End
         Begin VB.OptionButton opBP 
            BackColor       =   &H00FFFFFF&
            Caption         =   "REPAIR && PAINT"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   3600
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">999999999"
         PromptChar      =   ">"
      End
      Begin VB.CheckBox chGroup 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   7440
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   4
         Left            =   3600
         TabIndex        =   4
         Top             =   600
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin TrueOleDBGrid70.TDBGrid GridFlate 
         Height          =   3855
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6800
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "GROUP"
         Columns(0).DataField=   "GROUP"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KODEMDL"
         Columns(1).DataField=   "KODEMDL"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "MDL"
         Columns(2).DataField=   "MDL"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "JOB"
         Columns(3).DataField=   "JOB"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TIPE"
         Columns(4).DataField=   "TIPE"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "ORDER"
         Columns(5).DataField=   "ORDER"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "RATE"
         Columns(6).DataField=   "RATE"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "KETERANGAN"
         Columns(7).DataField=   "KETERANGAN"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   4
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "CHECK"
         Columns(8).DataField=   "CHECK"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=3281"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8196"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=8196"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=3281"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=8196"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
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
         DeadAreaBackColor=   6876390
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0"
         _StyleDefs(74)  =   "Named:id=34:Heading"
         _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=34,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=35:Footing"
         _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   "Named:id=36:Selected"
         _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=37:Caption"
         _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(83)  =   "Named:id=38:HighlightRow"
         _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=39:EvenRow"
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
         _StyleDefs(93)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(94)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(95)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(96)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   1
         Left            =   3615
         TabIndex        =   17
         Top             =   240
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99,99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   7440
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   9
         Left            =   7680
         TabIndex        =   22
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   3
         Left            =   11160
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin VB.Frame Frame3 
         BackColor       =   &H0068ECE6&
         Height          =   735
         Left            =   5400
         TabIndex        =   28
         Top             =   960
         Width           =   5655
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tipe Job BP"
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
            Index           =   7
            Left            =   3960
            TabIndex        =   37
            Top             =   405
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Keterangan"
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
            Index           =   6
            Left            =   3960
            TabIndex        =   36
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Rate"
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
            Index           =   5
            Left            =   2760
            TabIndex        =   34
            Top             =   405
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Job Order"
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
            Index           =   4
            Left            =   2760
            TabIndex        =   33
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tipe Job"
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
            Index           =   3
            Left            =   1560
            TabIndex        =   32
            Top             =   405
            Width           =   1095
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
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
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   31
            Top             =   120
            Width           =   975
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Kode Model"
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
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   400
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Group Model"
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
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   1215
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   39
         Top             =   1920
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3240
         Top             =   1920
         Width           =   60
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Job, Model :"
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
         TabIndex        =   40
         Top             =   1965
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Order"
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
         Left            =   2730
         TabIndex        =   12
         Top             =   645
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Job"
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
         Left            =   6675
         TabIndex        =   21
         Top             =   645
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   10800
         Top             =   600
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Left            =   6840
         TabIndex        =   20
         Top             =   285
         Width           =   435
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   9960
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Left            =   1065
         TabIndex        =   19
         Top             =   1005
         Width           =   345
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3495
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label16 
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
         Left            =   990
         TabIndex        =   18
         Top             =   285
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Job Order"
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
         TabIndex        =   13
         Top             =   645
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   570
         TabIndex        =   11
         Top             =   1365
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   2280
         Top             =   600
         Width           =   60
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   6480
         Top             =   600
         Width           =   60
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7395
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13044
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "    Tabel Flate Rate"
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
      Left            =   7200
      TabIndex        =   9
      Top             =   7200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
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
      Left            =   8160
      TabIndex        =   14
      Top             =   7200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
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
      Left            =   9120
      TabIndex        =   15
      Top             =   7200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
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
      Left            =   10080
      TabIndex        =   16
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12648447
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
   Begin AIFCmp1.asxPowerButton cmdCetak 
      Height          =   360
      Left            =   5880
      TabIndex        =   35
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Cetak"
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
   Begin AIFCmp1.asxPowerButton cmdDisplay 
      Height          =   360
      Left            =   480
      TabIndex        =   38
      Top             =   7200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Display All"
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
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   5625
      TabIndex        =   41
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   12648447
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
   Begin ARProgBarCtrl.ARProgressBar ProgBar 
      Height          =   255
      Left            =   1800
      TabIndex        =   42
      Top             =   7200
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
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
End
Attribute VB_Name = "frmHOFlateRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsGrid As New ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim newRecord As Boolean
Dim pesan As String
Dim boleh As Boolean
Dim SSQL2 As String
Dim statusSBE As String
Dim cols As TrueOleDBGrid70.Columns

Private Sub ber_sih()
1        On Error GoTo ber_sih_Error

2         For i = 0 To 9
3             Select Case i
                  Case 6
4                     Text(i) = ">>,>>"
5                     Text(i).Mask = "99,99"
6                 Case 7
7                     Text(i) = ">>>>>>>>>"
8                     Text(i).Mask = "999999999"
9                 Case Else
10                    Text(i) = ""
11            End Select
12        Next
13        statusSBE = ""

14       On Error GoTo 0
15       Exit Sub

ber_sih_Error:

16        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure ber_sih of Form frmHOFlateRate")
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
1        On Error GoTo buka_Tutup_Error

2         Text(0).Enabled = Flg
3         Text(2).Enabled = Flg
4         If gDiv = "SVC" Then
5             Text(4).Enabled = Flg
6             Text(6).Enabled = flg1
7             Text(9).Enabled = Flg
8         ElseIf gDiv = "BP" Then
9             Text(7).Enabled = flg1
10            Text(8).Enabled = flg1
11            Text(9).Enabled = flg1
12        End If
13        Text(5).Enabled = flg1
14        Text(9).BackColor = &HC0E0FF

15       On Error GoTo 0
16       Exit Sub

buka_Tutup_Error:

17        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure buka_Tutup of Form frmHOFlateRate")
End Sub

Private Sub refreshGrid(idx As Integer)
      Dim nCol As Integer
      Dim rsRec As New ADODB.Recordset
              
          ' PAKAI RECORDSET
1        On Error GoTo refreshGrid_Error

2         If rsRec.State Then rsRec.Close
3         If gDiv = "SVC" Then
4             If idx = 1 Then
5                 Set rsRec = objO.FetchDataHO("SELECT fs_grp_mdl as [GROUP], mdl as KodeMdl, MODEL, a.JOB, '' as TIPE, fs_item as [ORDER], RATE, KETERANGAN, SBE FROM ho_flate a WITH (NOLOCK) INNER JOIN job b WITH (NOLOCK) ON a.job = b.job and A.FS_LOG_DEL = '' and B.FS_LOG_DEL = '' and MDL = '" & Text(10) & "' ORDER BY FS_GRP_MDL, MDL, a.JOB, FS_ITEM")
6             Else
7                 Set rsRec = objO.FetchDataHO("SELECT fs_grp_mdl as [GROUP], mdl as KodeMdl, MODEL, a.JOB, '' as TIPE, fs_item as [ORDER], RATE, KETERANGAN, SBE FROM ho_flate a WITH (NOLOCK) INNER JOIN job b WITH (NOLOCK) ON a.job = b.job and A.FS_LOG_DEL = '' and B.FS_LOG_DEL = ''  and MDL = '" & Text(10) & "'ORDER BY FS_GRP_MDL, MDL, a.JOB, FS_ITEM")
8             End If
9             nCol = 1
10        ElseIf gDiv = "BP" Then
11            Set rsRec = objO.FetchDataHO("SELECT fs_grp_mdl as [GROUP],mdl as KodeMdl, MODEL, a.JOB, tipeBP as TIPE, fs_item as [ORDER], HARGA, KETERANGAN, SBE FROM ho_flate a WITH (NOLOCK) INNER JOIN job b WITH (NOLOCK) ON a.job = b.job and A.FS_LOG_DEL = '' and B.FS_LOG_DEL = ''  and MDL = '" & Text(10) & "'ORDER BY FS_GRP_MDL, MDL, a.JOB, FS_ITEM")
12            nCol = 2
13        End If
14        If Not rsRec.EOF Then
15            Call createRS
16            Do While Not rsRec.EOF
17                rsGrid.AddNew
18                rsGrid!Group = Trim(rsRec("GROUP"))
19                rsGrid!kodeMDL = rsRec("KODEMDL")
20                rsGrid!MDL = rsRec("MODEL")
21                rsGrid!job = rsRec("JOB")
22                rsGrid!tipe = IIf(gDiv = "SVC", "", rsRec("TIPE"))
23                rsGrid!Order = rsRec("ORDER")
24                rsGrid!Rate = rsRec(6)
25                rsGrid!keterangan = rsRec("KETERANGAN")
26                rsGrid!SBE = rsRec("SBE")
27                rsGrid!Check = False
28                rsGrid.Update
29                rsRec.MoveNext
30            Loop
31        Else
32            Set rsGrid = Nothing
33        End If
34        GridFlate.DataSource = rsGrid
35        Call HeadColoms(nCol)

36       On Error GoTo 0
37       Exit Sub

refreshGrid_Error:

38        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form frmHOFlateRate")
End Sub

Private Sub createRS()
1        On Error GoTo createRS_Error

2         On Error GoTo vb_error
          
3         If rsGrid.State Then rsGrid.Close
4         Set rsGrid = Nothing
5         With rsGrid
6             .Fields.Append "GROUP", adVarChar, 20
7             .Fields.Append "KODEMDL", adVarChar, 25
8             .Fields.Append "MDL", adVarChar, 50
9             .Fields.Append "JOB", adVarChar, 5
10            .Fields.Append "TIPE", adVarChar, 30
11            .Fields.Append "ORDER", adVarChar, 50
12            .Fields.Append "RATE", adVarChar, 10
13            .Fields.Append "KETERANGAN", adVarChar, 30
14            .Fields.Append "SBE", adVarChar, 1
15            .Fields.Append "CHECK", adBoolean
16            .Open
17        End With
18        Set GridFlate.DataSource = rsGrid
19        Exit Sub
          
vb_error:
20       MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateRs"), vbCritical

21       On Error GoTo 0
22       Exit Sub

createRS_Error:

23        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRS of Form frmHOFlateRate")
End Sub

Private Sub HeadColoms(idx As Integer)
          
1        On Error GoTo HeadColoms_Error

2         GridFlate.Columns(0).Width = IIf(chPilih(0).value = Unchecked, 0, 1000)
3         GridFlate.Columns(0).DataField = "GROUP"
4         GridFlate.Columns(0).Caption = "GROUP MODEL"
5         GridFlate.Columns(1).Width = IIf(chPilih(1).value = Unchecked, 0, 2000)
6         GridFlate.Columns(1).DataField = "KODEMDL"
7         GridFlate.Columns(1).Caption = "KODE MODEL"
8         GridFlate.Columns(2).Width = IIf(chPilih(2).value = Unchecked, 0, 4000)
9         GridFlate.Columns(2).DataField = "MDL"
10        GridFlate.Columns(2).Caption = "MODEL"
11        GridFlate.Columns(3).Width = IIf(chPilih(3).value = Unchecked, 0, 800)
12        GridFlate.Columns(3).DataField = "JOB"
13        GridFlate.Columns(3).Caption = "TIPE JOB"
          
14        If idx = 1 Then
15            chPilih(7).value = Unchecked
16            GridFlate.Columns(4).Width = IIf(chPilih(7).value = Unchecked, 0, 2000)
17            GridFlate.Columns(4).Visible = IIf(chPilih(7).value = Unchecked, False, True)
18            GridFlate.Columns(5).Width = IIf(chPilih(4).value = Unchecked, 0, 6000)
19            GridFlate.Columns(6).Width = IIf(chPilih(5).value = Unchecked, 0, 800)
20            GridFlate.Columns(7).Width = IIf(chPilih(6).value = Unchecked, 0, 2000)
21        ElseIf idx = 2 Then
22            chPilih(7).value = Checked
23            GridFlate.Columns(4).Width = IIf(chPilih(7).value = Unchecked, 0, 2000)
24            GridFlate.Columns(5).Width = IIf(chPilih(4).value = Unchecked, 0, 5000)
25            GridFlate.Columns(6).Width = IIf(chPilih(5).value = Unchecked, 0, 2000)
26            GridFlate.Columns(7).Width = IIf(chPilih(6).value = Unchecked, 0, 2000)
27        End If
28        GridFlate.Columns(4).DataField = "TIPE"
29        GridFlate.Columns(4).Caption = "TIPE"
30        GridFlate.Columns(5).DataField = "ORDER"
31        GridFlate.Columns(5).Caption = "JOB ORDER"
32        GridFlate.Columns(6).DataField = "RATE"
33        GridFlate.Columns(6).Caption = "RATE"
34        GridFlate.Columns(7).DataField = "KETERANGAN"
35        GridFlate.Columns(7).Caption = "KETERANGAN"
36        GridFlate.Columns(8).Width = 800
37        GridFlate.Columns(8).Alignment = dbgCenter

38       On Error GoTo 0
39       Exit Sub

HeadColoms_Error:

40        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure HeadColoms of Form frmHOFlateRate")
End Sub

Private Sub chGroup_Click()
    If chGroup.value = Checked Then
        Text(9).Enabled = True
        Text(9).BackColor = &H80000005
        Text(0).Enabled = False
        Text(0).BackColor = &HC0E0FF
        Shape4.Visible = True
    ElseIf chGroup.value = Unchecked Then
        Text(9).Enabled = False
        Text(0).Enabled = True
        Text(0).BackColor = &H80000005
    End If
End Sub

Private Sub chPilih_Click(Index As Integer)
    If rsGrid.State Then
    Select Case Index
        Case 0
            GridFlate.Columns(Index).Visible = IIf(chPilih(Index) = Unchecked, False, True)
        Case 1
            GridFlate.Columns(Index).Visible = IIf(chPilih(Index) = Unchecked, False, True)
        Case 2
            GridFlate.Columns(Index).Visible = IIf(chPilih(Index) = Unchecked, False, True)
        Case 3
            GridFlate.Columns(Index).Visible = IIf(chPilih(Index) = Unchecked, False, True)
        Case 4
            GridFlate.Columns(5).Visible = IIf(chPilih(Index) = Unchecked, False, True)
        Case 5
            GridFlate.Columns(6).Visible = IIf(chPilih(Index) = Unchecked, False, True)
        Case 6
            GridFlate.Columns(7).Visible = IIf(chPilih(Index) = Unchecked, False, True)
        Case 7
            GridFlate.Columns(4).Visible = IIf(chPilih(Index) = Unchecked, False, True)
    End Select
    End If
End Sub

Private Sub cmdDisplay_Click()
    Call refreshGrid(2)
End Sub

Private Sub cmdHapus_Click()
Dim SQL As String
'    If gDiv = "SVC" Then
        Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_flate WITH (NOLOCK) WHERE mdl like '" & Text(0) & "-%' AND job = '" & Text(2) & "' AND fs_item = '" & Text(4) & "' and FS_LOG_DEL = ''")
        SQL = "UPDATE ho_flate SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE mdl like '" & Text(0) & "-%' AND job = '" & Text(2) & "' AND fs_item = '" & Text(4) & "' and FS_LOG_DEL = ''"
        SSQL2 = "INSERT INTO TEMPKIRIM VALUES ('FLATE', '" & Text(0) & "|" & Text(4) & "', 'D')"
'    ElseIf gDiv = "BP" Then
'        Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_flate WHERE FS_GRP_MDL = '" & Text(9) & "' AND mdl = '" & Text(0) & "' AND job = '" & Text(2) & "' AND fs_item = '" & Text(4) & "' AND tipeBP = '" & Text(8) & "' and FS_LOG_DEL = ''")
'        SQL = "UPDATE ho_flate SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE mdl = '" & Text(0) & "' AND job = '" & Text(2) & "' AND fs_item = '" & Text(4) & "' AND tipeBP = '" & Text(8) & "' and FS_LOG_DEL = ''"
'        SSQL2 = ""
'    End If
    
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin Akan menghapus Flate Rate " & Text(4) & " dari Model " & Text(1) & " ?", vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO(SQL)
            Call objO.UpdateDataHO(SSQL2)
            Call ber_sih
            Call refreshGrid(1)
        End If
    Else
        MsgBox "Pilih data yang akan dihapus"
        Exit Sub
    End If
    If rsTemp.State <> adStateClosed Then rsTemp.Close
End Sub

Private Sub cmdImport_Click()
          
1         On Error GoTo cmdImport_Click_Error
          Dim ErrBegin As Boolean, konter As Long
          Dim CnExel As New ADODB.Connection
          Dim RcExcel As New ADODB.Recordset
          Dim Jml As Integer
2         konter = 0
3         ProgBar.Min = 0
4         CommonDialog1.FileName = ""
5         CommonDialog1.InitDir = App.Path
6         CommonDialog1.filter = "Exel File (*.xls)|*.xls"
7         CommonDialog1.ShowOpen
8         If CommonDialog1.FileName <> "" Then
9             If CnExel.State Then CnExel.Close
10            CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
11            GoSub ProsesImport
12        Else
13            Exit Sub
14        End If
15        Exit Sub
          
ProsesImport:
16        ErrBegin = True
          Dim VarNmFile As String
          Dim job As String
          Dim namaModel As String
          Dim rsExs As ADODB.Recordset
17        Call objHT.OpenTransactionHO
          
18        If RcExcel.State Then RcExcel.Close
19        RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
20        If RcExcel.RecordCount > 0 Then
21            ProgBar.Visible = True
22            ProgBar.Max = RcExcel.RecordCount
23            RcExcel.MoveFirst
24            Do Until RcExcel.EOF
25                DoEvents
26                konter = konter + 1
27                ProgBar.value = konter
28                job = ""
29                namaModel = ""
                  
30                If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
31                If Nul(RcExcel.Fields(1).value) = "" Then
32                    Call objHT.CloseTransactionHO_err
33                    ProgBar.Visible = False
34                    MsgBox "JOB baris ke-" & konter & " KOSONG!", vbCritical, "Error"
35                    Exit Sub
36                End If
37                If IsNumeric(RcExcel.Fields(3).value) = False Then
38                    Call objHT.CloseTransactionHO_err
39                    ProgBar.Visible = False
40                    MsgBox "RATE baris ke-" & konter & " bukan angka!", vbCritical, "Error"
41                    Exit Sub
42                End If
43                sSQL = "SELECT * FROM model WITH (NOLOCK) WHERE left(mdl," & Len(Trim(RcExcel.Fields(0).value)) & ") = '" & Trim(RcExcel.Fields(0).value) & "' and FS_LOG_DEL = ''"
44                Set rsTemp = objO.FetchDataHO(sSQL)
45                If rsTemp.EOF Then
46                    Call objHT.CloseTransactionHO_err
47                    ProgBar.Visible = False
48                    MsgBox "MODEL " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " salah!", vbCritical, "Error"
49                    Exit Sub
50                Else
51                    namaModel = Trim(rsTemp("fs_grp_mdl"))
52                End If
                  
53                sSQL = "SELECT * FROM job WITH (NOLOCK) WHERE job = '" & Trim(RcExcel.Fields(1).value) & "' and FS_LOG_DEL = ''"
54                Set rsTemp = objO.FetchDataHO(sSQL)
                  
55                If rsTemp.EOF Then
56                    Call objHT.CloseTransactionHO_err
57                    ProgBar.Visible = False
58                    MsgBox "JOB " & Trim(RcExcel.Fields(1).value) & " baris ke-" & konter & " salah!", vbCritical, "Error"
59                    Exit Sub
60                Else
61                    job = Trim(rsTemp("remark"))
62                End If
63                Set rsTemp = objO.FetchDataHO("SELECT * FROM MOdel WITH (NOLOCK) WHERE left(mdl," & Len(Trim(RcExcel.Fields(0).value)) & ") = '" & Trim(RcExcel.Fields(0).value) & "' and FS_LOG_DEL = ''")
64                If rsTemp.RecordCount > 0 Then
65                rsTemp.MoveFirst
66                Do Until rsTemp.EOF
67                    Set rsExs = objO.FetchDataHO("Select * from Ho_flate WITH (NOLOCK) where MDL = '" & Trim(rsTemp("MDL")) & "' and fs_Item = '" & Trim(RcExcel.Fields(2).value) & "' and job = '" & Trim(RcExcel.Fields(1).value) & "' and FS_LOG_DEL = ''")
68                    If rsExs.RecordCount = 0 Then
69                        Call objHT.InsertData_Ho("INSERT INTO ho_flate (FS_GRP_MDL,MDL,MODEL,JOB,REMARK,FS_ITEM,Rate,KETERANGAN,TIPEBP,HARGA,FS_LOG_DEL,FD_TGL_DEL,SBE) VALUES ('" & Trim(rsTemp("FS_GRP_MDL")) & "','" & Trim(rsTemp("MDL")) & "','" & Trim(rsTemp("DESC")) & "','" & Trim(RcExcel.Fields(1).value) & "','" & job & "','" & Trim(RcExcel.Fields(2).value) & "','" & Trim(Replace(Replace(Replace(Replace(Trim(RcExcel.Fields(3).value), ">", ""), " ", ""), ",", "."), " ", "")) & "','" & Trim(RcExcel.Fields(4).value) & "','','0','','', '" & Trim(statusSBE) & "')")
70                    Else
71                        Call objHT.InsertData_Ho("UPDATE ho_flate SET rate = '" & Replace(Replace(Replace(Trim(RcExcel.Fields(3).value), ">", ""), " ", ""), ",", ".") & "', keterangan = '" & Trim(RcExcel.Fields(4).value) & "' WHERE mdl = '" & Trim(rsTemp("MDL")) & "' AND job = '" & Trim(RcExcel.Fields(1).value) & "' AND fs_item = '" & Trim(RcExcel.Fields(2).value) & "' and FS_LOG_DEL = ''")
72                    End If
73                    rsTemp.MoveNext
74                Loop
75                SSQL2 = "INSERT INTO TEMPKIRIM VALUES ('FLATE', '" & Trim(RcExcel.Fields(0).value) & "|" & Trim(RcExcel.Fields(2).value) & "', 'S')"
76                End If
77                RcExcel.MoveNext
78            Loop
79        End If
80        RcExcel.Close
81        CnExel.Close
          '   ProgBar.Visible = False
          '  ' IsiGrid
          '   DbCon.CommitTrans
          '   ErrBegin = False
82        Call objHT.CloseTransactionHO
83        MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
84        Call refreshGrid(1)
85        ProgBar.Visible = False
86        On Error GoTo 0
87        Exit Sub

cmdImport_Click_Error:
88    ProgBar.Visible = False
89        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdImport_Click of Form frmHOFlateRate")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
1        On Error GoTo cmdSimpan_Click_Error

2         If cmdSimpan.Caption = "&Tambah" Then
3             Call ber_sih
4             Call buka_Tutup(True, True)
5             cmdUbah.Caption = "&Batal"
6             cmdSimpan.Caption = "&Simpan"
7         ElseIf cmdSimpan.Caption = "&Simpan" Then
8             boleh = True
9             pesan = "Lengkapi data "
              
10            If gDiv = "SVC" Then
11                If Text(0) = "" Then
12                    pesan = pesan + "Model. "
13                    boleh = False
14                End If
15                If Text(2) = "" Then
16                    pesan = pesan + "Tipe Job. "
17                    boleh = False
18                End If
19                If Text(4) = "" Then
20                    pesan = pesan + "Job Order. "
21                    boleh = False
22                End If
23                If Text(6) = ">>,>>" Then
24                    pesan = pesan + "Rate. "
25                    boleh = False
26                End If
27            ElseIf gDiv = "BP" Then
28                If Text(9) = "" Then
29                    pesan = pesan + "Group Model. "
30                    boleh = False
31                End If
32                If Text(2) = "" Then
33                    pesan = pesan + "Tipe Job. "
34                    boleh = False
35                End If
36                If Text(4) = "" Then
37                    pesan = pesan + "Job Order. "
38                    boleh = False
39                End If
40                If Text(8) = ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" Then
41                    pesan = pesan + "Tipe BP. "
42                    boleh = False
43                End If
44                If Text(7) = ">>>>>>>>>" Then
45                    pesan = pesan + "Harga. "
46                    boleh = False
47                End If
48            End If
          
49            If boleh Then
                  Dim rsExs As ADODB.Recordset
50                Set rsTemp = objO.FetchDataHO("SELECT * FROM MOdel WITH (NOLOCK) WHERE left(mdl," & Len(Text(0).Text) & ") = '" & Text(0) & "' and FS_LOG_DEL = ''")
51                If rsTemp.RecordCount > 0 Then
52                    rsTemp.MoveFirst
53                    Do Until rsTemp.EOF
54                        Set rsExs = objO.FetchDataHO("Select * from Ho_flate WITH (NOLOCK) where MDL = '" & Trim(rsTemp("MDL")) & "' and fs_Item = '" & Text(4) & "' and job = '" & Text(2) & "' and FS_LOG_DEL = ''")
55                        If rsExs.RecordCount = 0 Then
56                            Call objO.UpdateDataHO("INSERT INTO ho_flate (FS_GRP_MDL,MDL,MODEL,JOB,REMARK,FS_ITEM,Rate,KETERANGAN,TIPEBP,HARGA,FS_LOG_DEL,FD_TGL_DEL,SBE) VALUES ('" & Trim(rsTemp("FS_GRP_MDL")) & "','" & Trim(rsTemp("MDL")) & "','" & Trim(rsTemp("DESC")) & "','" & Text(2) & "','" & Text(3) & "','" & Text(4) & "','" & Trim(Replace(Replace(Replace(Replace(Text(6), ">", ""), " ", ""), ",", "."), " ", "")) & "','" & Text(5) & "','','0','','', '" & Trim(statusSBE) & "')")
57                        Else
58                            Call objO.UpdateDataHO("UPDATE ho_flate SET rate = '" & Replace(Replace(Replace(Text(6), ">", ""), " ", ""), ",", ".") & "', keterangan = '" & Text(5) & "' WHERE mdl = '" & Trim(rsTemp("MDL")) & "' AND job = '" & Text(2) & "' AND fs_item = '" & Text(4) & "' and FS_LOG_DEL = ''")
59                        End If
                          
60                        Set rsExs = objO.FetchDataHO("Select * from ACTRATE WITH (NOLOCK) where MDL = '" & Trim(rsTemp("MDL")) & "' and fs_Item = '" & Text(4) & "' and job = '" & Text(2) & "' and FS_LOG_DEL = ''")
61                        If rsExs.EOF Then
62                            Call objO.UpdateDataHO("INSERT INTO ACTRATE (MDL,MODEL,JOB,REMARK,FS_ITEM,LEVELMECH,KETLEV,Rate,KETERANGAN,OTHER,SBE,FS_LOG_DEL,FD_TGL_DEL) VALUES ( '" & rsTemp("MDL") & "', '" & rsTemp("MODEL") & "', '" & Trim(Text(2)) & "','" & Text(3) & "', '" & Text(4) & "', '','','" & Replace((CDbl(Replace(Trim(Text(6)), ">", "")) * 0.8), ",", ".") & "', '" & Text(5) & "', 0, '" & Trim(statusSBE) & "','" & Trim(rsTemp("FS_LOG_DEL")) & "', '" & rsTemp("FD_TGL_DEL") & "')")
63                        Else
64                            Call objO.UpdateDataHO("UPDATE ACTRATE SET MODEL = '" & rsTemp("MODEL") & "', JOB = '" & Trim(Text(2)) & "', REMARK = '" & Text(2) & "', RATE = " & Replace((CDbl(Trim(Text(6))) * 0.8), ",", ".") & ", SBE = '" & Trim(statusSBE) & "',KETERANGAN = '" & Text(5) & "', FS_LOG_DEL = '" & Trim(rsTemp("FS_LOG_DEL")) & "', FD_TGL_DEL = '" & Format(rsTemp("FD_TGL_DEL"), "MM/DD/YYYY HH:NN:SS") & "', Other = 0 WHERE FS_ITEM = '" & Text(4) & "' and MDL = '" & Trim(rsTemp("MDL")) & "' and FS_LOG_DEL = ''")
65                        End If
                                                                        
66                        rsTemp.MoveNext
67                    Loop
68                    SSQL2 = "INSERT INTO TEMPKIRIM VALUES ('FLATE', '" & Text(0) & "|" & Text(4) & "', 'S')"
69                End If

70                objO.UpdateDataHO (SSQL2)
71                If rsTemp.State <> adStateClosed Then rsTemp.Close
                  
72                Call ber_sih
73                Call buka_Tutup(False, False)
74                Call refreshGrid(1)
75                cmdUbah.Caption = "&Ubah"
76                cmdSimpan.Caption = "&Tambah"
77            Else
78                MsgBox pesan, vbCritical, "KELENGKAPAN"
79            End If
80        End If

81       On Error GoTo 0
82       Exit Sub

cmdSimpan_Click_Error:

83        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmHOFlateRate")
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        newRecord = False
        Call buka_Tutup(False, True)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
    End If
End Sub

Private Sub Form_Load()
    If gDiv = "SVC" Then
        Label4.Caption = "Rate"
        chGroup.Visible = False
        Shape5.Visible = False
        Text(8).Visible = False
        Label5.Visible = False
        Text(7).Visible = False
        chPilih(7).Visible = False
    ElseIf gDiv = "BP" Then
        Label4.Caption = "Harga"
        chGroup.Visible = True
        Shape5.Visible = True
        Text(8).Visible = True
        Label5.Visible = True
        Text(7).Left = Text(6).Left
        Text(7).Top = Text(6).Top
        Text(7).Visible = True
        chPilih(7).Visible = True
    End If
    Text(9).Enabled = False
    For i = 0 To 6
        chPilih(i).value = Checked
    Next
    Call ber_sih
    Call buka_Tutup(False, False)
    Call LoadModel
    Call refreshGrid(1)
End Sub

Private Sub LoadModel()
1        On Error GoTo LoadModel_Error

2         Text(10) = objO.FetchData_String("Select top 1 MDL from Model WITH (NOLOCK)")

3        On Error GoTo 0
4        Exit Sub

LoadModel_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure LoadModel of Form frmHOFlateRate")
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridFlate.Width = Frame1.Width - 200
    GridFlate.Height = Frame1.Height - Text(5).Top - Text(5).Height - 900
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdImport.Left = cmdSimpan.Left - cmdImport.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    cmdDisplay.Top = cmdSimpan.Top
End Sub

Private Sub GridFlate_FilterChange()
1        On Error GoTo GridFlate_FilterChange_Error

2         Set cols = GridFlate.Columns
          Dim C As Integer
          
3         C = GridFlate.col
4         GridFlate.HoldFields
5         rsGrid.filter = getFilter()
6         GridFlate.col = C
7         GridFlate.EditActive = True

8        On Error GoTo 0
9        Exit Sub

GridFlate_FilterChange_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GridFlate_FilterChange of Form frmHOFlateRate")
End Sub

Private Function getFilter() As String
      Dim tmp As String
      Dim n As Integer
1        On Error GoTo getFilter_Error

2         For Each col In cols
3             If Trim(col.FilterText) <> "" Then
4                 n = n + 1
5                 If n > 1 Then
6                     tmp = tmp & " AND "
7                 End If
8                 tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
9             End If
10        Next col
11        getFilter = tmp

12       On Error GoTo 0
13       Exit Function

getFilter_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form frmHOFlateRate")
End Function

Private Sub GridFlate_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsGrid.State <> adStateClosed Then
        If Not rsGrid.EOF Then
            On Error Resume Next
            Text(0) = Left(rsGrid.Fields(1), (InStr(1, rsGrid.Fields(1), "-") - 1)) '("[KODE MODEL]")
            Text(1) = rsGrid.Fields(2)  '("MODEL")
            Text(2) = rsGrid("JOB")
            Text(4) = rsGrid.Fields(5) '("JOB ORDER")
            If gDiv = "SVC" Then
                Text(6) = Format(Replace(IIf(rsGrid("RATE") = "", 0, rsGrid("RATE")), ".", ","), "00.00")
            ElseIf gDiv = "BP" Then
                Text(7) = rsGrid("HARGA")
            End If
            Text(5) = rsGrid("KETERANGAN")
            Text(9) = rsGrid.Fields(0)  '("GROUP MODEL")
            Text(8) = rsGrid("TIPE")
            statusSBE = rsGrid("SBE")
        End If
    End If
End Sub

Private Sub opBP_DblClick(Index As Integer)
    Text(8).Mask = ""
    Text(8) = Replace(opBP(Index).Caption, "&&", "&")
    Frame2.Visible = False
End Sub

Private Sub opBP_GotFocus(Index As Integer)
    opBP(Index).value = True
End Sub

Private Sub opBP_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text(8).Mask = ""
        Text(8) = Replace(opBP(Index).Caption, "&&", "&")
        Frame2.Visible = False
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
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "GRPMDL"
                frmCariHO.show (vbModal)
            Case 10
                gMainKeyCari = Text(Index)
                gSearchOption = "GRPMDL"
                frmCariHO.show (vbModal)
            Case 2
                If Text(0) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCariHO.show (vbModal)
                Else
                    MsgBox "Isi Model Kendaraan", vbCritical, "KELENGKAPAN"
                    Text(0).SetFocus
                    Exit Sub
                End If
            Case 4
                If Text(2) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = Text(2)
                    frmCariHO.show (vbModal)
                Else
                    MsgBox "Isi Tipe Job Order", vbCritical, "KELENGKAPAN"
                    Text(2).SetFocus
                    Exit Sub
                End If
            Case 8
                Frame2.Visible = True
                opBP(0).SetFocus
            Case 9
                gMainKeyCari = Text(Index)
                gSearchOption = "GROUP_MODEL"
                frmCariHO.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As New ADODB.Recordset

           
1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 0
3                 If Text(Index) <> "" Then
4                     sSQL = "SELECT * FROM model WITH (NOLOCK) WHERE left(mdl," & Len(Text(Index).Text) & ") = '" & Text(Index) & "' and FS_LOG_DEL = ''"
5                     Set rsSearch = objO.FetchDataHO(sSQL)
6                     If rsSearch.BOF Then
7                         Cancel = True
8                         gMainKeyCari = Text(Index)
9                         gSearchOption = "GRPMDL"
10                        frmCariHO.show (vbModal)
11                    Else
12                        Text(1) = rsSearch("DESC")
13                        Text(9) = rsSearch("FS_GRP_MDL")
14                    End If
15                End If
16            Case 10
17                If Text(Index) <> "" Then
18                    Text(10) = objO.FetchData_String("Select top 1 MDL from Model WITH (NOLOCK) Where MDL like '" & Text(10) & "-%'")
19                    Call refreshGrid(1)
20                End If
21            Case 2
22                If Text(Index) <> "" Then
23                    If Text(0) <> "" Then
24                        sSQL = "SELECT * FROM job WITH (NOLOCK) WHERE job = '" & Text(Index) & "' and FS_LOG_DEL = ''"
25                        Set rsSearch = objO.FetchDataHO(sSQL)
26                        If rsSearch.BOF Then
27                            Cancel = True
28                            gMainKeyCari = Text(Index)
29                            gSearchOption = "JENIS_KERJA"
30                            frmCariHO.show (vbModal)
31                        Else
32                            Text(3) = rsSearch("REMARK")
33                        End If
34                    Else
35                        MsgBox "Isi Model Kendaraan", vbCritical, "KELENGKAPAN"
36                        Text(0).SetFocus
37                        Exit Sub
38                    End If
39                End If
40            Case 4
41                If Text(Index) <> "" Then
42                    If Text(2) <> "" Then
43                        sSQL = "SELECT * FROM jobOrder WITH (NOLOCK) WHERE FS_Item = '" & Text(Index) & "' and FS_LOG_DEL = ''"
44                        Set rsSearch = objO.FetchDataHO(sSQL)
45                        If rsSearch.BOF Then
46                            Cancel = True
47                            gMainKeyCari = Text(Index)
48                            gSearchOption = "JOB_ORDER"
49                            gBantuan = Text(2)
50                            frmCariHO.show (vbModal)
51                        Else
52                            Text(5) = rsSearch("KETERANGAN")
53                            statusSBE = Trim(rsSearch("SBE"))
54                        End If
55                    Else
56                        MsgBox "Isi Tipe Job Order", vbCritical, "KELENGKAPAN"
57                        Text(2).SetFocus
58                        Exit Sub
59                    End If
60                End If
61            Case 8
62                If Text(Index) = "" Or Text(Index) = ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" Then
63                    Frame2.Visible = True
64                    opBP(0).SetFocus
65                End If
66        End Select

67       On Error GoTo 0
68       Exit Sub

text_Validate_Error:

69        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmHOFlateRate")
           
End Sub
