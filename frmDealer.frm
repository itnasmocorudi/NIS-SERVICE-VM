VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmDealer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Master Kode Dealer"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   14535
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7245
      Left            =   435
      TabIndex        =   7
      Top             =   480
      Width           =   13980
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1755
         MaxLength       =   99
         TabIndex        =   3
         Top             =   1162
         Width           =   5340
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   1755
         TabIndex        =   8
         Top             =   585
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Index           =   0
         Left            =   1755
         TabIndex        =   0
         Top             =   270
         Width           =   3180
         _ExtentX        =   5609
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   5040
         TabIndex        =   9
         Top             =   270
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   3600
         TabIndex        =   10
         Top             =   2100
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   1755
         TabIndex        =   2
         Top             =   885
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Height          =   270
         Index           =   4
         Left            =   1755
         TabIndex        =   11
         Top             =   1485
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   5
         Left            =   1755
         TabIndex        =   12
         Top             =   1800
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
      Begin TrueOleDBGrid70.TDBGrid gridDealer 
         Height          =   4455
         Left            =   105
         TabIndex        =   4
         Top             =   2625
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   7858
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
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).DataField=   ""
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
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
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
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1755
         TabIndex        =   1
         Top             =   2100
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   3400
         TabIndex        =   18
         Top             =   2107
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   8
         Left            =   1755
         TabIndex        =   20
         Top             =   1162
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   195
         Left            =   1140
         TabIndex        =   21
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pelanggan"
         Height          =   195
         Left            =   390
         TabIndex        =   19
         Top             =   2175
         Width           =   1185
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Dealer"
         Height          =   195
         Left            =   735
         TabIndex        =   17
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Left            =   1200
         TabIndex        =   16
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penagihan"
         Height          =   195
         Left            =   390
         TabIndex        =   15
         Top             =   915
         Width           =   1230
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   195
         Left            =   1245
         TabIndex        =   14
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   1845
         Width           =   495
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   120
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   12045
      TabIndex        =   6
      Top             =   7830
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   10800
      TabIndex        =   5
      Top             =   7830
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8055
      Left            =   0
      Picture         =   "frmDealer.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmDealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsGrid As New ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns

Private Sub bersih()
    If Log.RPLHistori = 1 Then
        Text(0).BackColor = &HC0E0FF
        Text(2).BackColor = &HC0E0FF
        Text3.BackColor = &HC0E0FF
        Text3.Locked = True
    End If
    For i = 0 To 7
        Select Case i
            Case 3
            Case Else
                Text(i) = ""
        End Select
    Next
End Sub

Private Sub headGrid()
    With GridDealer
        .Columns(0).Width = 1700
        .Columns(0).Caption = "Kode Dealer"
        .Columns(0).DataField = "FS_KD_DEALER"
        .Columns(0).Locked = True
        .Columns(1).Width = 2500
        .Columns(1).Caption = "Nama Dealer"
        .Columns(1).DataField = "FS_NM_DEALER"
        .Columns(1).Locked = True
        .Columns(2).Width = 3000
        .Columns(2).Caption = "Alamat Dealer"
        .Columns(2).DataField = "FS_ALAMAT"
        .Columns(2).Locked = True
        .Columns(3).Width = 1500
        .Columns(3).Caption = "Kota Dealer"
        .Columns(3).DataField = "FS_KOTA"
        .Columns(3).Locked = True
        .Columns(4).Width = 2000
        .Columns(4).Caption = "NPWP"
        .Columns(4).DataField = "FS_NPWP"
        .Columns(4).Locked = True
        .Columns(5).Width = 1000
        .Columns(5).Caption = "Kode Pelanggan"
        .Columns(5).DataField = "KDPLG"
        .Columns(5).Locked = True
        .Columns(6).Width = 2700
        .Columns(6).Caption = "Nama Penagihan"
        .Columns(6).DataField = "FS_NM_TAGIH"
        .Columns(6).Locked = True
    End With
End Sub

Private Sub refreshGrid()
          
1         If rsGrid.State Then rsGrid.Close
2         If Log.RPLHistori = 0 Then
3             Set rsGrid = objO.FetchData("SELECT FS_KD_DEALER,FS_NM_DEALER,FS_ALAMAT,FS_KOTA,FS_NPWP,KDPLG,FS_NM_TAGIH FROM T_DEALER WITH(NOLOCK) WHERE FS_LOG_DEL = ''")
4         Else
              ''Set rsGrid = objO.FetchData("SELECT FS_KD_DEALER,FS_NM_DEALER,FS_ALAMAT,FS_KOTA,FS_NPWP,KDPLG,FS_NM_TAGIH FROM [" & dbReplikasiHO & "]..T_DEALER WITH(NOLOCK) WHERE FS_LOG_DEL = ''")
              '' ==================================================================================
              '' 20181011 | REV | MB THANIA || '' biar ga bingung mana yg belum di link
              '' ==================================================================================
5             Set rsGrid = objO.FetchData("SELECT a.FS_KD_DEALER,a.FS_NM_DEALER,a.FS_ALAMAT,a.FS_KOTA,a.FS_NPWP,isnull(b.KDPLG,'') as KDPLG,a.FS_NM_TAGIH FROM [" & dbReplikasiHO & "]..T_DEALER a WITH(NOLOCK) LEFT JOIN T_DEALER B WITH (NOLOCK) ON a.FS_KD_DEALER = b.FS_KD_DEALER  WHERE a.FS_LOG_DEL = ''")
6         End If
7         If Not rsGrid.EOF Then
8             Set GridDealer.DataSource = rsGrid
9             Call headGrid
10        End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
          
1         If Trim(Text(0)) <> "" Then
2             If Trim(Text(7)) = "" Then
3                 If MsgBox("Kode Pelanggan dealer akan dikosongkan ?", vbDefaultButton2 + vbYesNo, "Kode Kosong") = vbYes Then
4                     GoTo lanjut
5                 Else
6                     Exit Sub
7                 End If
8             End If
              
9             If Trim(Text(2)) = "" Then
10                If MsgBox("Nama Penagihan dealer akan dikosongkan ?", vbDefaultButton2 + vbYesNo, "Nama Kosong") = vbYes Then
11                    GoTo lanjut
12                Else
13                    Exit Sub
14                End If
15            End If
lanjut:
16            If Log.RPLHistori = 0 Then
17                Call objO.UpdateData("UPDATE T_DEALER SET KDPLG = '" & Trim(Text(7)) & "',FS_NM_TAGIH = '" & Trim(Text(2)) & "', FS_NPWP = '" & Trim(Text(5)) & "', FS_ALAMAT = '" & Trim(Text3) & "', FS_KOTA = '" & Trim(Text(4)) & "'  WHERE FS_KD_DEALER = '" & Trim(Text(0)) & "'")
18            Else
19                Set rsTemp = objO.FetchData("SELECT * FROM T_DEALER WITH (NOLOCK) WHERE FS_KD_DEALER = '" & Trim(Text(0)) & "'")
20                If Not rsTemp.EOF Then
21                    Call objO.UpdateData("UPDATE T_DEALER SET KDPLG = '" & Trim(Text(7)) & "',FS_NM_TAGIH = '" & Trim(Text(2)) & "', FS_NPWP = '" & Trim(Text(5)) & "', FS_ALAMAT = '" & Trim(Text3) & "', FS_KOTA = '" & Trim(Text(4)) & "'  WHERE FS_KD_DEALER = '" & Trim(Text(0)) & "'")
22                Else
23                    Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_DEALER WITH (NOLOCK) WHERE FS_KD_DEALER = '" & Trim(Text(0)) & "'")
24                    If Not rsTemp.EOF Then
25                        Call objO.InsertData("INSERT INTO T_DEALER (FS_KD_DEALER, FS_NM_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, FS_PKP, KODECABANG, VPNCABANG, userCbg, passCbg, INTERNAL, KDPLG, FS_LOG_DEL, FD_TGL_DEL, FS_NM_TAGIH, HARGANRM) VALUES ('" & rsTemp("FS_KD_DEALER") & "', '" & rsTemp("FS_NM_DEALER") & "', '" & rsTemp("FS_ALAMAT") & "', '" & rsTemp("FS_KOTA") & "', '" & rsTemp("FS_NPWP") & "', '" & rsTemp("FS_PKP") & "', " _
                                                             & "'" & rsTemp("KODECABANG") & "', '" & rsTemp("VPNCABANG") & "', '" & rsTemp("USERCBG") & "', '" & rsTemp("PASSCBG") & "','" & CDbl(rsTemp("INTERNAL")) & "', '" & Trim(rsTemp("KDPLG")) & "','','','" & Trim(rsTemp("FS_NM_TAGIH")) & "','" & Trim(rsTemp("harganrm")) & "')")
26                        Call objO.UpdateData("UPDATE T_DEALER SET KDPLG = '" & Trim(Text(7)) & "',FS_NM_TAGIH = '" & Trim(Text(2)) & "', FS_NPWP = '" & Trim(Text(5)) & "', FS_ALAMAT = '" & Trim(Text3) & "', FS_KOTA = '" & Trim(Text(4)) & "'  WHERE FS_KD_DEALER = '" & Trim(Text(0)) & "'")
                          
27                    Else
28                        MsgBox "Gagal mengambil data.", vbInformation + vbOKOnly
29                        Call refreshGrid
30                        Exit Sub
31                    End If
32                End If
                  
33            End If
              
34            MsgBox "Kode Pelanggan Dealer berhasil di Update .", vbInformation + vbOKOnly
35            Call refreshGrid
36        Else
37            MsgBox "Dealer Tidak bisa menambahkan master Dealer Baru." & vbCr & "Hubungi HO Service", vbInformation + vbOKOnly
38        End If
End Sub

Private Sub Form_Load()
    Call bersih
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
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

Private Sub GridDealer_FilterChange()
    On Error GoTo errHandler
    Set cols = GridDealer.Columns
    Dim C As Integer
    
    C = GridDealer.col
    GridDealer.HoldFields
    rsGrid.filter = getFilter()
    If rsGrid.EOF Then
        For Each col In GridDealer.Columns
            col.FilterText = ""
        Next col
        rsGrid.filter = ""
    End If
    GridDealer.col = C
    GridDealer.EditActive = True
    
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub GridDealer_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsSearch As ADODB.Recordset
Dim isReplikasi As Boolean
    isReplikasi = False
    If Log.RPLHistori = 0 Then
        Set rsSearch = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Trim(GridDealer.Columns(0)) & "'")
    Else
    
        Set rsSearch = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Trim(GridDealer.Columns(0)) & "'")
        Set rsSearch = objO.FetchData("SELECT a.FS_KD_DEALER,a.FS_NM_DEALER,a.FS_ALAMAT,a.FS_KOTA,a.FS_NPWP,isnull(b.KDPLG,'') as KDPLG,a.FS_NM_TAGIH FROM [" & dbReplikasiHO & "]..T_DEALER a WITH(NOLOCK) LEFT JOIN T_DEALER B WITH (NOLOCK) ON a.FS_KD_DEALER = b.FS_KD_DEALER  WHERE a.FS_KD_DEALER = '" & Trim(GridDealer.Columns(0)) & "' and a.FS_LOG_DEL = ''")
        If rsSearch.EOF Then
            isReplikasi = True
            Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_DEALER a WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Trim(GridDealer.Columns(0)) & "'")
        End If
        
    
    End If
        If Not rsSearch.EOF Then
        Text(0) = Trim(rsSearch("FS_KD_DEALER"))
        Text(1) = Trim(rsSearch("FS_NM_DEALER"))
        Text(2) = Trim(rsSearch("FS_NM_TAGIH"))
        Text3 = Trim(rsSearch("FS_ALAMAT"))
        Text(4) = Trim(rsSearch("FS_KOTA"))
        Text(5) = Trim(rsSearch("FS_NPWP"))
        If isReplikasi = False Then
            Text(7) = Trim(rsSearch("KDPLG"))
            Text(6) = Left(Text(7), 3)
        Else '
            Text(7) = ""
            Text(6) = ""
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
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.show (vbModal)
            Case 7
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLG"
                'gSearchOption = "KODEPLGSID"
                frmCariAcess.show (vbModal)
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
      Dim rsSearch As ADODB.Recordset
      Dim rsHelp As ADODB.Recordset

1         Select Case Index
              Case 0
2                 If Text(Index) <> "" Then
3                     Set rsSearch = objO.FetchData("select * from t_dealer WITH(NOLOCK) where fs_kd_dealer = '" & Text(Index) & "' and FS_LOG_DEL = ''")
4                     If rsSearch.BOF Then
5                         Cancel = True
6                         gMainKeyCari = Text(Index)
7                         gBantuan = "ALL"
8                         gSearchOption = "DEALER"
9                         frmCari.show (vbModal)
10                    Else
11                        Text(1) = Trim(rsSearch("FS_NM_DEALER"))
12                        Text(2) = Trim(rsSearch("FS_NM_TAGIH"))
13                        Text3 = Trim(rsSearch("FS_ALAMAT"))
14                        Text(4) = Trim(rsSearch("FS_KOTA"))
15                        Text(5) = Trim(rsSearch("FS_NPWP"))
16                        Text(7) = Trim(rsSearch("KDPLG"))
17                        Text(6) = Trim(Left(Text(7), 3))
18                    End If
19                End If
20            Case 7
21                If Text(Index) <> "" Then
22                        Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WITH(NOLOCK) WHERE kodeLgn = '" & Text(Index) & "'")
23                        If rsSearch.EOF Then
24                            gMainKeyCari = Text(Index)
25                            gSearchOption = "KODEPLG"
26                            frmCariAcess.show (vbModal)
27                        Else
28                            Text(6) = Nul(rsSearch("KODEWIL"))
29                            Text3 = Left(Nul(rsSearch("ALAMAT1")), 30)
30                            Text(4) = Left(Nul(rsSearch("ALAMAT2")), 30)
31                            Text(5) = Nul(rsSearch("NPWP"))
32                        End If
                      
33                End If
34        End Select
End Sub

