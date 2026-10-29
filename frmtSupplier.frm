VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtSupplier 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Master Supplier"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   17430
   ScaleWidth      =   23760
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmtSupplier.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   -120
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   480
      TabIndex        =   10
      Top             =   0
      Width           =   12780
      Begin VB.TextBox txtPekerjaan 
         Height          =   1455
         Left            =   6960
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   5620
      End
      Begin VB.CheckBox chJurnal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Non Jurnal"
         Height          =   195
         Left            =   11040
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chPKP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PKP"
         Height          =   195
         Left            =   9960
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   1395
         TabIndex        =   1
         Top             =   585
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Left            =   1395
         TabIndex        =   0
         Top             =   270
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   3045
         TabIndex        =   11
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
         Index           =   7
         Left            =   3240
         TabIndex        =   14
         Top             =   270
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
         Left            =   1395
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
         Height          =   270
         Index           =   3
         Left            =   1395
         TabIndex        =   3
         Top             =   1185
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Height          =   270
         Index           =   4
         Left            =   1395
         TabIndex        =   4
         Top             =   1485
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Height          =   270
         Index           =   5
         Left            =   1395
         TabIndex        =   5
         Top             =   1800
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Height          =   270
         Index           =   6
         Left            =   1395
         TabIndex        =   6
         Top             =   2100
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin TrueOleDBGrid70.TDBGrid gridSupplier 
         Height          =   4575
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8070
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
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD7FFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HD5E1F7&"
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
         Index           =   8
         Left            =   4875
         TabIndex        =   21
         Top             =   270
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   7095
         Left            =   30
         Top             =   120
         Width           =   12700
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pekerjaan"
         Height          =   195
         Left            =   6960
         TabIndex        =   19
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telepon"
         Height          =   195
         Left            =   630
         TabIndex        =   17
         Top             =   1845
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   195
         Left            =   930
         TabIndex        =   16
         Top             =   1515
         Width           =   330
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   195
         Left            =   780
         TabIndex        =   15
         Top             =   915
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Supplier"
         Height          =   255
         Left            =   195
         TabIndex        =   12
         Top             =   285
         Width           =   1065
      End
   End
   Begin ucXPButton.XPButton cmdHapus 
      Height          =   495
      Left            =   10440
      TabIndex        =   23
      Top             =   7320
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
      MICON           =   "frmtSupplier.frx":217E
      PICN            =   "frmtSupplier.frx":219A
      PICH            =   "frmtSupplier.frx":2722
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
      Left            =   9240
      TabIndex        =   24
      Top             =   7320
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
      MICON           =   "frmtSupplier.frx":2D5E
      PICN            =   "frmtSupplier.frx":2D7A
      PICH            =   "frmtSupplier.frx":328D
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
      Left            =   12000
      TabIndex        =   25
      Top             =   7320
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
      MICON           =   "frmtSupplier.frx":382F
      PICN            =   "frmtSupplier.frx":384B
      PICH            =   "frmtSupplier.frx":3D1F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   13200
      X2              =   360
      Y1              =   7920
      Y2              =   7920
   End
End
Attribute VB_Name = "frmtSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsGrid As ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
Dim rsSearch As ADODB.Recordset

    If Trim(text(8)) <> "" Then
        Set rsSearch = objO.FetchData("SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE SUPPLIER = '" & Trim(text(8)) & "'")
        If Not rsSearch.EOF Then
            MsgBox "Kode supplier sudah digunakan dalam Order !!", vbCritical, "Supplier tidak bisa dihapus"
        Else
            Call objO.UpdateData("UPDATE T_SUPPLIER SET FS_LOG_DEL = '" & Trim(gSA) & "', FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "' WHERE KDSUPP = '" & Trim(text(8)) & "'")
            Call bersih
            Call refreshGrid
        End If
    End If
End Sub

Private Sub cmdSimpan_Click()
Dim boleh As Boolean
Dim pesan As String
Dim rsSimpan As ADODB.Recordset
    
    On Error GoTo vbError
    boleh = True
    pesan = ""
    If Trim(text(0)) = "" Then
        boleh = False
        pesan = "Kode Supplier. "
    End If
    If Trim(text(1)) = "" Then
        boleh = False
        pesan = "Nama Supplier. "
    End If
    If Trim(txtPekerjaan) = "" Then
        boleh = False
        pesan = "Pekerjaan Sublet. "
    End If
    If chPKP = Unchecked Then
        If MsgBox("Apakah Supplier Non PKP ?", vbYesNo, "PKP / NON PKP") = vbNo Then
            boleh = False
            pesan = "Status PKP. "
        End If
    End If
    If chJurnal = Checked Then
        If MsgBox("Apakah transaksi sublet dengan supplier tsb tidak dijurnal ?", vbYesNo, "Jurnal / Non Jurnal") = vbNo Then
            boleh = False
            pesan = "Status Jurnal. "
        End If
    End If
    
    If boleh Then
               
        If Trim(text(8)) = "" Then
            Dim kdSPP As String
            Set rsSimpan = objO.FetchData("SELECT * FROM T_SUPPLIER WITH (NOLOCK) ORDER BY KDSUPP DESC")
            If Not rsSimpan.EOF Then
                kdSPP = Trim(rsSimpan("KDSUPP")) + 1
            Else
                kdSPP = 1
            End If
            text(8) = Format(kdSPP, "#00000000")
        End If
    
        Call saveLogMenu(gSA, ComptName, gVersi, "frmtSupplier", "cmdSimpan", 0)
        erTranAll = False
        onTR = True
        Call objHT.OpenTransactionBengkel
        
        Set rsSimpan = objO.FetchData("SELECT * FROM T_SUPPLIER WITH (NOLOCK) WHERE KDSUPP = '" & Trim(text(8)) & "'")
        If Not rsSimpan.EOF Then
            Call objHT.UpdateData_BengkelOnly("UPDATE T_SUPPLIER SET KODEAP = '" & Trim(text(0)) & "', SUPPLIER = '" & Trim(text(1)) & "', ALAMAT1 = '" & Trim(text(2)) & "', ALAMAT2 = '" & Trim(text(3)) & "',KOTA = '" & Trim(text(4)) & "', " & _
                           "TELPON = '" & Trim(text(5)) & "', CP = '" & Trim(text(6)) & "',PEKERJAAN = '" & Trim(txtPekerjaan) & "',PKP = '" & IIf(chPKP = Checked, 1, 0) & "',NONJURNAL = '" & IIf(chJurnal = Checked, 1, 0) & "' WHERE KDSUPP = '" & Trim(text(8)) & "'")
        Else
            Call objHT.InsertData_BengkelOnly("INSERT INTO T_SUPPLIER VALUES ( '" & Trim(text(8)) & "','" & Trim(text(1)) & "','" & Trim(text(2)) & "','" & Trim(text(3)) & "','" & Trim(text(4)) & "', '" & Trim(text(5)) & "','" & Trim(text(6)) & "','" & Trim(text(0)) & "','" & Trim(txtPekerjaan) & "','" & IIf(chPKP = Checked, 1, 0) & "','" & IIf(chJurnal = Checked, 1, 0) & "','','') ")
        End If

        If erTranAll = False Then
            Call objHT.CloseTransactionBengkel
            onTR = False
            Call saveLogMenu(gSA, ComptName, gVersi, "frmtSupplier", "cmdSimpan", 1)
            MsgBox "Supplier berhasil disimpan", vbInformation + vbOKOnly
            Call bersih
            Call refreshGrid
        Else
            Call objHT.CloseTransactionBengkel_err
            onTR = False
            Call saveLogMenu(gSA, ComptName, gVersi, "frmtSupplier", "cmdSimpan", 2)
            MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
        End If
            
    Else
        MsgBox pesan, vbCritical
    End If
    Exit Sub
    
vbError:
    Call objHT.CloseTransactionBengkel_err
    onTR = False
    Call saveLogMenu(gSA, ComptName, gVersi, "frmtSupplier", "cmdSimpan", 2)
    MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
    Exit Sub
End Sub

Private Sub Form_Load()
    Call bersih
    Call refreshGrid
End Sub

Private Sub refreshGrid()
    
    If Log.standarisasiSublet = 1 Then
        Set rsGrid = objO.FetchData("SELECT KODEAP,KDSUPP,SUPPLIER,ALAMAT1+' '+ALAMAT2 AS ALMT,KOTA, TELPON,CP,'' AS PEKERJAAN,PKP,NONJURNAL FROM T_SUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY KODEAP")
    Else
        Set rsGrid = objO.FetchData("SELECT KODEAP,KDSUPP,SUPPLIER,ALAMAT1+' '+ALAMAT2 AS ALMT,KOTA, TELPON,CP,PEKERJAAN,PKP,NONJURNAL FROM T_SUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY KODEAP")
    End If
    
    If Not rsGrid.EOF Then
        Set GridSupplier.DataSource = rsGrid
        Call headGrid
    End If
End Sub

Private Sub headGrid()
    With GridSupplier
        .Columns(0).Width = 1200
        .Columns(0).Caption = "Kode"
        .Columns(0).DataField = "KODEAP"
        .Columns(0).Locked = True
        .Columns(1).Width = 1500
        .Columns(1).Caption = "Kode Supplier"
        .Columns(1).DataField = "KDSUPP"
        .Columns(1).Locked = True
        .Columns(2).Width = 2000
        .Columns(2).Caption = "Supplier"
        .Columns(2).DataField = "SUPPLIER"
        .Columns(2).Locked = True
        .Columns(3).Width = 4000
        .Columns(3).Caption = "Alamat"
        .Columns(3).DataField = "ALMT"
        .Columns(3).Locked = True
        .Columns(4).Width = 1500
        .Columns(4).Caption = "Kota"
        .Columns(4).DataField = "KOTA"
        .Columns(4).Locked = True
        .Columns(5).Width = 2000
        .Columns(5).Caption = "Telpon"
        .Columns(5).DataField = "TELPON"
        .Columns(5).Locked = True
        .Columns(6).Width = 2000
        .Columns(6).Caption = "Contact Person"
        .Columns(6).DataField = "CP"
        .Columns(6).Locked = True
        .Columns(7).Width = 4000
        .Columns(7).Caption = "Sublet"
        .Columns(7).DataField = "PEKERJAAN"
        .Columns(7).Locked = True
        .Columns(8).Width = 800
        .Columns(8).Caption = "PKP"
        .Columns(8).DataField = "PKP"
        .Columns(8).Locked = True
        .Columns(9).Width = 800
        .Columns(9).Caption = "Non Jurnal"
        .Columns(9).DataField = "NONJURNAL"
        .Columns(9).Locked = True
    End With
End Sub

Private Sub bersih()
    For i = 0 To 8
        Select Case i
            Case Else
                text(i) = ""
        End Select
    Next
End Sub


Private Sub gridSupplier_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsSearch As ADODB.Recordset
    Set rsSearch = objO.FetchData("SELECT * FROM T_SUPPLIER WITH (NOLOCK) WHERE KDSUPP = '" & Trim(GridSupplier.Columns(1)) & "'")
    If Not rsSearch.EOF Then
        text(0) = Trim(rsSearch("KODEAP"))
        text(1) = Trim(rsSearch("SUPPLIER"))
        text(2) = Trim(rsSearch("ALAMAT1"))
        text(3) = Trim(rsSearch("ALAMAT2"))
        text(4) = Trim(rsSearch("KOTA"))
        text(5) = Trim(rsSearch("TELPON"))
        text(6) = Trim(rsSearch("CP"))
        txtPekerjaan = Trim(rsSearch("PEKERJAAN"))
        text(8) = Trim(rsSearch("KDSUPP"))
        chPKP = IIf(rsSearch("PKP") = 1, Checked, Unchecked)
        chJurnal = IIf(rsSearch("NONJURNAL") = 1, Checked, Unchecked)
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
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
                gMainKeyCari = text(Index)
                gSearchOption = "KODEPLGAP"
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

    Select Case Index
        Case 0
            If text(Index) <> "" Then
                text(7) = Left(text(Index), 3)
                Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp from Apmst WITH (NOLOCK) where KodeLgn = '" & Trim(text(Index)) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "KODEPLGAP"
                    frmCariAcess.show (vbModal)
                Else
                    Set rsHelp = objO.FetchData("SELECT * FROM T_SUPPLIER WITH (NOLOCK) WHERE KODEAP = '" & Trim(text(Index)) & "' AND FS_LOG_DEL = ''")
                    If Not rsHelp.EOF Then
                        text(8) = Trim(rsHelp("KDSUPP"))
                        text(0) = Trim(rsHelp("KODEAP"))
                        text(1) = Trim(rsHelp("SUPPLIER"))
                        text(2) = Trim(rsHelp("ALAMAT1"))
                        text(3) = Trim(rsHelp("ALAMAT2"))
                        text(4) = Trim(rsHelp("KOTA"))
                        text(5) = Trim(rsHelp("TELPON"))
                        text(6) = Trim(rsHelp("CP"))
                        txtPekerjaan = Trim(rsHelp("PEKERJAAN"))
                        chPKP = IIf(Trim(rsHelp("PKP")) = 1, Checked, Unchecked)
                        chJurnal = IIf(Trim(rsHelp("NONJURNAL")) = 1, Checked, Unchecked)
                    Else
                        text(0) = Trim(rsSearch("KODELGN"))
                        text(1) = Trim(rsSearch("NAMALGN"))
                        text(2) = Trim(rsSearch("ALAMAT1"))
                        text(3) = Trim(rsSearch("ALAMAT2"))
                        text(4) = Trim(rsSearch("ALAMAT3"))
                        text(5) = Trim(rsSearch("TELEPON"))
                        text(6) = ""
                        text(8) = ""
                        txtPekerjaan = ""
                        chPKP = IIf(Trim(rsSearch("NPWP")) <> "", Checked, Unchecked)
                        chJurnal = Unchecked
                    End If
                End If
            Else
                Call bersih
            End If
    End Select
End Sub
