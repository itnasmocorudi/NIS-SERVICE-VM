VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmOSubletSentralValid 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pemakaian Sublet"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   11715
   Begin VB.CommandButton cmdSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali Ke Menu"
      Height          =   495
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8130
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   11085
      Begin TrueOleDBGrid70.TDBGrid gridSublet 
         Height          =   4815
         Left            =   0
         TabIndex        =   39
         Top             =   3120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8493
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No. PO"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No Nota"
         Columns(1).DataField=   "NoNota"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Supplier"
         Columns(2).DataField=   ""
         Columns(2).DropDown=   "TDBDropDown1"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nama Sublet"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Satuan"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Qty"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Harga Satuan"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Subtotal"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(24)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=8196"
         Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(29)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(33)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(7)._ColStyle=8196"
         Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=212,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   930
         TabIndex        =   2
         Top             =   750
         Width           =   4815
         _ExtentX        =   8493
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   930
         TabIndex        =   3
         Top             =   1035
         Width           =   4815
         _ExtentX        =   8493
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   930
         TabIndex        =   4
         Top             =   1350
         Width           =   4815
         _ExtentX        =   8493
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   930
         TabIndex        =   5
         Top             =   1650
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   9
         Left            =   6750
         TabIndex        =   6
         Top             =   450
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   10
         Left            =   6750
         TabIndex        =   7
         Top             =   750
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   13
         Left            =   6750
         TabIndex        =   8
         Top             =   1350
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   12
         Left            =   6750
         TabIndex        =   9
         Top             =   1035
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   11
         Left            =   10170
         TabIndex        =   10
         Top             =   750
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   450
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   1
         Left            =   930
         TabIndex        =   12
         Top             =   450
         Width           =   1845
         _ExtentX        =   3254
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   14
         Left            =   6750
         TabIndex        =   13
         Top             =   1650
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   15
         Left            =   6750
         TabIndex        =   14
         Top             =   1935
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   17
         Left            =   9750
         TabIndex        =   15
         Top             =   435
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   7
         Left            =   930
         TabIndex        =   16
         Top             =   1935
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   930
         TabIndex        =   17
         Top             =   2235
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
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   2805
         TabIndex        =   18
         Top             =   450
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   16
         Left            =   6750
         TabIndex        =   19
         Top             =   2235
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   8760
         Top             =   6720
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
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
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   10920
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   10920
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3360
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3360
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemakaian Sublet"
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
         Left            =   4440
         TabIndex        =   43
         Top             =   2685
         Width           =   2265
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   10920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   10920
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3360
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3360
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
         Left            =   3720
         TabIndex        =   42
         Top             =   45
         Width           =   3675
      End
      Begin VB.Label SgnRekom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Rekom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   8280
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Mesin"
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
         Left            =   5955
         TabIndex        =   33
         Top             =   2280
         Width           =   645
      End
      Begin VB.Label Label9 
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
         TabIndex        =   32
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label8 
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
         TabIndex        =   31
         Top             =   1980
         Width           =   315
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   8640
         TabIndex        =   30
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka"
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
         Left            =   5820
         TabIndex        =   29
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label2 
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
         Left            =   5955
         TabIndex        =   28
         Top             =   1695
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
         TabIndex        =   27
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label19 
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
         Left            =   6120
         TabIndex        =   26
         Top             =   1365
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
         Left            =   6180
         TabIndex        =   25
         Top             =   765
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
         Left            =   5955
         TabIndex        =   24
         Top             =   465
         Width           =   645
      End
      Begin VB.Label Label5 
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
         TabIndex        =   23
         Top             =   1695
         Width           =   375
      End
      Begin VB.Label Label4 
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
         TabIndex        =   22
         Top             =   1395
         Width           =   330
      End
      Begin VB.Label Label3 
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
         TabIndex        =   21
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label6 
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
         TabIndex        =   20
         Top             =   795
         Width           =   405
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11055
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   2640
         Width           =   11055
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2805
      TabIndex        =   36
      Top             =   120
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   0
      Left            =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
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
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   0
      Picture         =   "frmOSubletSentralValid.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total :"
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
      Left            =   480
      TabIndex        =   41
      Top             =   8520
      Width           =   1875
   End
   Begin VB.Label grandTotal 
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
      Left            =   2520
      TabIndex        =   40
      Top             =   8520
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   615
      TabIndex        =   37
      Top             =   165
      Width           =   570
   End
End
Attribute VB_Name = "frmOSubletSentralValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSublet As New ADODB.Recordset
Dim rsSearch As ADODB.Recordset
Dim rsHelp As ADODB.Recordset
Dim vPDPGL As String
Dim vHutangGL As String
Dim vPpnGL As String

Private Function createSublet() As Boolean
    On Error GoTo vb_error
    
    If rsSublet.State Then rsSublet.Close
    Set rsSublet = Nothing
    With rsSublet
        .Fields.Append "NOPO", adVarChar, 25
        .Fields.Append "NONOTA", adVarChar, 25
        .Fields.Append "SUPP", adVarChar, 15
        .Fields.Append "SUBLET", adVarChar, 50 '40
        .Fields.Append "SAT", adVarChar, 10
        .Fields.Append "QTY", adDouble
        .Fields.Append "HARGASAT", adDouble
        .Fields.Append "SUBTOTAL", adDouble
        .Open
    End With
    If rsSublet.State Then createSublet = True
    Exit Function
    
vb_error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub headSublet()
    
    With gridSublet
        .Columns(0).Caption = "No. PO"
        .Columns(0).DataField = "NOPO"
        .Columns(0).Width = 1400
        
        .Columns(1).Caption = "No Nota"
        .Columns(1).DataField = "nonota"
        .Columns(1).Width = 900
        
        .Columns(2).Caption = "Supplier"
        .Columns(2).DataField = "SUPP"
        .Columns(2).Width = 1200
        
        .Columns(3).Caption = "Pekerjaan Sublet"
        .Columns(3).DataField = "SUBLET"
        .Columns(3).Width = 2600
        
        .Columns(4).Caption = "Satuan"
        .Columns(4).DataField = "SAT"
        .Columns(4).Width = 900
        
        .Columns(5).Caption = "Qty"
        .Columns(5).DataField = "QTY"
        .Columns(5).Width = 600
        
        .Columns(6).Caption = "Harga Satuan"
        .Columns(6).DataField = "HARGASAT"
        .Columns(6).Width = 1200
        
        .Columns(7).Caption = "Sub Total"
        .Columns(7).DataField = "SUBTOTAL"
        .Columns(7).Width = 1200
    End With
End Sub

Private Sub hapusForm()
    For i = 1 To 17
        Select Case i
            Case 9
                text(i) = 0
            Case 17
                text(i) = "99-99-9999"
            Case Else
                text(i) = ""
        End Select
    Next
    If rsSublet.State Then rsSublet.Close
    grandTotal = 0

'    With Adodc1
'        .ConnectionString = SQLConACC
'        .CommandType = adCmdText
'        .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST WITH(NOLOCK)) ORDER BY KODEAP"
'        .Refresh
'    End With
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
      Dim pesanIT As String
      Dim noJurnal As Boolean
      'Dim rsAPtrn     As ADODB.Recordset

1         On Error GoTo vbError
2         boleh = True
3         pesan = ""
4         pesanIT = ""
          Dim sum_bengkel As Double
          
          
5         If rsSublet.State <> adStateClosed Then
6             rsSublet.MoveFirst
7             Do While Not rsSublet.EOF
8                 If rsSublet("SUPP") = "" Then
9                     boleh = False
10                    pesan = pesan + "Kode Supplier untuk " & rsSublet("SUBLET") & ". " & vbCr
11                Else
12                    Set rsHelp = objO.FetchDataACC("SELECT * FROM APMST WITH(NOLOCK) WHERE KODELGN = '" & NulSID(rsSublet("SUPP")) & "'")
13                    If rsHelp.EOF Then
14                        boleh = False
15                        pesan = pesan + "Kode Supplier " & NulSID(rsSublet("SUPP")) & " untuk " & rsSublet("SUBLET") & ". " & vbCr
16                    End If
17                End If
18                If rsSublet("HARGASAT") = "0" Then
19                    If MsgBox("Yakin harga beli " & rsSublet("SUBLET") & " 0 (Nol) ?", vbYesNo, "Validasi Harga") = vbNo Then
20                        boleh = False
21                        pesan = pesan + "Harga Beli Satuan " & rsSublet("SUBLET") & ". " & vbCr
22                    End If
23                End If
24                If rsSublet("nonota") = "" Then
25                    boleh = False
26                    pesan = pesan + "No Nota untuk " & rsSublet("SUBLET") & ". " & vbCr
27                End If
                  
                  ''2016-07-22 || cek total bengkel dengan aptrn
28                sum_bengkel = sum_bengkel + (Abs(CDbl(rsSublet("HARGASAT"))) * CDbl(rsSublet("QTY")))
29                rsSublet.MoveNext
30            Loop
31        End If
          Dim rsTemp As ADODB.Recordset
          
          
32        rsSublet.MoveFirst
33        If CDbl(Right(NisDb, 2)) <> CDbl(Month(isnow)) Then
34            boleh = False
35            pesan = pesan + "Beda Fiskal Accounting tidak dapat dibuat invoice. " & vbCr & "Pindah Fiskal Accounting pada bulan bersangkutan. "
36        End If
          
          ''2015-11-25
          ''tambah validasi tempjobsublet untuk wo tsb harus kosong
          ''Masih ada user yang membuka menu order dengan no order = ''
37        Set rsHelp = objO.FetchData("SELECT * FROM TEMPJOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "'")
38        If Not rsHelp.EOF Then
39            rsHelp.MoveFirst
40            boleh = False
41            pesan = pesan + "Masih ada user yang membuka menu order dengan no order = " & rsHelp!no_wo
42        End If
43        If boleh Then
      '44            Set rsTemp = objO.FetchData("select queryTransferAptrnSublet from t_profil with (nolock)")
      '45            If Not rsTemp.EOF Then
      '46                sSQL = NulSID(rsTemp!queryTransferAptrnSublet)
      '47                sSQL = Replace(sSQL, "~", "'")
      '48                sSQL = Replace(sSQL, "#KODEDEALER#", "'" & Right(gDb, 2) & "'")
      '49                sSQL = Replace(sSQL, "#LOGACC#", Log.LogACC)
      '50                Set rsAPtrn = objO.FetchData(sSQL)
      '51            End If
44             Set rsTemp = objO.FetchDataACC("select isnull(SUM(jumlahtrn),0) total from aptrn where nobukti like '" & "S-" & Trim(text(0)) & "%'")
45             If Not rsTemp.EOF Then
46               rsTemp.MoveFirst
                 
                 ''If rsTemp!total <> sum_bengkel And rsTemp!total <> 0 Then
                 
47               If Not (rsTemp!total <> sum_bengkel And rsTemp!total = 0) Then
                    ''insert data ke selisih sublet
48                  Call objO.InsertData("INSERT INTO [SelisihSublet] ([no_wo] ,[no_chassis] ,[sum_bengkel] ,[sum_ap] ,[tglEntry],KodeUser,menu) VALUES ('" & Trim(text(0)) & "','" & Trim(text(15)) & "'," & sum_bengkel & "," & CDbl(rsTemp!total) & ",GETDATE(),'" & gSA & "','Validasi Sublet While Insert')")
49               End If
50             End If
              
              
51            Set rsTemp = objO.FetchDataACC("select * from aptrn where nofaktur = '" & "S-" & Trim(text(0)) & "'")
52            If Not rsTemp.EOF Then
53               rsTemp.MoveFirst
54               If objO.FetchData_DoubleACC("select isnull(SUM(jumlahtrn),0) from aptrn where nofaktur = '" & "S-" & Trim(text(0)) & "'") = 0 Then
55                    If objO.FetchData_Count("select * from [" & Log.LogACC & "]..aptrn where nobukti not like 'S-%' and nofaktur = '" & "S-" & Trim(text(0)) & "'") > 0 Then
56                        MsgBox "Sudah ada tagihan untuk no wo ini di akunting." & vbCrLf & "Proses tidak dapat dilanjutkan", vbInformation
57                        Exit Sub
58                    Else
                      
59                    End If
60               Else
61                    MsgBox "Sudah ada tagihan untuk no wo ini di akunting." & vbCrLf & "Proses tidak dapat dilanjutkan", vbInformation
62                    Exit Sub
63               End If
64            End If
              Dim jobSub(100) As String
              Dim iSub As Integer
65            iSub = 0
66            Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletSentralValid", "cmdSimpan", 0)
67            erTranAll = False
68            onTR = True
69            Call objHT.OpenTransactionAll


              ' SIMPAN SUBLET
              
70            If rsSublet.State Then
71                If Not rsSublet.EOF Then
                      Dim isDyna As Boolean
                      'PENAMBAHAN PENGECEKAN PENYERAGAMAN AKUN
                      'TANGGAL 04-09-2014
                      'RUDI KRISTANTO
                      
72                    If Log.isAkun = 1 Then
73                        isDyna = IIf(InStr(1, UCase(text(12).text), "DYNA") > 0, True, False)
74                        Set rsGL = objO.Get_AkunGL_NEW_Sublet("OTHERS", isDyna, False)
75                        If isDyna = True Then
76                            vPDPGL = cekKodeAkun(Trim(rsGL.Fields(0).value))
77                            vPpnGL = cekKodeAkun(Trim(rsGL.Fields(1).value))
78                            vHutangGL = cekKodeAkun(Trim(rsGL.Fields(2).value))
79                        Else
80                            vPDPGL = cekKodeAkun(Trim(rsGL.Fields(0).value))
81                            vPpnGL = cekKodeAkun(Trim(rsGL.Fields(1).value))
82                            vHutangGL = cekKodeAkun(Trim(rsGL.Fields(2).value))
83                        End If
84                    Else
85                        Set rsGL = objHT.FetchDataBengkel("SELECT FS_KD_PDP, FS_KD_PPN, FS_KD_HUTANG FROM T_AKUN WITH(NOLOCK)")
86                        If Not rsGL.EOF Then
87                            vPDPGL = cekKodeAkun(Trim(rsGL("FS_KD_PDP")))
88                            vPpnGL = cekKodeAkun(Trim(rsGL("FS_KD_PPN")))
89                            vHutangGL = cekKodeAkun(Trim(rsGL("FS_KD_HUTANG")))
90                        End If
91                    End If
                      
                      
92                    Do While Not rsSublet.EOF
                          Dim sudahValid As Boolean
93                        sudahValid = False
94                        Set rsHelp = objHT.FetchDataBengkel("SELECT Valid FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND NO_CHASSIS = '" & Trim(text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'")
95                        If Not rsHelp.EOF Then
96                            If rsHelp("VALID") = 1 Then sudahValid = True
97                        End If
                          
98                        If Not sudahValid Then
                              ' INPUT PDP TERHADAP HUTANG SUPPLIER
                                  Dim vDelIndex As String
                                  Dim vKodeLgn As String
                                  Dim vTglTempo As Date
                                  Dim vTglSkr As Date
                                  Dim vTBeli As Double
                                  Dim vTPPN As Double
                                  Dim vTotal As Double
                                  Dim vKeterangan As String
                                  
99                                vDelIndex = gKode & Left(McDelIndex(7), 8)
100                               vKodeLgn = Trim(rsSublet("SUPP"))
101                               vTglSkr = Format(isnow, "dd/mm/yyyy hh:nn")
102                               vTglTempo = vTglSkr
103                               vTotal = SpyRound(CDbl(rsSublet("HARGASAT"))) * CDbl(rsSublet("QTY")) ''Abs(CDbl(rsSublet("HARGASAT"))) * CDbl(rsSublet("QTY"))
104                               vTPPN = SpyRound(vTotal * (profilACC.nilaippn / profilACC.nilaitotal)) ''(10 / 110))
105                               vTBeli = vTotal - vTPPN
                                  
106                               If vPDPGL = "" Or vPpnGL = "" Or vHutangGL = "" Then
107                                   pesanIT = pesanIT + "Kode GL ada yang kosong !"
108                                   GoTo vbError
109                               End If
                                  
110                               noJurnal = False
111                               Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM T_SUPPLIER WITH(NOLOCK) WHERE KODEAP = '" & Trim(rsSublet("SUPP")) & "'")
112                               If Not rsSearch.EOF Then
113                                   vKeterangan = rsSearch("SUPPLIER") & " (" & rsSublet("SUBLET") & ") No. WO :" & Trim(text(0))
114                                   If Trim(rsSearch("PKP")) = "0" Then
                                          'vPpnGL = ""
115                                       vTBeli = vTBeli + vTPPN
116                                       vTPPN = 0
117                                       vTotal = vTBeli
118                                   End If
119                                   If Trim(rsSearch("NONJURNAL")) = "1" Then
120                                       noJurnal = True
121                                   End If
122                               End If
                                  
                                  ' update status sublet jadi valid
                                  'Call objHT.UpdateData_Bengkel("UPDATE JOBSUBLET SET HARGA = '" & Round(CDbl(vTotal) / CDbl(rsSublet("QTY"))) & "', SUPPLIER = '" & Trim(rsSublet("SUPP")) & "', VALID = 1, TGLVALID = GETDATE() WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'")
                                  'jobSub(iSub) = "UPDATE JOBSUBLET SET HARGA = '" & Round(CDbl(vTotal) / CDbl(rsSublet("QTY"))) & "', SUPPLIER = '" & Trim(rsSublet("SUPP")) & "', VALID = 1, TGLVALID = '" & Format(vTglSkr, "MM/DD/YYYY HH:NN") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'"
123                               jobSub(iSub) = "UPDATE JOBSUBLET SET NONOTA = '" & NulSID(rsSublet!nonota) & "',HARGA = '" & Round(CDbl(vTotal) / CDbl(rsSublet("QTY"))) & "', SUPPLIER = '" & Trim(rsSublet("SUPP")) & "', VALID = 1, TGLVALID = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "' WHERE NO_WO = '" & Trim(text(0)) & "' AND NO_CHASSIS = '" & Trim(text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'"
124                               iSub = iSub + 1
                                  
125                               If gConnecAAS = True Then
126                                   If Not noJurnal Then
127                                       If Abs(CDbl(rsSublet("HARGASAT"))) <> 0 Then
                                              ' INPUT DI APTRN
                                              ' PDP
128                                           sSQL1 = "insert into APtrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                                                        & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, tglEntry,kodejurnal) " _
                                                        & "values " _
                                                        & " ('" & Trim(vKodeLgn) & "', '" & "S-" & Trim(text(0)) & "', 'C'," _
                                                        & " '" & "S-" & Trim(text(0)) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', 'PDP Sublet " & rsSublet("SUBLET") & " No. WO : " & Trim(text(0)) & "'," _
                                                        & " " & SimpanKoma2Titik(vTBeli) & ", '" & gSA & "'," _
                                                        & " '" & vDelIndex & "','" & gKode & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vHutangGL & "')"
129                                           Call objHT.InsertData_ACC(sSQL1)
                                              
                                              ' PPN
130                                           sSQL1 = "insert into APtrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                                                        & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, tglEntry,kodejurnal) " _
                                                        & "values " _
                                                        & " ('" & Trim(vKodeLgn) & "', '" & "S-" & Trim(text(0)) & "', 'F'," _
                                                        & " '" & "S-" & Trim(text(0)) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', 'PPN Sublet " & rsSublet("SUBLET") & " No. WO : " & Trim(text(0)) & "'," _
                                                        & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "'," _
                                                        & " '" & vDelIndex & "','" & gKode & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vHutangGL & "')"
131                                           Call objHT.InsertData_ACC(sSQL1)
                                              
                                              ' INPUT DI APMST
132                                           sSQL1 = "Update APMST set SALDOJALAN = SALDOJALAN + " & SimpanKoma2Titik(vTBeli) & ", TGLBELIAHR = '" & Format(isnow, "MM/DD/YYYY") & "' where KodeLgn= '" & Trim(vKodeLgn) & "'"
133                                           Call objHT.UpdateData_ACC(sSQL1)
                                              
134                                           sSQL1 = "Update APMST set SALDOJALAN = SALDOJALAN + " & SimpanKoma2Titik(vTPPN) & ", TGLBELIAHR = '" & Format(isnow, "MM/DD/YYYY") & "' where KodeLgn= '" & Trim(vKodeLgn) & "'"
135                                           Call objHT.UpdateData_ACC(sSQL1)
                      
                                              ' INPUT DI GLTRN
136                                           If vTotal > 0 Then
137                                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(text(0)) & "', '" & vPDPGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTBeli) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
138                                               Call objHT.InsertData_ACC(sSQL1)
                                                  
139                                               If vTPPN > 0 Then
140                                                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(text(0)) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
141                                                   Call objHT.InsertData_ACC(sSQL1)
142                                               End If
                                                  
143                                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(text(0)) & "', '" & vHutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTotal) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
144                                               Call objHT.InsertData_ACC(sSQL1)
                                                  
                                                  ' INPUT DI GLMST
145                                               sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTBeli) & ", " _
                                                        & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(vTBeli) & " " _
                                                        & " where KodeGl = '" & vPDPGL & "'"
146                                               Call objHT.UpdateData_ACC(sSQL1)
                                                  
147                                               If vTPPN > 0 Then
148                                                   sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTPPN) & ", " _
                                                        & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(vTPPN) & " " _
                                                        & " where KodeGl = '" & vPpnGL & "'"
149                                                   Call objHT.UpdateData_ACC(sSQL1)
150                                               End If
                                                  
151                                               sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTotal) & ", " _
                                                        & " JKreditBi = JKreditBi + " & SimpanKoma2Titik(vTotal) & " " _
                                                        & " where KodeGl = '" & vHutangGL & "'"
152                                               Call objHT.UpdateData_ACC(sSQL1)
153                                           End If
154                                       End If
155                                   End If
156                               End If
157                           End If
158                       rsSublet.MoveNext
159                   Loop
                      Dim xI As Integer
160                   For xI = 0 To 100
161                       If jobSub(xI) <> "" Then
162                           Call objHT.UpdateData_Bengkel(jobSub(xI))
163                       Else
164                           Exit For
165                       End If
166                   Next xI
167               End If
                  
                  ''20190612 penambahan transfer data aptrn_transfer di replbengkel99
                  '' -----------------------------------------------------------------
      '            Set rsTemp = objHT.FetchDataBengkel("select queryTransferAptrnSublet from t_profil with (nolock)")
      '            If Not rsTemp.EOF Then
      '                sSQL = NulSID(rsTemp!queryTransferAptrnSublet)
      '                sSQL = Replace(sSQL, "~", "'")
      '                sSQL = Replace(sSQL, "#KODEDEALER#", "'" & Right(gDb, 2) & "'")
      '                sSQL = Replace(sSQL, "#LOGACC#", Log.LogACC)
      '                'Set rsTemp = objHT.FetchDataBengkel(sSQL)
      '                Call objHT.InsertData_Bengkel(sSQL)
      '            End If
      '162               If rsAPtrn.State = adStateOpen Then
      '163                   If Not rsAPtrn.EOF Then
      '164                       rsAPtrn.MoveFirst
      '165                       Do While Not rsAPtrn.EOF
      '166                           sSQL = "INSERT INTO [ReplBengkel99]..[APTRN_TRANSFER]([KodeDealer],[KodeSup],[KodeLgn],[NoFaktur],[TypeTrn],[NoBukti],[TglTrn],[TglJthTmp],[TglTrnFaktur],[TglJtpFaktur],[Keterangan],[JumlahTrn],[TglEntry],[Kodeuser],[Kodesumber],[KodeJurnal],[DelIndex]) VALUES ( "
      '167                           sSQL = sSQL & " '" & Right(gDb, 2) & "','" & NulSID(rsAPtrn!KDSUPP) & "','" & NulSID(rsAPtrn!kodeLgn) & "','" & NulSID(rsAPtrn!NoFaktur) & "','" & NulSID(rsAPtrn!TYPETRN) & "','" & NulSID(rsAPtrn!Nobukti) & "','" & Format(NulSID(rsAPtrn!TGLTRN), "YYYY-MM-DD") & "','" & Format(NulSID(rsAPtrn!TGLJTHTMP), "YYYY-MM-DD") & "','" & Format(NulSID(rsAPtrn!TGLTRNFAKTUR), "YYYY-MM-DD") & "','" & Format(NulSID(rsAPtrn!TGLJTPFAKTUR), "YYYY-MM-DD") & "','" & NulSID(rsAPtrn!keterangan) & "'"
      '168                           sSQL = sSQL & " ,'" & NulSID(rsAPtrn!JUMLAHTRN) & "','" & Format(NulSID(rsAPtrn!tglentry), "YYYY-MM-DD HH:MM:SS") & "','" & NulSID(rsAPtrn!KODEUSER) & "','" & NulSID(rsAPtrn!KODESUMBER) & "','" & NulSID(rsAPtrn!KODEJURNAL) & "','" & NulSID(rsAPtrn!DELINDEX) & "')"
      '                              Call objHT.InsertData_Bengkel(sSQL)
      '169                           rsAPtrn.MoveNext
      '170                       Loop
      '171                   End If
      '172               End If
168               If erTranAll = False Then
169                   Call objHT.CloseTransactionAll
                      ''2016-07-22
170                   Set rsTemp = objO.FetchDataACC("select isnull(SUM(jumlahtrn),0) total from aptrn where nobukti like '" & "S-" & Trim(text(0)) & "%'")
171                     If Not rsTemp.EOF Then
172                       rsTemp.MoveFirst
                          ''If rsTemp!total <> sum_bengkel And rsTemp!total <> 0 Then
173                       If Not (rsTemp!total = sum_bengkel And rsTemp!total <> 0) Then
                             ''insert data ke selisih sublet
174                          Call objO.InsertData("INSERT INTO [SelisihSublet] ([no_wo] ,[no_chassis] ,[sum_bengkel] ,[sum_ap] ,[tglEntry],KodeUser,menu) VALUES ('" & Trim(text(0)) & "','" & Trim(text(15)) & "'," & sum_bengkel & "," & CDbl(rsTemp!total) & ",GETDATE(),'" & gSA & "','Validasi Sublet After Insert')")
175                       End If
176                     End If
177                   onTR = False
178                   Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletSentralValid", "cmdSimpan", 1)
179                   MsgBox "Data berhasil disimpan "
180                   cmdSimpan.Enabled = False
181               Else
182                   Call objHT.CloseTransactionAll_Err
183                   onTR = False
184                   Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletSentralValid", "cmdSimpan", 2)
185                   MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
186               End If

187           End If
188       Else
189           MsgBox "Data belum lengkap !!" & vbCr & pesan
190       End If
191       Exit Sub

vbError:
192       Call objHT.CloseTransactionAll_Err
193       onTR = False
194       Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletSentralValid", "cmdSimpan", 2)
195       MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
196       Exit Sub
End Sub

Private Sub Form_Activate()
    
    If orderAktif Or validAktif Or subRetAktif Then
        MsgBox "Tutup Menu sebelumnya terlebih Dahulu !!", vbCritical
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    subletAktif = True
    grandTotal = "0"
    Call hapusForm
End Sub

Private Sub Form_Resize()
    Image1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subletAktif = False
End Sub

Private Sub gridSublet_AfterColUpdate(ByVal ColIndex As Integer)
Dim rsHlp As New ADODB.Recordset

    Select Case ColIndex
        Case 5
            With gridSublet
                .Columns(5) = Abs(CDbl(.Columns(5)))
                rsSublet("HARGASAT") = Abs(CDbl(.Columns(5)))
                .Columns(6) = CDbl(.Columns(4)) * Abs(CDbl(.Columns(5)))
                rsSublet.MoveFirst
                grandTotal = 0
                Do While Not rsSublet.EOF
                    grandTotal = Format(grandTotal + CDbl(rsSublet("SUBTOTAL")), "##,0")
                    rsSublet.MoveNext
                Loop
            End With
        Case 1
            If Len(NulSID(rsSublet!nonota)) > 10 Then
                rsSublet!nonota = Mid(NulSID(rsSublet!nonota), 1, 10)
            End If
            
'            With gridSublet
'                Set rsHelp = objO.FetchData("SELECT * FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE KODEAP = '" & Trim(.Columns(1)) & "' AND FS_LOG_DEL = ''")
'                If rsHelp.EOF Then
'                    rsSublet.Find "SUBLET = '" & Trim(.Columns(2)) & "'", , adSearchForward, 1
'                    If Not rsSublet.EOF Then
'                        rsSublet("SUPP") = ""
'                    End If
'                Else
'                    Set rsHlp = objO.FetchDataACC("SELECT * FROM APMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(.Columns(1)) & "'")
'                    If Not rsHlp.EOF Then
'                        rsSublet("SUPP") = UCase(Trim(.Columns(1)))
'                    Else
'                        MsgBox "Kode Supplier tidak terdaftar di Accounting !"
'                        rsSublet.Find "SUBLET = '" & Trim(.Columns(2)) & "'", , adSearchForward, 1
'                        If Not rsSublet.EOF Then
'                            rsSublet("SUPP") = ""
'                        End If
'                    End If
'                End If
'            End With
    End Select
    Set gridSublet.DataSource = rsSublet
End Sub

Private Sub GridSublet_ButtonClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 1
            With Adodc1
                .ConnectionString = SQLConACC
                .CommandType = adCmdText
                .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST WITH(NOLOCK)) ORDER BY KODEAP"
                .Refresh
            End With
    End Select
End Sub

Private Sub GridSublet_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 18
            With Adodc1
                .ConnectionString = SQLConACC
                .CommandType = adCmdText
                .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST WITH(NOLOCK)) ORDER BY KODEAP"
                .Refresh
            End With
    End Select
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
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
                gSearchOption = "VALIDSUBLET"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If text(Index).BackColor = &HC0E0FF Then KeyAscii = 0
End Sub

Private Function cariValid(nWO As String, nChassis As String) As Integer
Dim rsCari As ADODB.Recordset

    Set rsCari = objO.FetchData("SELECT SUM(VALID) AS OK FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(nWO) & "' AND NO_CHASSIS = '" & Trim(nChassis) & "'")
    If Not rsCari.EOF Then
        cariValid = CDbl(rsCari("OK"))
    End If
End Function

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
1         Select Case Index
              Case 0
                  ' NO WO VALIDASI
2                 If text(Index) <> "" Then
                      ''Set rsSearch = objO.FetchData("SELECT B.FS_KD_PLG, B.FS_KD_WIL, B.CUSTOMER, B.CUST_ADRES, B.KOTA, B.TELPON, GSM, CDMA, B.FS_TH_MOBIL, B.MDL, " _
                                                 & "B.MODEL, B.FMDL, B.COLOR, B.NO_POLISI, B.NO_CHASSIS, C.NO_ENGINE, C.DDATE, NOPO, NAMABARANG, KODESATUAN, QTY, HARGA, SUPPLIER, NOINV, VALID FROM JOBSUBLET A WITH(NOLOCK) " _
                                                 & "INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS " _
                                                 & "INNER JOIN CUSTDAT C WITH(NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS WHERE B.NOINV = '' AND A.NO_WO = '" & Trim(text(Index)) & "'ORDER BY A.NO_WO")     '
3                     sSQL = "SELECT A.NONOTA,B.FS_KD_PLG, B.FS_KD_WIL, B.CUSTOMER, B.CUST_ADRES, B.KOTA, B.TELPON, GSM, CDMA, "
4                     sSQL = sSQL & " B.FS_TH_MOBIL, B.MDL, B.MODEL, B.FMDL, B.COLOR, B.NO_POLISI, B.NO_CHASSIS, C.NO_ENGINE,"
5                     sSQL = sSQL & " C.DDATE, NOPO, NAMABARANG, A.KODESATUAN, QTY, E.HARGABELI HARGA, D.KODEAP AS SUPPLIER, NOINV, VALID,D.PKP"
6                     sSQL = sSQL & " FROM JOBSUBLET A WITH(NOLOCK)"
7                     sSQL = sSQL & " INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS"
8                     sSQL = sSQL & " INNER JOIN CUSTDAT C WITH(NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS"
9                     sSQL = sSQL & " INNER JOIN [" & dbReplikasiHO & "]..MSTSUBLET E ON E.KDSUBLET = A.KODEITEM"
10                    sSQL = sSQL & " LEFT JOIN T_SUPPLIER D WITH (NOLOCK) ON E.KDSUPP = D.KDSUPP"
11                    sSQL = sSQL & " WHERE B.NOINV = '' AND A.NO_WO = '" & Trim(text(Index)) & "'"
12                    sSQL = sSQL & " ORDER BY A.NO_WO"
                      
13                    Set rsSearch = objO.FetchData(sSQL)
                      
14                    If Not rsSearch.EOF Then
15                        Call hapusForm
16                        If Trim(rsSearch("NOINV")) = "" Then
17                            If CDbl(rsSearch.RecordCount) > cariValid(Trim(text(Index)), Trim(rsSearch("NO_CHASSIS"))) Then
18                                text(1) = Trim(rsSearch("FS_KD_PLG"))
19                                text(2) = Trim(rsSearch("FS_KD_WIL"))
20                                text(3) = Trim(rsSearch("CUSTOMER"))
21                                text(4) = Left(Trim(rsSearch("CUST_ADRES")), 60)
22                                text(5) = Trim(rsSearch("KOTA"))
23                                text(6) = Trim(rsSearch("TELPON"))
24                                text(7) = Trim(rsSearch("GSM"))
25                                text(8) = Trim(rsSearch("CDMA"))
26                                text(9) = Trim(rsSearch("FS_TH_MOBIL"))
27                                text(10) = Trim(rsSearch("MDL"))
28                                text(11) = Trim(rsSearch("MODEL"))
29                                text(12) = Trim(rsSearch("FMDL"))
30                                text(13) = Trim(rsSearch("COLOR"))
31                                text(14) = Trim(rsSearch("NO_POLISI"))
32                                text(15) = Trim(rsSearch("NO_CHASSIS"))
33                                text(16) = Trim(rsSearch("NO_ENGINE"))
34                                text(17).Mask = ""
35                                text(17) = Nul(IIf(Replace(rsSearch("DDATE"), "/", "-") = "01-01-3000", "", IIf(Replace(rsSearch("DDATE"), "/", "-") = "01-01-1900", "", Replace(rsSearch("DDATE"), "/", "-"))))
                                  'Trim(rsSearch("DDATE"))
                                  
36                                Call createSublet
37                                Do While Not rsSearch.EOF
38                                    If rsSearch("VALID") = 0 Then
39                                        If rsSublet.State Then
40                                            With rsSublet
41                                                .AddNew
42                                                !NOPO = Trim(rsSearch("NOPO"))
43                                                !nonota = Trim(rsSearch("NONOTA"))
44                                                !SUPP = NulSID(rsSearch("SUPPLIER"))
45                                                !SUBLET = Trim(rsSearch("NAMABARANG"))
46                                                !SAT = Trim(rsSearch("KODESATUAN"))
47                                                !qty = CDbl(Trim(rsSearch("QTY")))
48                                                If NulSID(rsSearch!pkp) = "1" Then
49                                                    !HARGASAT = SpyRound(CDbl(Trim(rsSearch("HARGA"))) * profilACC.nilaitotal / profilACC.nilaidpp, 0)
50                                                    !Subtotal = SpyRound(CDbl(Trim(rsSearch("QTY"))) * (CDbl(Trim(rsSearch("HARGA"))) * profilACC.nilaitotal / profilACC.nilaidpp), 0)
51                                                Else
52                                                    !HARGASAT = SpyRound(CDbl(Trim(rsSearch("HARGA"))), 0)
53                                                    !Subtotal = SpyRound(CDbl(Trim(rsSearch("QTY"))) * CDbl(Trim(rsSearch("HARGA"))), 0)
54                                                End If
55                                                .Update
56                                            End With
57                                            grandTotal = Format(grandTotal + CDbl(Trim(rsSearch("QTY"))) * CDbl(Trim(rsSearch("HARGA"))), "##,0")
58                                        End If
59                                    End If
60                                    rsSearch.MoveNext
61                                Loop
62                                cmdSimpan.Enabled = True
63                                Set gridSublet.DataSource = rsSublet
64                                Call headSublet
65                            Else
66                                MsgBox "WO telah divalidasi "
67                            End If
68                        Else
69                            MsgBox "WO telah dibuat invoice. Retur Invoice"
70                        End If
71                    Else
72                        MsgBox "WO tersebut tidak ada !!"
73                    End If
74                End If
75        End Select
End Sub

